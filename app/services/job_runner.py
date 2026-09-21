# -*- coding: utf-8 -*-
"""Job Runner service — orchestrates the full comparative generation pipeline.

Executes as a FastAPI BackgroundTask:
1. Download/read input files
2. Parse text from each document
3. MAP: Extract structured data from each quote in parallel via Gemini
4. REDUCE: Consolidate into unified data structure
5. Generate AI recommendation and worst markers
6. Render Word document
7. Upload to Supabase Storage
8. Send webhook callback to BPMS/frontend
"""
from __future__ import annotations

import asyncio
import json
import logging
import tempfile
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any

import httpx

from app.core.config import settings
from app.core.exceptions import CallbackError
from app.services.consolidator import consolidate_quotes
from app.services.docx_renderer import render_comparative
from app.services.document_parser import download_file, parse_document
from app.services.llm_extractor import (
    extract_project_meta,
    extract_quote,
    generate_recommendation,
    generate_worst_markers,
)
from app.services.storage import upload_to_supabase

logger = logging.getLogger(__name__)


# --------------------------------------------------------------------------- #
# Job state management (in-memory for simplicity)                              #
# --------------------------------------------------------------------------- #

class JobStatus(str, Enum):
    QUEUED = "queued"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"


@dataclass
class JobState:
    """Tracks the state of a comparative generation job."""
    job_id: str
    process_id: str
    status: JobStatus = JobStatus.QUEUED
    download_url: str | None = None
    error: str | None = None
    created_at: datetime = field(default_factory=datetime.now)
    updated_at: datetime = field(default_factory=datetime.now)


# In-memory job store (replace with Redis/DB for production multi-instance)
_jobs: dict[str, JobState] = {}


def get_job(job_id: str) -> JobState | None:
    """Retrieve a job's current state."""
    return _jobs.get(job_id)


def create_job(process_id: str) -> JobState:
    """Create and register a new job."""
    job_id = str(uuid.uuid4())
    state = JobState(job_id=job_id, process_id=process_id)
    _jobs[job_id] = state
    return state


def _update_job(job_id: str, **kwargs: Any) -> None:
    """Update a job's state."""
    if job_id in _jobs:
        for k, v in kwargs.items():
            setattr(_jobs[job_id], k, v)
        _jobs[job_id].updated_at = datetime.now()


# --------------------------------------------------------------------------- #
# Callback                                                                     #
# --------------------------------------------------------------------------- #

async def _send_callback(
    callback_url: str,
    payload: dict,
    retries: int = 3,
) -> None:
    """Send a webhook callback to the BPMS/frontend.

    Args:
        callback_url: URL to POST the result to.
        payload: JSON payload with job_id, status, download_url, etc.
        retries: Number of retry attempts.
    """
    for attempt in range(1, retries + 1):
        try:
            async with httpx.AsyncClient(timeout=30.0) as client:
                response = await client.post(
                    callback_url,
                    json=payload,
                    headers={"Content-Type": "application/json"},
                )
                response.raise_for_status()
                logger.info(
                    "Callback enviado exitosamente a %s (intento %d)",
                    callback_url,
                    attempt,
                )
                return
        except Exception as e:
            logger.warning(
                "Error enviando callback (intento %d/%d): %s",
                attempt,
                retries,
                e,
            )
            if attempt < retries:
                await asyncio.sleep(2 ** attempt)  # Exponential backoff
            else:
                raise CallbackError(
                    message=f"No se pudo enviar el callback después de {retries} intentos",
                    detail=str(e),
                ) from e


# --------------------------------------------------------------------------- #
# Main pipeline                                                                #
# --------------------------------------------------------------------------- #

async def run_comparative_job(
    job_id: str,
    process_id: str,
    uploaded_files: list[tuple[str, bytes]] | None = None,
    file_urls: list[str] | None = None,
    callback_url: str = "",
    tomador: str = "",
) -> None:
    """Execute the full comparative generation pipeline.

    This function runs as a BackgroundTask and handles its own error reporting.

    Args:
        job_id: Unique job identifier.
        process_id: External process ID (Flokzu / frontend).
        uploaded_files: List of (filename, bytes) from multipart upload.
        file_urls: List of URLs to download quotes from.
        callback_url: Webhook URL for result notification.
        tomador: TOMADOR field value (empty = to fill later).
    """
    _update_job(job_id, status=JobStatus.PROCESSING)
    logger.info("=== Iniciando job %s (process_id=%s) ===", job_id, process_id)

    try:
        # ---- Step 1: Gather all documents ----
        documents: list[tuple[str, bytes]] = []

        if uploaded_files:
            documents.extend(uploaded_files)
            logger.info("Recibidos %d archivos subidos", len(uploaded_files))

        if file_urls:
            logger.info("Descargando %d archivos desde URLs", len(file_urls))
            download_tasks = [download_file(url) for url in file_urls]
            downloaded = await asyncio.gather(*download_tasks, return_exceptions=True)
            for result in downloaded:
                if isinstance(result, Exception):
                    logger.error("Error descargando archivo: %s", result)
                    raise result
                documents.append(result)

        if not documents:
            raise ValueError("No se recibieron archivos para procesar")

        logger.info("Total documentos a procesar: %d", len(documents))

        # ---- Step 2: Parse text from each document ----
        texts: list[str] = []
        for filename, file_bytes in documents:
            logger.info("Parseando: %s (%d bytes)", filename, len(file_bytes))
            text = parse_document(filename, file_bytes)
            texts.append(text)
            logger.info("Texto extraído de %s: %d caracteres", filename, len(text))

        # ---- Step 3: MAP — Extract each quote in parallel with Gemini ----
        logger.info("Extrayendo %d cotizaciones con Gemini (paralelo)...", len(texts))
        extraction_tasks = [extract_quote(text) for text in texts]
        quotes = await asyncio.gather(*extraction_tasks)
        logger.info(
            "Cotizaciones extraídas: %s",
            [q.nombre_compania for q in quotes],
        )

        # ---- Step 4: Extract project metadata ----
        logger.info("Extrayendo metadata del proyecto...")
        project_meta = await extract_project_meta(texts)

        # ---- Step 5: Generate recommendation and worst markers in parallel ----
        logger.info("Generando recomendación y marcadores...")

        # Build a preliminary consolidated structure for the AI analysis
        preliminary = consolidate_quotes(
            quotes=list(quotes),
            project_meta=project_meta,
            tomador=tomador,
        )
        preliminary_dict = json.loads(preliminary.model_dump_json())

        rec_task = generate_recommendation(preliminary_dict)
        markers_task = generate_worst_markers(preliminary_dict)
        recommendation, worst_markers = await asyncio.gather(rec_task, markers_task)

        # ---- Step 6: REDUCE — Final consolidation ----
        consolidated = consolidate_quotes(
            quotes=list(quotes),
            project_meta=project_meta,
            tomador=tomador,
            recommendation=recommendation,
            worst_markers=worst_markers,
        )
        data_dict = json.loads(consolidated.model_dump_json())
        logger.info("Consolidación completada: %d aseguradoras", len(consolidated.aseguradoras))

        # ---- Step 7: Render Word document ----
        output_filename = f"COMPARATIVO_FASE_2_{process_id}_{job_id[:8]}.docx"
        output_path = str(Path(tempfile.gettempdir()) / output_filename)

        assets_dir = str(settings.ASSETS_DIR)
        logger.info("Renderizando Word en: %s (assets: %s)", output_path, assets_dir)
        render_comparative(data_dict, output_path, assets_dir)

        # ---- Step 8: Upload to Supabase Storage ----
        logger.info("Subiendo a Supabase Storage...")
        download_url = await upload_to_supabase(output_path)
        logger.info("Archivo subido: %s", download_url)

        # ---- Step 9: Update job state ----
        _update_job(
            job_id,
            status=JobStatus.COMPLETED,
            download_url=download_url,
        )

        # ---- Step 10: Send callback ----
        if callback_url:
            await _send_callback(
                callback_url,
                {
                    "job_id": job_id,
                    "process_id": process_id,
                    "status": "completed",
                    "download_url": download_url,
                },
            )

        logger.info("=== Job %s completado exitosamente ===", job_id)

    except Exception as e:
        error_msg = f"{type(e).__name__}: {e}"
        logger.exception("=== Job %s FALLÓ: %s ===", job_id, error_msg)

        _update_job(
            job_id,
            status=JobStatus.FAILED,
            error=error_msg,
        )

        # Try to send failure callback
        if callback_url:
            try:
                await _send_callback(
                    callback_url,
                    {
                        "job_id": job_id,
                        "process_id": process_id,
                        "status": "failed",
                        "error": error_msg,
                    },
                )
            except Exception as cb_err:
                logger.error(
                    "No se pudo enviar callback de error: %s", cb_err
                )
