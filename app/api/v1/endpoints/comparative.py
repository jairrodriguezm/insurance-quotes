# -*- coding: utf-8 -*-
"""Comparative API endpoints — handles job creation and status queries.

POST /api/v1/comparatives/jobs  → Accepts files (multipart or URLs), returns 202 + job_id
GET  /api/v1/comparatives/jobs/{job_id}/status → Returns current job status
"""
from __future__ import annotations

import json
import logging
from typing import Optional

from fastapi import APIRouter, BackgroundTasks, File, Form, HTTPException, UploadFile

from app.core.config import settings
from app.schemas.comparative import JobResponse, JobStatusResponse
from app.services.job_runner import JobStatus, create_job, get_job, run_comparative_job

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/comparatives", tags=["Comparatives"])


@router.post(
    "/jobs",
    status_code=202,
    response_model=JobResponse,
    summary="Create a comparative analysis job",
    description=(
        "Accepts insurance quote files (PDF/DOCX) via multipart upload or "
        "as a list of URLs. Returns immediately with a job_id. The processing "
        "runs in the background and notifies via callback_url when done."
    ),
)
async def create_comparative_job(
    background_tasks: BackgroundTasks,
    process_id: str = Form(
        ...,
        description="Process identifier from Flokzu BPMS or the frontend",
    ),
    callback_url: str = Form(
        "",
        description="Webhook URL to notify when the job completes or fails (opcional)",
    ),
    tomador: str = Form(
        "",
        description="TOMADOR field value. Leave empty to fill later.",
    ),
    files: Optional[list[UploadFile]] = File(
        None,
        description="Quote files to process (PDF or DOCX)",
    ),
    file_urls: Optional[str] = Form(
        None,
        description='JSON array of URLs to download quote files from, e.g. ["https://...pdf", "https://...docx"]',
    ),
) -> JobResponse:
    """Create a new comparative analysis job.

    Accepts quote files in two ways (can be combined):
    - **Multipart upload:** Send files directly via the `files` form field.
    - **URL list:** Provide a JSON array of URLs in the `file_urls` form field.

    Returns HTTP 202 with a `job_id` to track progress.
    """
    # Validate that at least one source of files is provided
    has_uploaded = files and any(f.filename for f in files)
    parsed_urls: list[str] = []

    if file_urls:
        try:
            parsed_urls = json.loads(file_urls)
            if not isinstance(parsed_urls, list):
                raise ValueError("file_urls must be a JSON array")
        except (json.JSONDecodeError, ValueError) as e:
            raise HTTPException(
                status_code=400,
                detail=f"file_urls must be a valid JSON array of strings: {e}",
            )

    if not has_uploaded and not parsed_urls:
        raise HTTPException(
            status_code=400,
            detail="Debe proporcionar al menos un archivo (files) o una URL (file_urls)",
        )

    # Validate file sizes and types
    uploaded_files: list[tuple[str, bytes]] = []
    if has_uploaded:
        max_size = settings.MAX_FILE_SIZE_MB * 1024 * 1024
        for f in files:
            if not f.filename:
                continue
            ext = f.filename.rsplit(".", 1)[-1].lower() if "." in f.filename else ""
            if ext not in ("pdf", "docx", "doc"):
                raise HTTPException(
                    status_code=400,
                    detail=f"Tipo de archivo no soportado: {f.filename}. Solo PDF y DOCX.",
                )
            content = await f.read()
            if len(content) > max_size:
                raise HTTPException(
                    status_code=400,
                    detail=f"El archivo {f.filename} excede el límite de {settings.MAX_FILE_SIZE_MB} MB",
                )
            uploaded_files.append((f.filename, content))

    # Create job and launch background processing
    job = create_job(process_id)
    logger.info(
        "Job creado: %s (process_id=%s, %d archivos, %d URLs)",
        job.job_id,
        process_id,
        len(uploaded_files),
        len(parsed_urls),
    )

    background_tasks.add_task(
        run_comparative_job,
        job_id=job.job_id,
        process_id=process_id,
        uploaded_files=uploaded_files if uploaded_files else None,
        file_urls=parsed_urls if parsed_urls else None,
        callback_url=callback_url,
        tomador=tomador,
    )

    return JobResponse(job_id=job.job_id)


@router.get(
    "/jobs/{job_id}/status",
    response_model=JobStatusResponse,
    summary="Get job status",
    description="Check the current status of a comparative analysis job.",
)
async def get_job_status(job_id: str) -> JobStatusResponse:
    """Query the current status of a comparative job."""
    job = get_job(job_id)
    if job is None:
        raise HTTPException(
            status_code=404,
            detail=f"Job no encontrado: {job_id}",
        )
    return JobStatusResponse(
        job_id=job.job_id,
        status=job.status.value,
        download_url=job.download_url,
        error=job.error,
    )
