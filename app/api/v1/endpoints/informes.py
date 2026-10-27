# -*- coding: utf-8 -*-
"""Management Reports (Informes de Gestión) API endpoints.

Completely isolated from the insurance quotes comparative logic.
Endpoints:
- POST /api/v1/informes/inspect -> Inspect uploaded .xlsx (periods, suggested clients, cutoff date)
- POST /api/v1/informes/jobs -> Generate .docx management report from .xlsx
- GET  /api/v1/informes/jobs/{job_id}/status -> Check report job status & summary KPIs
- GET  /api/v1/informes/jobs/{job_id}/download -> Direct file download endpoint
"""
from __future__ import annotations

import json
import logging
import re
import tempfile
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from fastapi import APIRouter, BackgroundTasks, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse

from app.core.config import settings
from app.schemas.informe_gestion import (
    InformeInspectResponse,
    InformeJobResponse,
    InformeJobStatusResponse,
)
from app.services.document_parser import download_file
from app.services.informe_gestion_parser import inspect_flokzu_xlsx, process_flokzu_xlsx
from app.services.informe_gestion_renderer import render_informe_gestion
from app.services.storage import upload_to_supabase

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/informes", tags=["Informes de Gestión"])


@dataclass
class InformeJobState:
    """State of a management report generation job."""
    job_id: str
    status: str = "queued"
    download_url: str | None = None
    local_path: str | None = None
    filename: str | None = None
    error: str | None = None
    resumen: dict[str, Any] | None = None
    created_at: datetime = field(default_factory=datetime.now)


_informe_jobs: dict[str, InformeJobState] = {}


def _parse_search_terms(raw_terms: str) -> list[str]:
    """Parse search terms from JSON array or comma/semicolon-separated string."""
    if not raw_terms or not raw_terms.strip():
        return []
    txt = raw_terms.strip()
    if txt.startswith("["):
        try:
            parsed = json.loads(txt)
            if isinstance(parsed, list):
                return [str(x).strip() for x in parsed if str(x).strip()]
        except Exception:
            pass
    return [t.strip() for t in re.split(r"[,;]+", txt) if t.strip()]


@router.post(
    "/inspect",
    response_model=InformeInspectResponse,
    summary="Inspect an uploaded Flokzu .xlsx export",
    description=(
        "Reads a Flokzu .xlsx file and returns available months/periods, "
        "suggested clients/entities with request counts and NITs, and cutoff date."
    ),
)
async def inspect_informe_xlsx(
    file: UploadFile = File(..., description="Archivo Excel (.xlsx) exportado de Flokzu"),
) -> InformeInspectResponse:
    """Inspect an uploaded Excel file to populate frontend selectors."""
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(
            status_code=400,
            detail="El archivo debe ser un documento Excel (.xlsx).",
        )
    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="El archivo Excel está vacío.")

    try:
        return inspect_flokzu_xlsx(content)
    except Exception as e:
        logger.exception("Error inspeccionando archivo Excel de informe: %s", e)
        raise HTTPException(
            status_code=400,
            detail=f"No se pudo leer el archivo Excel: {e}",
        )


async def _run_informe_job(
    job_id: str,
    file_bytes: bytes | None,
    file_url: str | None,
    periodo: str,
    cliente_nombre_corto: str,
    cliente_nombre_completo: str,
    cliente_nit: str,
    terminos_busqueda: list[str],
    base_url: str,
) -> None:
    """Background worker to process the .xlsx and render the .docx management report."""
    job = _informe_jobs.get(job_id)
    if not job:
        return
    job.status = "processing"

    try:
        if file_bytes is None and file_url:
            _, file_bytes = await download_file(file_url)
        if not file_bytes:
            raise ValueError("No se recibieron datos del archivo Excel.")

        report_data = process_flokzu_xlsx(
            file_bytes=file_bytes,
            periodo=periodo,
            cliente_nombre_corto=cliente_nombre_corto,
            cliente_nombre_completo=cliente_nombre_completo,
            cliente_nit=cliente_nit,
            terminos_busqueda=terminos_busqueda,
        )

        if report_data["kpis"]["tramitadas"] == 0:
            raise ValueError(
                f"No se encontraron solicitudes para '{cliente_nombre_corto}' en el período '{periodo}'."
            )

        safe_client = re.sub(r"[^A-Za-z0-9_-]+", "_", report_data["cliente_nombre_corto"]).strip("_") or "CLIENTE"
        safe_period = re.sub(r"[^A-Za-z0-9_-]+", "-", report_data["periodo"]).strip("-") or "PERIODO"
        output_filename = f"Informe_Gestion_{safe_client}_{safe_period}.docx"
        output_path = Path(tempfile.gettempdir()) / f"{job_id[:8]}_{output_filename}"

        render_informe_gestion(
            data=report_data,
            output_path=output_path,
            assets_dir=settings.ASSETS_DIR,
        )

        job.local_path = str(output_path)
        job.filename = output_filename

        # Upload to cloud storage (Supabase) if configured; fallback to direct API download URL
        download_url: str | None = None
        if settings.SUPABASE_URL and settings.SUPABASE_KEY:
            try:
                download_url = await upload_to_supabase(str(output_path))
            except Exception as sup_err:
                logger.warning("Fallback a descarga directa (error en almacenamiento nube): %s", sup_err)

        if not download_url:
            download_url = f"{base_url.rstrip('/')}/api/v1/informes/jobs/{job_id}/download"

        job.download_url = download_url
        job.resumen = {
            "cliente": report_data["cliente_nombre_completo"],
            "periodo": report_data["periodo_titulo"],
            "fecha_corte": report_data["fecha_corte"],
            "kpis": report_data["kpis"],
        }
        job.status = "completed"
        logger.info("Informe job %s completado exitosamente (%s)", job_id, download_url)

    except Exception as e:
        logger.exception("Error generando informe de gestión (job %s): %s", job_id, e)
        job.status = "failed"
        job.error = str(e)


@router.post(
    "/jobs",
    status_code=202,
    response_model=InformeJobResponse,
    summary="Create a Management Report (.docx) generation job from an .xlsx file",
)
async def create_informe_job(
    request: Request,
    background_tasks: BackgroundTasks,
    periodo: str = Form("2026-09", description="Período en formato YYYY-MM (ej. 2026-09) o TODOS"),
    cliente_nombre_corto: str = Form("SONACOL", description="Nombre corto o sigla del cliente (ej. SONACOL)"),
    cliente_nombre_completo: str = Form(
        "SONACOL – Solarte Nacional de Construcciones S.A.S.",
        description="Nombre completo del cliente para la portada del informe",
    ),
    cliente_nit: str = Form(
        "830.129.289-8",
        description="NIT del cliente (ej. 830.129.289-8)",
    ),
    terminos_busqueda: str = Form(
        "",
        description="Términos de búsqueda separados por coma para filtrar el cliente en el Excel",
    ),
    file: Optional[UploadFile] = File(None, description="Archivo Excel (.xlsx) exportado de Flokzu"),
    file_url: Optional[str] = Form(None, description="URL alternativa para descargar el archivo .xlsx"),
) -> InformeJobResponse:
    """Create and launch a management report generation job."""
    has_file = file is not None and bool(file.filename)
    if not has_file and not file_url:
        raise HTTPException(
            status_code=400,
            detail="Debe adjuntar un archivo Excel (.xlsx) o proporcionar file_url.",
        )

    file_bytes: bytes | None = None
    if has_file and file is not None:
        if not (file.filename or "").lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(
                status_code=400,
                detail="Formato no soportado. Solo se admiten archivos Excel (.xlsx).",
            )
        file_bytes = await file.read()
        if not file_bytes:
            raise HTTPException(status_code=400, detail="El archivo Excel subido está vacío.")

    job_id = str(uuid.uuid4())
    _informe_jobs[job_id] = InformeJobState(job_id=job_id)

    terms_list = _parse_search_terms(terminos_busqueda)
    base_url = str(request.base_url)

    background_tasks.add_task(
        _run_informe_job,
        job_id=job_id,
        file_bytes=file_bytes,
        file_url=file_url,
        periodo=periodo,
        cliente_nombre_corto=cliente_nombre_corto,
        cliente_nombre_completo=cliente_nombre_completo,
        cliente_nit=cliente_nit,
        terminos_busqueda=terms_list,
        base_url=base_url,
    )

    return InformeJobResponse(status="queued", job_id=job_id)


@router.get(
    "/jobs/{job_id}/status",
    response_model=InformeJobStatusResponse,
    summary="Get Management Report job status",
)
async def get_informe_job_status(job_id: str) -> InformeJobStatusResponse:
    """Check the status of a management report generation job."""
    job = _informe_jobs.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail=f"Job de informe no encontrado: {job_id}")
    return InformeJobStatusResponse(
        job_id=job.job_id,
        status=job.status,
        download_url=job.download_url,
        error=job.error,
        resumen=job.resumen,
    )


@router.get(
    "/jobs/{job_id}/download",
    summary="Download the generated Management Report (.docx)",
)
async def download_informe_docx(job_id: str) -> FileResponse:
    """Direct download endpoint for the generated management report .docx."""
    job = _informe_jobs.get(job_id)
    if job is None or not job.local_path or not Path(job.local_path).exists():
        raise HTTPException(status_code=404, detail="El archivo del informe no está disponible.")
    return FileResponse(
        path=job.local_path,
        filename=job.filename or f"Informe_Gestion_{job_id[:8]}.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
