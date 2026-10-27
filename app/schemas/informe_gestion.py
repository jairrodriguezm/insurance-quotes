# -*- coding: utf-8 -*-
"""Pydantic schemas for Management Reports (Informes de Gestión)."""
from __future__ import annotations

from typing import Optional
from pydantic import BaseModel, Field


class PeriodoOption(BaseModel):
    """Available month/period in the uploaded Flokzu Excel file."""
    periodo: str = Field(..., description="YYYY-MM period identifier, e.g. '2026-09'")
    label: str = Field(..., description="Human-readable period label, e.g. 'Septiembre de 2026'")
    count: int = Field(..., description="Total requests in this period")


class ClienteSuggestion(BaseModel):
    """Detected client/entity from the uploaded Flokzu Excel file."""
    id: str = Field(..., description="Normalized client identifier or key")
    nombre_corto: str = Field(..., description="Short name or acronym, e.g. 'SONACOL'")
    nombre_completo: str = Field(..., description="Full display name for report title")
    nit: str = Field("", description="Formatted NIT if known or detected, e.g. '830.129.289-8'")
    terminos_busqueda: list[str] = Field(
        default_factory=list,
        description="Keywords to match in Tomador, Asegurado, or Integrantes",
    )
    total_solicitudes: int = Field(0, description="Total matching requests across the file")
    solicitudes_por_periodo: dict[str, int] = Field(
        default_factory=dict,
        description="Count of matching requests per YYYY-MM period",
    )


class InformeInspectResponse(BaseModel):
    """Response returned when inspecting a Flokzu .xlsx export."""
    ramo: str = Field(..., description="Insurance branch, e.g. 'Seguros de cumplimiento'")
    ramo_codigo: str = Field(..., description="Short branch code, e.g. 'CUMP'")
    fecha_corte: str = Field(..., description="Cutoff date YYYY-MM-DD from the Excel export")
    total_filas: int = Field(..., description="Total request rows in the Excel file")
    periodos: list[PeriodoOption] = Field(default_factory=list)
    clientes_sugeridos: list[ClienteSuggestion] = Field(default_factory=list)


class InformeJobResponse(BaseModel):
    """Response returned when creating a management report generation job."""
    status: str = "queued"
    job_id: str
    download_url: Optional[str] = None


class InformeJobStatusResponse(BaseModel):
    """Status response for a management report generation job."""
    job_id: str
    status: str
    download_url: Optional[str] = None
    error: Optional[str] = None
    resumen: Optional[dict] = None
