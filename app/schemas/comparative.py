# -*- coding: utf-8 -*-
"""Pydantic v2 schemas for comparative jobs, requests, responses, and consolidated document data."""
from __future__ import annotations

from datetime import date
from typing import Literal, Optional, Union
from pydantic import BaseModel, Field, HttpUrl


class JobRequest(BaseModel):
    """Request body for creating a comparative job (used with JSON body, not multipart)."""

    process_id: str = Field(description="ID del proceso en Flokzu BPMS o frontend")
    callback_url: str = Field(description="URL del webhook para notificar cuando termine")
    tomador: str = Field(default="", description="Nombre del tomador. Vacío si se llenará después.")
    file_urls: Optional[list[str]] = Field(
        default=None,
        description="Lista de URLs de las cotizaciones a procesar (alternativa a subir archivos)",
    )


class JobResponse(BaseModel):
    """Immediate 202 Accepted response."""

    status: Literal["queued"] = "queued"
    job_id: str


class JobStatusResponse(BaseModel):
    """Response for the job status endpoint."""

    job_id: str
    status: Literal["queued", "processing", "completed", "failed"]
    download_url: Optional[str] = None
    error: Optional[str] = None


class CallbackPayload(BaseModel):
    """Payload sent to the callback URL when the job completes."""

    job_id: str
    process_id: str
    status: Literal["completed", "failed"]
    download_url: Optional[str] = None
    error: Optional[str] = None


class Recommendation(BaseModel):
    """AI-generated recommendation for the best insurance option."""

    aseguradora_id: str = Field(description="ID de la aseguradora recomendada")
    opcion: str = Field(description="Opción recomendada, ej: 'Opción 1 – LEG 2/96 100%'")
    vinetas: list[str] = Field(
        description="Lista de argumentos objetivos que sustentan la recomendación"
    )


class AnnexConfig(BaseModel):
    """Configuration for which annexes to include in the document."""

    leg: bool = True
    tasa_prorroga: bool = True
    preguntas_cliente: list[dict[str, str]] = Field(default_factory=list)
    observaciones_adicionales: list[str] = Field(default_factory=list)


class MetaInfo(BaseModel):
    """Complete metadata for the comparative document."""

    fecha: str = Field(default_factory=lambda: date.today().strftime("%d/%m/%Y"))
    tipo_cobertura: str = "TODO RIESGO CONSTRUCCIÓN Y MONTAJE"
    tomador: str = ""
    asegurado: str = ""
    beneficiario: str = ""
    vigencia_construccion: Optional[dict[str, str]] = None
    vigencia_mantenimiento: Optional[dict[str, str]] = None
    ubicacion: str = ""
    descripcion_proyecto: str = ""
    valor_asegurado: Optional[Union[int, float]] = None


class InsuranceCompanyData(BaseModel):
    """Full data for one insurance company in the consolidated comparative."""

    id: str
    nombre: str
    logo: str = ""
    coaseguro: Optional[list[dict]] = None
    coaseguro_nota: str = ""
    validez_oferta: str = "NO ESPECIFICA"
    opciones: list[dict] = Field(default_factory=list)
    tasa_prorroga: str = "NO ESPECIFICA"
    coberturas: dict[str, Optional[Union[int, float, str]]] = Field(default_factory=dict)
    deducibles: dict[str, Optional[str]] = Field(default_factory=dict)
    subjetividades: list[str] = Field(default_factory=list)


class ConsolidatedData(BaseModel):
    """Complete consolidated data structure (equivalent to datos.json).

    This is the structure consumed by the Word renderer.
    """

    meta: MetaInfo = Field(default_factory=MetaInfo)
    aseguradoras: list[InsuranceCompanyData] = Field(default_factory=list)
    marcadores_peor: dict[str, list[str]] = Field(default_factory=dict)
    recomendacion: Optional[Recommendation] = None
    anexos: AnnexConfig = Field(default_factory=AnnexConfig)
