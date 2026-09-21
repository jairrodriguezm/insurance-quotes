# -*- coding: utf-8 -*-
"""Pydantic v2 schemas for individual quote extraction and project metadata."""
from __future__ import annotations

from typing import Optional, Union
from pydantic import BaseModel, Field


class CoaseguroEntry(BaseModel):
    """Entry for a co-insurance participant."""

    nombre: str = Field(description="Nombre de la compañía participante en el coaseguro")
    participacion: str = Field(description="Porcentaje de participación, ej: '35%'")
    logo: Optional[str] = Field(default=None, description="Nombre del archivo de logo, ej: 'mundial.png'")


class QuoteOption(BaseModel):
    """A pricing option within a quote."""

    etiqueta: str = Field(description="Identificador de la opción, ej: 'Opción 1', 'Opción 2'")
    tasa: str = Field(description="Tasa en por mil (‰), ej: '1,10 ‰'")
    prima: float = Field(description="Valor numérico de la prima en pesos colombianos, sin separadores de miles")
    modalidad: str = Field(description="Modalidad de aseguramiento, ej: 'LEG 2/96 100% del valor asegurado'")


class ExtractedQuote(BaseModel):
    """Structured extraction of a single insurance quote from a document.

    This schema is used as the response_json_schema for Gemini structured output.
    Field descriptions guide the LLM extraction.
    """

    nombre_compania: str = Field(
        description="Nombre completo oficial de la compañía aseguradora tal como aparece en la cotización"
    )
    id_compania: str = Field(
        description="ID normalizado en minúsculas sin espacios ni tildes. Ejemplos: 'sura', 'chubb', 'axa_colpatria', 'mundial', 'hdi', 'zurich', 'bolivar'"
    )
    opciones: list[QuoteOption] = Field(
        description="Lista de opciones de cotización. Cada opción tiene tasa, prima y modalidad distintas."
    )
    coberturas: dict[str, Optional[Union[int, float, str]]] = Field(
        default_factory=dict,
        description=(
            "Diccionario de coberturas extraídas. Las claves DEBEN ser del catálogo estándar: "
            "danos_materiales, terremoto, tormenta_inundacion, mantenimiento_amplio, remocion_escombros, "
            "hmacc_amit, hurto_calificado, cronograma_avance, gastos_horas_extra, obras_zona_sismica, "
            "bienes_fuera_sitio, prueba_maquinaria, campamentos_almacenes, medidas_inundacion, "
            "proteccion_incendio, transportes_nacionales, siniestros_serie, obras_civiles_operacion, "
            "cimentacion_pilotaje, hundimiento_subsuelo, error_diseno, honorarios_profesionales, "
            "planos_documentos, gastos_extincion, condiciones_remocion, actos_autoridad, "
            "preservacion_bienes, rce, rce_contratistas, rce_patronal, rce_cruzada, rce_vehiculos, "
            "rce_cuidado_control, rce_contaminacion, rce_vibracion, rce_subterraneas, terrorismo. "
            "Los valores pueden ser numéricos (int/float para montos puros) o texto (str para descripciones). "
            "Si un dato NO se menciona en el documento, OMITIR la clave (no inventar). "
            "Si se excluye expresamente, usar 'N/A'."
        ),
    )
    deducibles: dict[str, Optional[str]] = Field(
        default_factory=dict,
        description=(
            "Diccionario de deducibles extraídos. Las claves DEBEN ser del catálogo: "
            "ded_incendio, ded_terremoto, ded_tormenta, ded_mantenimiento, ded_remocion, ded_hmacc, "
            "ded_hurto, ded_error_diseno, ded_cables, ded_adyacentes, ded_campamentos, ded_hundimiento, "
            "ded_aeronaves, ded_impericia, ded_corto_circuito, ded_rce, ded_rce_cruzada, "
            "ded_propiedades_existentes. "
            "Los valores siempre son texto: transcribir porcentaje y mínimo tal cual aparecen en la cotización. "
            "No convertir SMMLV a pesos ni viceversa."
        ),
    )
    subjetividades: list[str] = Field(
        default_factory=list,
        description="Lista de subjetividades, condiciones pendientes, garantías o documentos requeridos por la aseguradora.",
    )
    tasa_prorroga: str = Field(
        default="NO ESPECIFICA",
        description="Tasa de prórroga tal cual la especifica la aseguradora. Si no la menciona, usar 'NO ESPECIFICA'.",
    )
    coaseguro: Optional[list[CoaseguroEntry]] = Field(
        default=None,
        description="Si la cotización es en coaseguro (varias compañías), listar las participantes con su porcentaje.",
    )
    validez_oferta: str = Field(
        default="NO ESPECIFICA",
        description="Plazo de validez de la oferta, ej: '15/05/2026' o '15 días calendario'",
    )


class ProjectMeta(BaseModel):
    """Metadata about the insured project, inferred from the quotes."""

    tipo_cobertura: str = Field(
        default="TODO RIESGO CONSTRUCCIÓN Y MONTAJE",
        description="Tipo de póliza, ej: 'TODO RIESGO CONSTRUCCIÓN Y MONTAJE', 'TODO RIESGO DAÑO MATERIAL'",
    )
    asegurado: str = Field(
        default="",
        description="Nombre del asegurado / propietario del proyecto",
    )
    beneficiario: str = Field(
        default="",
        description="Beneficiario de la póliza",
    )
    vigencia_construccion: Optional[dict[str, str]] = Field(
        default=None,
        description="{'desde': 'dd/mm/aaaa', 'hasta': 'dd/mm/aaaa'}",
    )
    vigencia_mantenimiento: Optional[dict[str, str]] = Field(
        default=None,
        description="{'desde': '...', 'hasta': '...', 'tipo': 'Amplio', 'duracion': '12 meses'}",
    )
    ubicacion: str = Field(default="", description="Ubicación del proyecto")
    descripcion_proyecto: str = Field(default="", description="Descripción del proyecto a asegurar")
    valor_asegurado: Optional[Union[int, float]] = Field(
        default=None,
        description="Valor total asegurado en pesos colombianos, sin separadores",
    )
