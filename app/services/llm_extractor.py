# -*- coding: utf-8 -*-
"""LLM Extractor service — calls Gemini with structured outputs to extract insurance quote data.

Uses the Map pattern: each quote document is processed individually, producing
an ExtractedQuote that can later be consolidated.

Temperature is set to 0.0 to maximize determinism and prevent hallucinations.
"""
from __future__ import annotations

import json
import logging
from typing import Any

from google import genai
from google.genai import types

from app.core.config import settings
from app.core.exceptions import LLMExtractionError
from app.schemas.quote import ExtractedQuote, ProjectMeta
from app.schemas.comparative import Recommendation

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# System prompts
# ---------------------------------------------------------------------------

EXTRACTION_SYSTEM_PROMPT = """Eres un analista técnico de seguros experto. Tu tarea es extraer datos estructurados de una cotización de seguros colombiana.

## REGLAS ESTRICTAS
1. **TRANSCRIBE, NO INTERPRETES.** Copia los valores tal como aparecen en el documento. No redondees, no conviertas unidades, no deduzcas información.
2. Si un dato NO aparece en el documento, **OMITE la clave** del diccionario. No inventes valores.
3. Si la cotización excluye expresamente un amparo, usa el valor **"N/A"**.
4. Los montos puros (valor asegurado, primas) van como **números** (int o float) SIN separadores.
5. Los sublímites con texto descriptivo van como **strings** tal cual: "30% del siniestro. Máximo $10.000.000.000".
6. Los deducibles SIEMPRE son texto (string). Transcribe porcentaje y mínimo tal cual: "10% del siniestro, mínimo 5 SMMLV". No conviertas SMMLV a pesos.
7. La tasa va en formato texto: "1,10 ‰".
8. La prima es un NÚMERO en pesos colombianos.

## CATÁLOGO DE COBERTURAS (claves permitidas en "coberturas")
### Sección 1 — Todo Riesgo Construcción
- danos_materiales: Daños materiales (Cobertura A)
- terremoto: Terremoto, temblor y/o erupción volcánica (B)
- tormenta_inundacion: Tormenta e inundación (C)
- mantenimiento_amplio: Mantenimiento amplio (D)
- remocion_escombros: Remoción de escombros (G)
- hmacc_amit: Huelga, asonada, motín, conmoción civil o popular (HMACC/AMIT)
- hurto_calificado: Hurto calificado
- cronograma_avance: Cronograma de avance (Endoso 005)
- gastos_horas_extra: Horas extra, trabajo nocturno, flete (Endoso 006/007)
- obras_zona_sismica: Obras en zona sísmica (Endoso 008)
- bienes_fuera_sitio: Bienes almacenados fuera del sitio (Endoso 013)
- prueba_maquinaria: Prueba de maquinarias e instalaciones (Endoso 100)
- campamentos_almacenes: Campamentos y almacenes (Endoso 107/109/112)
- medidas_inundacion: Medidas de seguridad contra inundaciones (Endoso 110)
- proteccion_incendio: Protección contra incendio (Endoso 112)
- transportes_nacionales: Transportes nacionales
- siniestros_serie: Siniestros en serie (Endoso 114)
- obras_civiles_operacion: Obras civiles puestas en operación (Endoso 116)
- cimentacion_pilotaje: Cimentación por pilotaje y tablestacados (Endoso 121)
- hundimiento_subsuelo: Hundimiento y asentamiento del subsuelo (Endoso 214)
- error_diseno: Error de diseño (Cláusulas LEG 2/96 y LEG 3/06)
- honorarios_profesionales: Honorarios profesionales
- planos_documentos: Planos y documentos
- gastos_extincion: Gastos de extinción del siniestro
- condiciones_remocion: Remoción de escombros tras corrimiento de tierras (Endoso 111)
- actos_autoridad: Actos de autoridad
- preservacion_bienes: Preservación de bienes

### Sección 2 — Responsabilidad Civil Extracontractual
- rce: RCE (amparos E y F). Indicar si es límite único combinado o separado.
- rce_contratistas: Contratistas y subcontratistas independientes
- rce_patronal: Civil patronal
- rce_cruzada: Civil cruzada. Anotar si el sublímite está dentro de E y F.
- rce_vehiculos: Vehículos propios y no propios
- rce_cuidado_control: Bienes bajo cuidado, tenencia y control (Endoso 119)
- rce_contaminacion: Contaminación súbita y accidental
- rce_vibracion: Vibración y debilitamiento de elementos portantes (Endoso 120)
- rce_subterraneas: Conducciones y servicios subterráneos (Endoso 102)

### Sección 3
- terrorismo: Terrorismo

## CATÁLOGO DE DEDUCIBLES (claves permitidas en "deducibles")
ded_incendio, ded_terremoto, ded_tormenta, ded_mantenimiento, ded_remocion,
ded_hmacc, ded_hurto, ded_error_diseno, ded_cables, ded_adyacentes,
ded_campamentos, ded_hundimiento, ded_aeronaves, ded_impericia,
ded_corto_circuito, ded_rce, ded_rce_cruzada, ded_propiedades_existentes

## PARTICULARIDADES POR ASEGURADORA
- SURA: sublímites como "% del siniestro, máximo $X"; RCE con límite único combinado E+F.
- Seguros Mundial: todo por endoso numerado, sublímites "evento / vigencia"; frecuentemente en coaseguro con Berkley.
- Chubb: estructura en Sección I / II / III; deducibles mínimos altos en SMMLV; suele traer tres opciones LEG.
- AXA Colpatria, Bolívar, Estado: verificar si terrorismo está incluido en HMACC o cotizado aparte.

Extrae toda la información del documento y devuelve el JSON estructurado."""


META_EXTRACTION_PROMPT = """Eres un analista técnico de seguros. Analiza las siguientes cotizaciones y extrae la información general del proyecto asegurado.

## REGLAS
1. Infiere los datos del proyecto a partir de las cotizaciones.
2. Si un dato no aparece en ninguna cotización, déjalo vacío.
3. El valor asegurado debe ser un número sin separadores.
4. Las fechas en formato dd/mm/aaaa.

Extrae la información general del proyecto a partir de estos textos de cotizaciones."""


RECOMMENDATION_PROMPT = """Eres un analista técnico de seguros de Multiriesgos de Colombia Ltda. Analiza los datos comparativos consolidados de las siguientes aseguradoras y genera una recomendación OBJETIVA y CONCISA.

## CRITERIOS DE EVALUACIÓN (en orden de prioridad)
1. **Amplitud de cobertura**: sublímites, coberturas adicionales, exclusiones.
2. **Prima**: relación costo-beneficio, no solo el precio más bajo.
3. **Deducibles**: porcentajes y mínimos en SMMLV comparados.
4. **Requisitos del financiador**: cumplimiento de LEG 2/96 sobre 100% del valor.
5. **Respaldo/Coaseguro**: solidez de la(s) aseguradora(s).

## REGLAS
1. La recomendación debe basarse EXCLUSIVAMENTE en los datos proporcionados.
2. Las viñetas deben ser argumentos objetivos y verificables.
3. NO inventes datos que no estén en el comparativo.
4. Incluye al menos un punto de validación o negociación pendiente si existe.

Datos consolidados:"""


WORST_MARKERS_PROMPT = """Eres un analista técnico de seguros. Analiza los datos consolidados y determina qué aseguradoras tienen la PEOR oferta en cada cobertura y deducible.

## REGLAS
1. Solo marca como "peor" cuando la diferencia es CLARAMENTE inferior y defendible ante un cliente.
2. Si todas las aseguradoras son similares en un rubro, NO incluyas esa clave.
3. Usa los IDs de aseguradora (ej: "sura", "chubb", "mundial").
4. Devuelve un diccionario donde la clave es la clave de cobertura/deducible y el valor es una lista de IDs.
5. NO marques rubros donde alguna aseguradora no especifica (NO ESPECIFICA no es "peor").

Datos consolidados:"""


# ---------------------------------------------------------------------------
# Client management
# ---------------------------------------------------------------------------

def _get_client() -> genai.Client:
    """Create a Gemini API client."""
    return genai.Client(api_key=settings.GEMINI_API_KEY)


# ---------------------------------------------------------------------------
# Extraction functions
# ---------------------------------------------------------------------------

async def extract_quote(document_text: str) -> ExtractedQuote:
    """Extract structured quote data from a document's text using Gemini.

    This is the MAP step of the Map-Reduce pattern.

    Args:
        document_text: Plain text extracted from a quote document.

    Returns:
        Structured extraction of the quote.

    Raises:
        LLMExtractionError: If extraction fails.
    """
    try:
        client = _get_client()

        response = await client.aio.models.generate_content(
            model=settings.GEMINI_MODEL,
            contents=[
                EXTRACTION_SYSTEM_PROMPT,
                f"## DOCUMENTO DE COTIZACIÓN A ANALIZAR:\n\n{document_text}",
            ],
            config=types.GenerateContentConfig(
                temperature=0.0,
                response_mime_type="application/json",
                response_json_schema=ExtractedQuote.model_json_schema(),
            ),
        )

        if not response.text:
            raise LLMExtractionError(
                message="Gemini no retornó texto en la respuesta"
            )

        result = ExtractedQuote.model_validate_json(response.text)
        logger.info(
            "Cotización extraída: %s (%d coberturas, %d deducibles)",
            result.nombre_compania,
            len(result.coberturas),
            len(result.deducibles),
        )
        return result

    except LLMExtractionError:
        raise
    except Exception as e:
        raise LLMExtractionError(
            message=f"Error al extraer cotización con Gemini: {e}",
            detail=str(e),
        ) from e


async def extract_project_meta(document_texts: list[str]) -> ProjectMeta:
    """Extract project metadata from multiple quote texts.

    Args:
        document_texts: List of plain text from all quote documents.

    Returns:
        Inferred project metadata.

    Raises:
        LLMExtractionError: If extraction fails.
    """
    try:
        client = _get_client()

        combined = "\n\n---\n\n".join(
            f"COTIZACIÓN {i + 1}:\n{text[:3000]}"
            for i, text in enumerate(document_texts)
        )

        response = await client.aio.models.generate_content(
            model=settings.GEMINI_MODEL,
            contents=[
                META_EXTRACTION_PROMPT,
                combined,
            ],
            config=types.GenerateContentConfig(
                temperature=0.0,
                response_mime_type="application/json",
                response_json_schema=ProjectMeta.model_json_schema(),
            ),
        )

        if not response.text:
            raise LLMExtractionError(
                message="Gemini no retornó metadata del proyecto"
            )

        return ProjectMeta.model_validate_json(response.text)

    except LLMExtractionError:
        raise
    except Exception as e:
        raise LLMExtractionError(
            message=f"Error al extraer metadata del proyecto: {e}",
            detail=str(e),
        ) from e


async def generate_recommendation(consolidated_data: dict) -> Recommendation:
    """Generate an AI recommendation based on consolidated comparative data.

    Args:
        consolidated_data: The full consolidated data dict.

    Returns:
        Recommendation with insurer ID, option, and justification bullets.

    Raises:
        LLMExtractionError: If generation fails.
    """
    try:
        client = _get_client()

        # Prepare a focused summary for the LLM
        summary = json.dumps(consolidated_data, ensure_ascii=False, indent=2)

        response = await client.aio.models.generate_content(
            model=settings.GEMINI_MODEL,
            contents=[
                RECOMMENDATION_PROMPT,
                summary,
            ],
            config=types.GenerateContentConfig(
                temperature=0.0,
                response_mime_type="application/json",
                response_json_schema=Recommendation.model_json_schema(),
            ),
        )

        if not response.text:
            raise LLMExtractionError(
                message="Gemini no retornó recomendación"
            )

        rec = Recommendation.model_validate_json(response.text)
        logger.info("Recomendación generada: %s - %s", rec.aseguradora_id, rec.opcion)
        return rec

    except LLMExtractionError:
        raise
    except Exception as e:
        raise LLMExtractionError(
            message=f"Error al generar recomendación: {e}",
            detail=str(e),
        ) from e


async def generate_worst_markers(consolidated_data: dict) -> dict[str, list[str]]:
    """Identify the worst offerings per coverage/deductible row.

    Args:
        consolidated_data: The full consolidated data dict.

    Returns:
        Dict mapping coverage/deductible keys to lists of insurer IDs.

    Raises:
        LLMExtractionError: If generation fails.
    """
    try:
        client = _get_client()

        summary = json.dumps(consolidated_data, ensure_ascii=False, indent=2)

        response = await client.aio.models.generate_content(
            model=settings.GEMINI_MODEL,
            contents=[
                WORST_MARKERS_PROMPT,
                summary,
            ],
            config=types.GenerateContentConfig(
                temperature=0.0,
                response_mime_type="application/json",
            ),
        )

        if not response.text:
            return {}

        markers = json.loads(response.text)
        if not isinstance(markers, dict):
            return {}

        logger.info("Marcadores 'peor' generados: %d claves", len(markers))
        return markers

    except Exception as e:
        logger.warning("Error al generar marcadores peor (no crítico): %s", e)
        return {}
