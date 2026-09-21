# -*- coding: utf-8 -*-
"""Consolidator service — Map-Reduce pattern for combining extracted quotes.

Takes individually extracted quotes (Map outputs) and consolidates them into
a single unified data structure compatible with the Word renderer (datos.json format).
"""
from __future__ import annotations

import logging
from datetime import date
from typing import Any

from app.schemas.quote import ExtractedQuote, ProjectMeta
from app.schemas.comparative import (
    AnnexConfig,
    ConsolidatedData,
    InsuranceCompanyData,
    MetaInfo,
    Recommendation,
)

logger = logging.getLogger(__name__)


def _build_company_data(quote: ExtractedQuote) -> InsuranceCompanyData:
    """Convert an extracted quote into the company data format for the renderer.

    Args:
        quote: Individual extracted quote.

    Returns:
        InsuranceCompanyData ready for the consolidated structure.
    """
    # Build coaseguro list in the renderer's expected format
    coaseguro = None
    if quote.coaseguro:
        coaseguro = [
            {
                "nombre": entry.nombre,
                "participacion": entry.participacion,
                "logo": entry.logo or "",
            }
            for entry in quote.coaseguro
        ]

    # Build options list
    opciones = [
        {
            "etiqueta": opt.etiqueta,
            "tasa": opt.tasa,
            "prima": opt.prima,
            "modalidad": opt.modalidad,
        }
        for opt in quote.opciones
    ]

    return InsuranceCompanyData(
        id=quote.id_compania,
        nombre=quote.nombre_compania,
        logo=f"{quote.id_compania}.png",
        coaseguro=coaseguro,
        coaseguro_nota="",
        validez_oferta=quote.validez_oferta,
        opciones=opciones,
        tasa_prorroga=quote.tasa_prorroga,
        coberturas=quote.coberturas,
        deducibles=quote.deducibles,
        subjetividades=quote.subjetividades,
    )


def _build_meta(
    project_meta: ProjectMeta | None,
    tomador: str = "",
) -> MetaInfo:
    """Build the MetaInfo from extracted project metadata and overrides.

    Args:
        project_meta: AI-extracted project metadata (may be None).
        tomador: Override for the TOMADOR field (empty = to be filled later).

    Returns:
        MetaInfo for the consolidated document.
    """
    if project_meta is None:
        return MetaInfo(
            fecha=date.today().strftime("%d/%m/%Y"),
            tomador=tomador,
        )

    return MetaInfo(
        fecha=date.today().strftime("%d/%m/%Y"),
        tipo_cobertura=project_meta.tipo_cobertura,
        tomador=tomador,
        asegurado=project_meta.asegurado,
        beneficiario=project_meta.beneficiario,
        vigencia_construccion=project_meta.vigencia_construccion,
        vigencia_mantenimiento=project_meta.vigencia_mantenimiento,
        ubicacion=project_meta.ubicacion,
        descripcion_proyecto=project_meta.descripcion_proyecto,
        valor_asegurado=project_meta.valor_asegurado,
    )


def _build_observations(quotes: list[ExtractedQuote]) -> list[str]:
    """Collect notable observations across all quotes.

    Args:
        quotes: List of extracted quotes.

    Returns:
        List of observation strings for the annexes.
    """
    observations: list[str] = []

    # Check for quotes with very different insured values
    values = [
        (q.nombre_compania, q.coberturas.get("danos_materiales"))
        for q in quotes
        if q.coberturas.get("danos_materiales") is not None
        and isinstance(q.coberturas.get("danos_materiales"), (int, float))
    ]
    if len(values) >= 2:
        nums = [v[1] for v in values]
        if isinstance(nums[0], (int, float)) and isinstance(nums[-1], (int, float)):
            if max(nums) > 0 and min(nums) / max(nums) < 0.9:
                observations.append(
                    "Se observa diferencia en el valor asegurado de daños materiales "
                    "entre aseguradoras. Verificar que todas coticen sobre el mismo valor."
                )

    return observations


def consolidate_quotes(
    quotes: list[ExtractedQuote],
    project_meta: ProjectMeta | None = None,
    tomador: str = "",
    recommendation: Recommendation | None = None,
    worst_markers: dict[str, list[str]] | None = None,
) -> ConsolidatedData:
    """Consolidate multiple extracted quotes into a single data structure.

    This is the REDUCE step of the Map-Reduce pattern.

    Args:
        quotes: List of individually extracted quotes (Map outputs).
        project_meta: AI-extracted project metadata.
        tomador: TOMADOR field value (empty = to be filled later).
        recommendation: AI-generated recommendation (optional).
        worst_markers: AI-generated worst-offering markers (optional).

    Returns:
        ConsolidatedData ready for the Word renderer.
    """
    logger.info("Consolidando %d cotizaciones", len(quotes))

    # Build company data entries
    aseguradoras = [_build_company_data(q) for q in quotes]

    # Build metadata
    meta = _build_meta(project_meta, tomador)

    # Build observations
    observations = _build_observations(quotes)

    # Build annex config
    anexos = AnnexConfig(
        leg=True,
        tasa_prorroga=True,
        observaciones_adicionales=observations,
    )

    consolidated = ConsolidatedData(
        meta=meta,
        aseguradoras=aseguradoras,
        marcadores_peor=worst_markers or {},
        recomendacion=recommendation,
        anexos=anexos,
    )

    logger.info(
        "Consolidación completada: %d aseguradoras, meta.tipo_cobertura=%s",
        len(aseguradoras),
        meta.tipo_cobertura,
    )

    return consolidated
