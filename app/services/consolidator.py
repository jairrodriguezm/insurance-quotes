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
from app.services.concept_mapper import ConceptMapper

logger = logging.getLogger(__name__)

_EMPTY_MARKERS = {"", "none", "null", "no especifica"}


def _normalize_insurer_id(raw_id: str, nombre: str = "") -> str:
    """Normalize insurer ID so multiple quotes from the same company share a canonical key."""
    s = (raw_id or "").strip().lower()
    nom = (nombre or "").strip().lower()
    canonical_aliases = {
        "seguros_del_estado": "estado",
        "seguros del estado": "estado",
        "estado": "estado",
        "zurich": "zurich",
        "axa_colpatria": "axa_colpatria",
        "axa colpatria": "axa_colpatria",
        "colpatria": "axa_colpatria",
        "mapfre": "mapfre",
        "sura": "sura",
        "suramericana": "sura",
        "bolivar": "bolivar",
        "seguros_bolivar": "bolivar",
        "chubb": "chubb",
        "hdi": "hdi",
        "mundial": "mundial",
        "allianz": "allianz",
        "solidaria": "solidaria",
        "previsora": "previsora",
        "equidad": "equidad",
        "berkley": "berkley",
    }
    for k, v in canonical_aliases.items():
        if k == s or k in s or k in nom:
            return v
    return s or "aseguradora"


def _build_company_data(quote: ExtractedQuote) -> InsuranceCompanyData:
    """Convert an extracted quote into the company data format for the renderer.

    Args:
        quote: Individual extracted quote.

    Returns:
        InsuranceCompanyData ready for the consolidated structure.
    """
    norm_id = _normalize_insurer_id(quote.id_compania, quote.nombre_compania)

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
        id=norm_id,
        nombre=quote.nombre_compania,
        logo=f"{norm_id}.png",
        coaseguro=coaseguro,
        coaseguro_nota="",
        validez_oferta=quote.validez_oferta,
        opciones=opciones,
        tasa_prorroga=quote.tasa_prorroga,
        coberturas=dict(quote.coberturas or {}),
        deducibles=dict(quote.deducibles or {}),
        asistencias=dict(quote.asistencias or {}),
        subjetividades=list(quote.subjetividades or []),
    )


def _merge_same_insurer_companies(
    companies: list[InsuranceCompanyData],
) -> list[InsuranceCompanyData]:
    """Merge multiple InsuranceCompanyData entries belonging to the same insurer.

    In Copropiedades, brokers frequently upload separate PDFs for Option 1 and
    Option 2 of the same insurer (e.g., Zurich VA vs Zurich Valor de la Pérdida).
    This merges them into a single column with ordered options ('Opción 1', 'Opción 2')
    and combines differing deductibles/coverages cleanly.
    """
    grouped: dict[str, list[InsuranceCompanyData]] = {}
    order: list[str] = []
    for comp in companies:
        cid = comp.id.lower()
        if cid not in grouped:
            grouped[cid] = []
            order.append(cid)
        grouped[cid].append(comp)

    merged_list: list[InsuranceCompanyData] = []
    for cid in order:
        group = grouped[cid]
        if len(group) == 1:
            single = group[0]
            if len(single.opciones) > 1:
                single.opciones.sort(key=lambda op: op.get("prima") or 999999999)
                for idx, op in enumerate(single.opciones, start=1):
                    etq = str(op.get("etiqueta") or "").strip()
                    if not etq or etq.lower() in {"opción 1", "opcion 1", "opción 2", "opcion 2"}:
                        op["etiqueta"] = f"Opción {idx}"
            merged_list.append(single)
            continue

        # Sort sub-quotes by lowest option premium so Opción 1 is lowest premium, Opción 2 next
        group.sort(
            key=lambda c: min((op.get("prima") or 999999999 for op in c.opciones), default=999999999)
        )
        base = group[0].model_copy(deep=True)

        combined_opts: list[dict[str, Any]] = []
        for comp in group:
            for op in comp.opciones:
                combined_opts.append(dict(op))
        combined_opts.sort(key=lambda op: op.get("prima") or 999999999)
        for idx, op in enumerate(combined_opts, start=1):
            op["etiqueta"] = f"Opción {idx}"
        base.opciones = combined_opts

        # Merge coberturas, deducibles, asistencias across sub-quotes
        for field_name in ("coberturas", "deducibles", "asistencias"):
            all_keys: list[str] = []
            for comp in group:
                for k in getattr(comp, field_name).keys():
                    if k not in all_keys:
                        all_keys.append(k)

            merged_dict: dict[str, Any] = {}
            for k in all_keys:
                vals = [getattr(comp, field_name).get(k) for comp in group]
                non_empty = [
                    v for v in vals
                    if v is not None and str(v).strip().lower() not in _EMPTY_MARKERS
                ]
                if not non_empty:
                    merged_dict[k] = vals[0]
                else:
                    norm_unique = {str(v).strip() for v in non_empty}
                    if len(norm_unique) == 1:
                        merged_dict[k] = non_empty[0]
                    elif field_name == "deducibles":
                        parts = []
                        for idx, v in enumerate(vals, start=1):
                            if v is not None and str(v).strip().lower() not in _EMPTY_MARKERS:
                                parts.append(f"Opción {idx}: {str(v).strip()}")
                        merged_dict[k] = " / ".join(parts)
                    else:
                        # Prefer numeric or most complete string value
                        merged_dict[k] = non_empty[0]
            setattr(base, field_name, merged_dict)

        # Merge subjetividades preserving order without duplicates
        seen_subj: set[str] = set()
        merged_subj: list[str] = []
        for comp in group:
            for s in comp.subjetividades:
                if s and s not in seen_subj:
                    seen_subj.add(s)
                    merged_subj.append(s)
        base.subjetividades = merged_subj
        merged_list.append(base)

    return merged_list


def _build_meta(
    project_meta: ProjectMeta | None,
    tomador: str = "",
    categoria: str = "",
) -> MetaInfo:
    """Build the MetaInfo from extracted project metadata and overrides.

    Args:
        project_meta: AI-extracted project metadata (may be None).
        tomador: Override for the TOMADOR field (empty = to be filled later).
        categoria: Selected insurance category.

    Returns:
        MetaInfo for the consolidated document.
    """
    effective_tomador = tomador.strip() or (project_meta.tomador if project_meta else "")

    if project_meta is None:
        return MetaInfo(
            fecha=date.today().strftime("%d/%m/%Y"),
            categoria=categoria,
            tomador=effective_tomador,
        )

    return MetaInfo(
        fecha=date.today().strftime("%d/%m/%Y"),
        tipo_cobertura=project_meta.tipo_cobertura,
        categoria=categoria or project_meta.tipo_cobertura,
        tomador=effective_tomador,
        asegurado=project_meta.asegurado,
        beneficiario=project_meta.beneficiario,
        vigencia_construccion=project_meta.vigencia_construccion,
        vigencia_mantenimiento=project_meta.vigencia_mantenimiento,
        ubicacion=project_meta.ubicacion,
        descripcion_proyecto=project_meta.descripcion_proyecto,
        valor_asegurado=project_meta.valor_asegurado,
        # Autos
        identificacion=project_meta.identificacion or project_meta.cedula,
        marca=project_meta.marca,
        placa=project_meta.placa,
        linea=project_meta.linea,
        modelo=project_meta.modelo,
        servicio=project_meta.servicio,
        zona_circulacion=project_meta.zona_circulacion,
        accesorios=project_meta.accesorios,
        # Hogar
        cedula=project_meta.cedula or project_meta.identificacion,
        direccion=project_meta.direccion or project_meta.ubicacion,
        ciudad=project_meta.ciudad,
        ano_construccion=project_meta.ano_construccion,
        valor_edificio=project_meta.valor_edificio,
        valor_muebles=project_meta.valor_muebles,
        valor_equipos=project_meta.valor_equipos,
        valor_arte=project_meta.valor_arte,
        valor_dinero=project_meta.valor_dinero,
        asegurado_actualmente=project_meta.asegurado_actualmente,
        siniestros_previos=project_meta.siniestros_previos,
        # Copropiedades
        valor_cimentacion=project_meta.valor_cimentacion,
        valor_maquinaria=project_meta.valor_maquinaria,
        valor_equipos_moviles=project_meta.valor_equipos_moviles,
        valor_rce=project_meta.valor_rce,
        valor_dno=project_meta.valor_dno,
        valor_manejo=project_meta.valor_manejo,
        numero_torres=project_meta.numero_torres,
        numero_apartamentos=project_meta.numero_apartamentos,
        numero_pisos=project_meta.numero_pisos,
        uso=project_meta.uso,
        tipo_construccion=project_meta.tipo_construccion,
        inf_copropiedad=project_meta.inf_copropiedad,
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
    categoria: str = "",
) -> ConsolidatedData:
    """Consolidate multiple extracted quotes into a single data structure.

    This is the REDUCE step of the Map-Reduce pattern.

    Args:
        quotes: List of individually extracted quotes (Map outputs).
        project_meta: AI-extracted project metadata.
        tomador: TOMADOR field value (empty = to be filled later).
        recommendation: AI-generated recommendation (optional).
        worst_markers: AI-generated worst-offering markers (optional).
        categoria: Selected insurance category.

    Returns:
        ConsolidatedData ready for the Word renderer.
    """
    logger.info("Consolidando %d cotizaciones (categoría: %s)", len(quotes), categoria)
    norm_cat = ConceptMapper.normalize_category(
        categoria or (project_meta.tipo_cobertura if project_meta else "")
    )

    # Build company data entries
    aseguradoras = [_build_company_data(q) for q in quotes]

    if norm_cat == "Copropiedades":
        # Merge multiple quote PDFs belonging to the same insurer (e.g. Option 1 & Option 2)
        aseguradoras = _merge_same_insurer_companies(aseguradoras)
        # In Copropiedades, rows in Section III and columns in V/VI/VII are ordered strictly
        # by ascending lowest option premium (e.g. Estado -> Zurich -> AXA -> Mapfre).
        aseguradoras.sort(
            key=lambda a: min(
                (op.get("prima") or 999999999 for op in a.opciones),
                default=999999999,
            )
        )
    else:
        # Sort insurers: recommended option first, then by ascending premium (best offer first)
        if recommendation and recommendation.aseguradora_id:
            rec_id = recommendation.aseguradora_id.lower()
            aseguradoras.sort(key=lambda a: (
                0 if (rec_id in a.id.lower() or a.id.lower() in rec_id) else 1,
                (a.opciones[0].get("prima") or 999999999) if a.opciones else 999999999,
            ))
        else:
            aseguradoras.sort(
                key=lambda a: (a.opciones[0].get("prima") or 999999999) if a.opciones else 999999999
            )

    # Build metadata
    meta = _build_meta(project_meta, tomador=tomador, categoria=categoria)

    # Build observations
    observations = _build_observations(quotes)
    if recommendation and getattr(recommendation, "observaciones_tecnicas", None):
        observations.extend(recommendation.observaciones_tecnicas)

    # Build annex config (Copropiedades ends at VIII. PREVALENCIA without TRC annexes)
    is_coprop = norm_cat == "Copropiedades"
    anexos = AnnexConfig(
        leg=not is_coprop,
        tasa_prorroga=not is_coprop,
        observaciones_adicionales=[] if is_coprop else observations,
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
