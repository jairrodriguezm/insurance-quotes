# -*- coding: utf-8 -*-
"""Parser and deterministic analytics engine for Flokzu Management Reports (.xlsx).

Reads raw Flokzu Excel exports (e.g. multiriesgos-CUMP-YYYY-MM-DD.xlsx),
extracts periods and clients, and computes 100% deterministic metrics,
tables, and narratives matching the exact reference format.
"""
from __future__ import annotations

import calendar
import io
import logging
import re
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime
from typing import Any

import openpyxl

from app.schemas.informe_gestion import (
    ClienteSuggestion,
    InformeInspectResponse,
    PeriodoOption,
)

logger = logging.getLogger(__name__)

# Spanish month names
MESES_ES = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}

RAMO_NAMES = {
    "CUMP": "Seguros de cumplimiento",
}

# Known corporate client profiles with exact canonical names, NITs, and aliases
KNOWN_CLIENTS = [
    {
        "id": "SONACOL",
        "nombre_corto": "SONACOL",
        "nombre_completo": "SONACOL – Solarte Nacional de Construcciones S.A.S.",
        "nit": "830.129.289-8",
        "terminos_busqueda": [
            "sonacol",
            "solarte nacional de construcciones",
        ],
        "flag_col": "Sonacol? (CUMP)",
    },
    {
        "id": "COMPENSAR",
        "nombre_corto": "Compensar",
        "nombre_completo": "Caja de Compensación Familiar Compensar",
        "nit": "860.066.942-7",
        "terminos_busqueda": [
            "compensar",
            "caja de compensacion familiar compensar",
        ],
        "flag_col": None,
    },
    {
        "id": "OCOBO",
        "nombre_corto": "Ocobo Construcciones",
        "nombre_completo": "Ocobo Construcciones S.A.S.",
        "nit": "",
        "terminos_busqueda": [
            "ocobo construcciones",
        ],
        "flag_col": None,
    },
]

# Canonical ordering of request types in Trazabilidad and Tiempos tables
CANONICAL_TIPO_ORDER = [
    "Cumplimiento con RC",
    "Cumplimiento",
    "Modificación",
    "Seriedad de oferta",
    "Judiciales",
]

# Uppercase acronyms and tokens to preserve when formatting entity names
UPPERCASE_ACRONYMS = {
    "SAS", "S.A.S", "S.A.S.", "SA", "S.A", "S.A.", "LTDA", "LTDA.",
    "BIC", "ESP", "E.S.P", "E.S.P.", "UT", "U.T.", "IDU", "INVIAS",
    "LHS", "SNC", "SENA", "ANI", "ICCU", "FFIE", "SECOP", "KVA", "RH",
    "SICS", "RC", "RCE", "D.C.", "DC", "IPS", "EPS", "ARL", "ESE",
    "E.S.E.", "CIA", "CÍA", "E", "&", "L", "V5", "P73", "5D", "3M",
}

LOWERCASE_CONNECTORS = {
    "de", "del", "la", "las", "el", "los", "y", "e", "o", "u", "en", "para", "por", "con", "a",
}


def _strip_accents(text: str) -> str:
    """Lowercase and strip accents for case/accent-insensitive matching."""
    if not text:
        return ""
    norm = unicodedata.normalize("NFD", str(text))
    return "".join(c for c in norm if unicodedata.category(c) != "Mn").lower().strip()


def format_colombian_nit(raw_nit: str) -> str:
    """Format a raw numeric NIT into standard Colombian format XXX.XXX.XXX-Y."""
    if not raw_nit:
        return ""
    raw = str(raw_nit).strip()
    if "." in raw and "-" in raw:
        return raw
    digits = re.sub(r"\D", "", raw)
    if not digits:
        return raw
    if len(digits) == 10 and digits[0] in ("8", "9"):
        base, dv = digits[:9], digits[9]
        return f"{base[:3]}.{base[3:6]}.{base[6:9]}-{dv}"
    if len(digits) == 9 and digits[0] in ("8", "9"):
        weights = [41, 37, 29, 23, 19, 17, 13, 7, 3]
        total = sum(int(d) * w for d, w in zip(digits, weights))
        rem = total % 11
        dv = str(rem) if rem in (0, 1) else str(11 - rem)
        return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{dv}"
    return raw


def format_entity_name(raw_name: str) -> str:
    """Format an uppercase entity name into corporate title case matching reference docx."""
    if not raw_name:
        return ""
    cleaned = " ".join(str(raw_name).strip().split())
    # Collapse trailing "S A S" into "S.A.S"
    cleaned = re.sub(r"\bS\s+A\s+S\.?$", "S.A.S", cleaned, flags=re.IGNORECASE)

    words = cleaned.split(" ")
    result: list[str] = []
    for idx, w in enumerate(words):
        w_clean = w.strip()
        w_upper = w_clean.upper()
        w_no_acc = _strip_accents(w_clean)

        if w_upper in UPPERCASE_ACRONYMS:
            result.append(w_upper)
        elif idx > 0 and w_no_acc in LOWERCASE_CONNECTORS:
            result.append(w_clean.lower())
        else:
            result.append(w_clean.capitalize())

    return " ".join(result)


def format_integrantes(raw_integrantes: str) -> str:
    """Format consortium members separated by newlines into 'Integrantes: A; B'."""
    if not raw_integrantes:
        return ""
    parts = [
        p.strip()
        for p in re.split(r"[\r\n;]+", str(raw_integrantes))
        if p and p.strip()
    ]
    if not parts:
        return ""
    formatted_parts = [format_entity_name(p) for p in parts]
    return "Integrantes: " + "; ".join(formatted_parts)


def format_spanish_date_long(dt_str: str) -> str:
    """Format 'YYYY-MM-DD' into '22 de octubre de 2026'."""
    try:
        dt = datetime.strptime(dt_str[:10], "%Y-%m-%d")
        mes = MESES_ES.get(dt.month, "")
        return f"{dt.day} de {mes} de {dt.year}"
    except Exception:
        return dt_str


def format_date_ddmmyyyy(dt_str: str) -> str:
    """Format 'YYYY-MM-DD...' into 'DD/MM/YYYY' matching Tables 4 and 8."""
    try:
        dt = datetime.strptime(dt_str[:10], "%Y-%m-%d")
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return dt_str[:10]


def format_decimal_es(val: float, decimals: int = 1) -> str:
    """Format float with comma as decimal separator (e.g. 11.2 -> '11,2')."""
    return f"{val:.{decimals}f}".replace(".", ",")


def _load_flokzu_rows(file_bytes: bytes) -> tuple[list[Any], list[str], list[tuple[Any, ...]]]:
    """Load Flokzu Excel bytes and return (meta_row, headers, data_rows)."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    if len(all_rows) < 2:
        raise ValueError("El archivo Excel no tiene el formato esperado de exportación Flokzu.")
    meta_row = list(all_rows[0])
    headers = [str(h or "").strip() for h in all_rows[1]]
    data_rows = [r for r in all_rows[2:] if r and any(cell is not None and str(cell).strip() != "" for cell in r)]
    return meta_row, headers, data_rows


def _resolve_col_indices(headers: list[str]) -> dict[str, int]:
    """Map logical field names to column indices by header name with safe fallbacks."""
    h_map: dict[str, int] = {}
    for idx, h in enumerate(headers):
        h_norm = _strip_accents(h)
        if h_norm not in h_map:
            h_map[h_norm] = idx

    def find_col(candidates: list[str], default_idx: int) -> int:
        for cand in candidates:
            c_norm = _strip_accents(cand)
            if c_norm in h_map:
                return h_map[c_norm]
        for cand in candidates:
            c_norm = _strip_accents(cand)
            for k, idx in h_map.items():
                if c_norm in k:
                    return idx
        return default_idx

    return {
        "id": find_col(["Identificador"], 0),
        "fecha_inicio": find_col(["Fecha Inicio"], 1),
        "finalizado": find_col(["Finalizado"], 3),
        "fecha_fin": find_col(["Fecha Finalización", "Fecha Finalizacion"], 4),
        "tiempo_vida": find_col(["Tiempo de vida"], 5),
        "tarea_actual": find_col(["Tarea Actual"], 8),
        "pagare_cump": find_col(["Cliente debe firmar pagaré Cumplimiento? (CUMP)"], 86),
        "pagare": find_col(["Cliente debe firmar pagaré? (CUMP)"], 87),
        "cartera": find_col(["Es necesario estado de Cartera del Cliente? (CUMP)"], 131),
        "nit_cliente": find_col(["Identificación del Cliente (CUMP)"], 158),
        "info_adicional": find_col(["Información adicional solicitada al Cliente (CUMP)"], 159),
        "integrantes": find_col(["Integrantes del Consorcio o Unión Temporal (CUMP)"], 162),
        "entidad": find_col(["Nombre de la Entidad (CUMP)"], 164),
        "remitente": find_col(["Nombre del remitente (CUMP)"], 165),
        "objecion": find_col(["Objeción presentada por Aseguradoras (CUMP)"], 176),
        "sonacol_flag": find_col(["Sonacol? (CUMP)"], 194),
        "subtipo": find_col(["Subtipo (CUMP)"], 199),
        "tipo_solicitud": find_col(["Tipo de Solicitud (CUMP)"], 205),
    }


def _extract_meta_info(meta_row: list[Any], data_rows: list[tuple[Any, ...]], cols: dict[str, int]) -> tuple[str, str, str]:
    """Extract (ramo_codigo, ramo_nombre, fecha_corte) from row 0 or data rows."""
    fecha_corte = ""
    ramo_codigo = "CUMP"
    if len(meta_row) > 3 and meta_row[3]:
        raw_corte = str(meta_row[3]).strip()
        if re.match(r"^\d{4}-\d{2}-\d{2}", raw_corte):
            fecha_corte = raw_corte[:10]
    if len(meta_row) > 5 and meta_row[5]:
        ramo_codigo = str(meta_row[5]).strip().upper()

    if not fecha_corte:
        dates: list[str] = []
        for r in data_rows:
            for c_key in ("fecha_inicio", "fecha_fin"):
                idx = cols[c_key]
                if idx < len(r) and r[idx]:
                    s = str(r[idx]).strip()[:10]
                    if re.match(r"^\d{4}-\d{2}-\d{2}$", s):
                        dates.append(s)
        fecha_corte = max(dates) if dates else datetime.now().strftime("%Y-%m-%d")

    ramo_nombre = RAMO_NAMES.get(ramo_codigo, f"Seguros ({ramo_codigo})")
    return ramo_codigo, ramo_nombre, fecha_corte


def inspect_flokzu_xlsx(file_bytes: bytes) -> InformeInspectResponse:
    """Inspect a Flokzu .xlsx file and return available periods and suggested clients."""
    meta_row, headers, data_rows = _load_flokzu_rows(file_bytes)
    cols = _resolve_col_indices(headers)
    ramo_codigo, ramo_nombre, fecha_corte = _extract_meta_info(meta_row, data_rows, cols)

    periodo_counts: Counter[str] = Counter()
    for r in data_rows:
        fi = str(r[cols["fecha_inicio"]] or "").strip()
        if len(fi) >= 7 and re.match(r"^\d{4}-\d{2}", fi[:7]):
            periodo_counts[fi[:7]] += 1

    periodos: list[PeriodoOption] = []
    for per in sorted(periodo_counts.keys(), reverse=True):
        yr, mo = int(per[:4]), int(per[5:7])
        mes_label = MESES_ES.get(mo, per).capitalize()
        periodos.append(
            PeriodoOption(
                periodo=per,
                label=f"{mes_label} de {yr}",
                count=periodo_counts[per],
            )
        )

    remitente_nits: dict[str, Counter[str]] = defaultdict(Counter)
    for r in data_rows:
        rem = _strip_accents(str(r[cols["remitente"]] or "")).upper()
        nit = str(r[cols["nit_cliente"]] or "").strip()
        if rem and nit:
            remitente_nits[rem][nit] += 1

    clientes_sugeridos: list[ClienteSuggestion] = []
    seen_terms: set[str] = set()

    for kc in KNOWN_CLIENTS:
        terms_norm = [_strip_accents(t) for t in kc["terminos_busqueda"]]
        per_counts: Counter[str] = Counter()
        total_c = 0
        for r in data_rows:
            s_ent = _strip_accents(str(r[cols["entidad"]] or ""))
            s_rem = _strip_accents(str(r[cols["remitente"]] or ""))
            s_int = _strip_accents(str(r[cols["integrantes"]] or ""))
            flag_match = (
                kc.get("flag_col")
                and cols.get("sonacol_flag") is not None
                and str(r[cols["sonacol_flag"]] or "").strip().lower() == "true"
            )
            text_match = any(t in s_ent or t in s_rem or t in s_int for t in terms_norm)
            if flag_match or text_match:
                total_c += 1
                fi = str(r[cols["fecha_inicio"]] or "").strip()[:7]
                if re.match(r"^\d{4}-\d{2}$", fi):
                    per_counts[fi] += 1

        if total_c > 0:
            for t in terms_norm:
                seen_terms.add(t)
            clientes_sugeridos.append(
                ClienteSuggestion(
                    id=kc["id"],
                    nombre_corto=kc["nombre_corto"],
                    nombre_completo=kc["nombre_completo"],
                    nit=kc["nit"],
                    terminos_busqueda=kc["terminos_busqueda"],
                    total_solicitudes=total_c,
                    solicitudes_por_periodo=dict(per_counts),
                )
            )

    entity_periods: dict[str, Counter[str]] = defaultdict(Counter)
    entity_raw_display: dict[str, str] = {}
    for r in data_rows:
        fi = str(r[cols["fecha_inicio"]] or "").strip()[:7]
        for col_key in ("entidad", "remitente"):
            raw_val = str(r[cols[col_key]] or "").strip()
            if not raw_val:
                continue
            norm_val = _strip_accents(raw_val).rstrip(".")
            if not norm_val or any(t in norm_val for t in seen_terms):
                continue
            entity_periods[norm_val][fi] += 1
            if norm_val not in entity_raw_display:
                entity_raw_display[norm_val] = raw_val

    top_others = sorted(
        entity_periods.items(),
        key=lambda kv: sum(kv[1].values()),
        reverse=True,
    )[:25]

    for norm_val, per_c in top_others:
        total_c = sum(per_c.values())
        if total_c < 3:
            continue
        raw_disp = entity_raw_display[norm_val]
        formatted_name = format_entity_name(raw_disp)
        upper_norm = norm_val.upper()
        raw_nit = ""
        if upper_norm in remitente_nits and remitente_nits[upper_norm]:
            raw_nit = remitente_nits[upper_norm].most_common(1)[0][0]
        clientes_sugeridos.append(
            ClienteSuggestion(
                id=upper_norm[:40],
                nombre_corto=formatted_name,
                nombre_completo=formatted_name,
                nit=format_colombian_nit(raw_nit) if raw_nit else "",
                terminos_busqueda=[norm_val],
                total_solicitudes=total_c,
                solicitudes_por_periodo=dict(per_c),
            )
        )

    return InformeInspectResponse(
        ramo=ramo_nombre,
        ramo_codigo=ramo_codigo,
        fecha_corte=fecha_corte,
        total_filas=len(data_rows),
        periodos=periodos,
        clientes_sugeridos=clientes_sugeridos,
    )


def _map_pending_motivo(row: tuple[Any, ...], cols: dict[str, int]) -> str:
    """Map a pending request's current task and flags to a human-readable Spanish reason."""
    tarea = str(row[cols["tarea_actual"]] or "").strip()
    tarea_norm = _strip_accents(tarea)

    reasons: list[str] = []
    if tarea_norm == "confirmar borrador":
        reasons.append("En espera de confirmación del borrador por el cliente")
    elif tarea_norm == "esperar respuesta del cliente":
        reasons.append("En espera de respuesta del cliente")
    elif tarea_norm == "esperar soporte de pago":
        reasons.append("En espera de soporte de pago por el cliente")
    elif tarea_norm == "espera respuesta cia de seguros":
        reasons.append("En espera de respuesta de la compañía de seguros")
    elif "adjudicacion de so" in tarea_norm:
        reasons.append("En espera hasta fecha de adjudicación de seriedad de oferta")
    elif "verificar en secop" in tarea_norm:
        reasons.append("En verificación de adjudicación del proceso en SECOP")
    elif "cargue de poliza en sics" in tarea_norm:
        reasons.append("En proceso de cargue de póliza en SICS")
    elif "emitir y enviar poliza" in tarea_norm:
        reasons.append("En proceso de emisión y envío de póliza")
    elif tarea:
        reasons.append(tarea)

    pagare_1 = str(row[cols["pagare_cump"]] or "").strip().lower() == "true"
    pagare_2 = str(row[cols["pagare"]] or "").strip().lower() == "true"
    cartera = str(row[cols["cartera"]] or "").strip().lower() == "true"

    if pagare_1 or pagare_2:
        reasons.append("Pendiente de firma de pagaré por el cliente")
    if cartera:
        reasons.append("Pendiente de estado de cartera del cliente")

    if not reasons:
        return "En trámite de gestión técnica"
    return " · ".join(reasons)


def _map_declined_motivo(row: tuple[Any, ...], cols: dict[str, int]) -> str:
    """Map a declined request to a clear reason based on available columns."""
    objecion = str(row[cols["objecion"]] or "").strip()
    if objecion and objecion.lower() not in ("false", "n/a", "na"):
        return objecion
    info_adic = str(row[cols["info_adicional"]] or "").strip()
    if info_adic:
        info_norm = _strip_accents(info_adic)
        if "contrato firmado" in info_norm:
            return "Pendiente envío de contrato firmado / documentación del cliente"
        if "estados financieros" in info_norm or "sarlaft" in info_norm:
            return "Documentación financiera y SARLAFT pendiente por el cliente"
        return "Sin respuesta o confirmación de condiciones por el cliente"
    return "Desistimiento o solicitud declinada en validación técnica"


def _parse_hours(tiempo_vida: Any, fecha_inicio_str: str, fecha_fin_str: str) -> int | None:
    """Extract integer elapsed hours from 'Tiempo de vida' (e.g. '357 Hs.') or dates."""
    if tiempo_vida is not None:
        m = re.search(r"(\d+)", str(tiempo_vida))
        if m:
            return int(m.group(1))
    if fecha_inicio_str and fecha_fin_str:
        try:
            dt1 = datetime.strptime(fecha_inicio_str[:19], "%Y-%m-%d %H:%M:%S")
            dt2 = datetime.strptime(fecha_fin_str[:19], "%Y-%m-%d %H:%M:%S")
            diff_hrs = int(round((dt2 - dt1).total_seconds() / 3600.0))
            return max(0, diff_hrs)
        except Exception:
            pass
    return None


def process_flokzu_xlsx(
    file_bytes: bytes,
    periodo: str,
    cliente_nombre_corto: str = "SONACOL",
    cliente_nombre_completo: str = "SONACOL – Solarte Nacional de Construcciones S.A.S.",
    cliente_nit: str = "830.129.289-8",
    terminos_busqueda: list[str] | None = None,
) -> dict[str, Any]:
    """Process the Flokzu .xlsx and compute all structured data for the management report."""
    meta_row, headers, data_rows = _load_flokzu_rows(file_bytes)
    cols = _resolve_col_indices(headers)
    ramo_codigo, ramo_nombre, fecha_corte = _extract_meta_info(meta_row, data_rows, cols)

    nombre_corto = (cliente_nombre_corto or "").strip() or "SONACOL"
    nombre_completo = (cliente_nombre_completo or "").strip() or nombre_corto
    nit_formatted = format_colombian_nit(cliente_nit)

    if not terminos_busqueda:
        if "sonacol" in _strip_accents(nombre_corto) or "solarte" in _strip_accents(nombre_completo):
            terminos_busqueda = ["sonacol", "solarte nacional de construcciones"]
        else:
            terminos_busqueda = [nombre_corto, nombre_completo]

    terms_norm = [
        _strip_accents(t).rstrip(".")
        for t in terminos_busqueda
        if t and _strip_accents(t).rstrip(".")
    ]
    is_sonacol = any("sonacol" in t or "solarte nacional" in t for t in terms_norm)
    is_all_clients = any(t in ("todos", "all", "*") for t in terms_norm)

    periodo_clean = (periodo or "").strip()
    if not periodo_clean or periodo_clean.upper() == "AUTO":
        inspect_res = inspect_flokzu_xlsx(file_bytes)
        periodo_clean = inspect_res.periodos[0].periodo if inspect_res.periodos else "2026-09"

    if re.match(r"^\d{4}-\d{2}$", periodo_clean):
        anio = int(periodo_clean[:4])
        mes_num = int(periodo_clean[5:7])
        mes_nombre = MESES_ES.get(mes_num, periodo_clean)
        ultimo_dia = calendar.monthrange(anio, mes_num)[1]
        periodo_titulo = f"{mes_nombre.capitalize()} de {anio}"
        periodo_narrativa = f"En {mes_nombre} de {anio}"
        rango_fechas_texto = f"entre el 1 y el {ultimo_dia} de {mes_nombre} de {anio}"
        universo_periodo_texto = f"recibidas en {mes_nombre} de {anio}"
    else:
        anio = datetime.now().year
        mes_nombre = "el periodo seleccionado"
        periodo_titulo = "Consolidado General"
        periodo_narrativa = "En el periodo analizado"
        rango_fechas_texto = "en el periodo seleccionado"
        universo_periodo_texto = "recibidas en el periodo analizado"

    matched_records: list[dict[str, Any]] = []
    count_asegurado = 0
    count_consorcio = 0
    count_tomador = 0
    pending_pagare_cartera_count = 0

    for r in data_rows:
        fi_raw = str(r[cols["fecha_inicio"]] or "").strip()
        if not fi_raw:
            continue
        if re.match(r"^\d{4}-\d{2}$", periodo_clean) and not fi_raw.startswith(periodo_clean):
            continue

        raw_ent = str(r[cols["entidad"]] or "").strip()
        raw_rem = str(r[cols["remitente"]] or "").strip()
        raw_int = str(r[cols["integrantes"]] or "").strip()

        s_ent = _strip_accents(raw_ent)
        s_rem = _strip_accents(raw_rem)
        s_int = _strip_accents(raw_int)
        flag_sonacol = (
            is_sonacol
            and cols.get("sonacol_flag") is not None
            and str(r[cols["sonacol_flag"]] or "").strip().lower() == "true"
        )

        in_rem = any(t in s_rem for t in terms_norm) if not is_all_clients else True
        in_int = any(t in s_int for t in terms_norm) if not is_all_clients else False
        in_ent = any(t in s_ent for t in terms_norm) if not is_all_clients else False

        if not is_all_clients and not (in_rem or in_int or in_ent or flag_sonacol):
            continue

        if in_int:
            rol = "consorcio"
            count_consorcio += 1
        elif in_rem and not in_ent:
            rol = "tomador"
            count_tomador += 1
        else:
            rol = "asegurado"
            count_asegurado += 1

        fin_raw = str(r[cols["finalizado"]] or "").strip()
        fin_norm = _strip_accents(fin_raw)
        if "emitida" in fin_norm:
            estado_cat = "emitida"
            estado_label = "Emitida"
        elif "novedad atendida" in fin_norm:
            estado_cat = "novedad"
            estado_label = "Atendida"
        elif "declinada" in fin_norm:
            estado_cat = "declinada"
            estado_label = "Declinada"
        else:
            estado_cat = "pendiente"
            estado_label = "Pendiente"
            p1 = str(r[cols["pagare_cump"]] or "").strip().lower() == "true"
            p2 = str(r[cols["pagare"]] or "").strip().lower() == "true"
            c1 = str(r[cols["cartera"]] or "").strip().lower() == "true"
            if p1 or p2 or c1:
                pending_pagare_cartera_count += 1

        ff_raw = str(r[cols["fecha_fin"]] or "").strip()
        horas: int | None = None
        dias: float | None = None
        if estado_cat != "pendiente":
            horas = _parse_hours(r[cols["tiempo_vida"]], fi_raw, ff_raw)
            if horas is not None:
                dias = horas / 24.0

        tipo_raw = str(r[cols["tipo_solicitud"]] or "").strip() or "Otro"
        subtipo_raw = str(r[cols["subtipo"]] or "").strip()
        if subtipo_raw and subtipo_raw.upper() not in ("N/A", "NA", "NINGUNO", ""):
            tipo_detalle = f"{tipo_raw} · {subtipo_raw}"
        else:
            tipo_detalle = tipo_raw

        tomador_fmt = format_entity_name(raw_rem) if raw_rem else "—"
        integrantes_fmt = format_integrantes(raw_int)
        if integrantes_fmt:
            tomador_tabla8 = f"{tomador_fmt}\n{integrantes_fmt}"
        else:
            tomador_tabla8 = tomador_fmt

        asegurado_fmt = format_entity_name(raw_ent) if raw_ent else "—"
        recibida_ddmmyyyy = format_date_ddmmyyyy(fi_raw)

        matched_records.append(
            {
                "id": str(r[cols["id"]] or "").strip(),
                "fecha_inicio_dt": fi_raw,
                "recibida": recibida_ddmmyyyy,
                "tipo": tipo_raw,
                "subtipo": subtipo_raw,
                "tipo_detalle": tipo_detalle,
                "tomador_fmt": tomador_fmt,
                "tomador_tabla8": tomador_tabla8,
                "asegurado_fmt": asegurado_fmt,
                "tomador_asegurado": f"{tomador_fmt} / {asegurado_fmt}",
                "estado_cat": estado_cat,
                "estado_label": estado_label,
                "horas": horas,
                "dias": dias,
                "dias_fmt": format_decimal_es(dias, 1) if dias is not None else "—",
                "motivo_pendiente": _map_pending_motivo(r, cols) if estado_cat == "pendiente" else "",
                "motivo_declinada": _map_declined_motivo(r, cols) if estado_cat == "declinada" else "",
                "rol": rol,
            }
        )

    matched_records.sort(key=lambda x: (x["fecha_inicio_dt"], x["id"]))

    total_tramitadas = len(matched_records)
    total_emitidas = sum(1 for x in matched_records if x["estado_cat"] == "emitida")
    total_novedades = sum(1 for x in matched_records if x["estado_cat"] == "novedad")
    total_declinadas = sum(1 for x in matched_records if x["estado_cat"] == "declinada")
    total_pendientes = sum(1 for x in matched_records if x["estado_cat"] == "pendiente")

    total_atendidas = total_emitidas + total_novedades
    total_cerradas = total_atendidas + total_declinadas

    efectividad_pct = (total_atendidas / total_tramitadas * 100.0) if total_tramitadas > 0 else 0.0
    gestion_pct = (total_cerradas / total_tramitadas * 100.0) if total_tramitadas > 0 else 0.0

    closed_with_time = [x for x in matched_records if x["estado_cat"] != "pendiente" and x["horas"] is not None]
    if closed_with_time:
        avg_hours = sum(x["horas"] for x in closed_with_time) / len(closed_with_time)
        avg_days = avg_hours / 24.0
        avg_hours_int = int(round(avg_hours))
        tiempo_promedio_kpi = f"{format_decimal_es(avg_days, 1)} días"
        tiempo_promedio_tabla5 = f"{format_decimal_es(avg_days, 1)} días ({avg_hours_int} horas)"

        sorted_days = sorted(x["dias"] for x in closed_with_time)
        n_c = len(sorted_days)
        if n_c % 2 == 1:
            med_days = sorted_days[n_c // 2]
        else:
            med_days = (sorted_days[n_c // 2 - 1] + sorted_days[n_c // 2]) / 2.0
        mediana_tabla5 = f"{format_decimal_es(med_days, 1)} días"

        fastest = min(closed_with_time, key=lambda x: x["horas"])
        slowest = max(closed_with_time, key=lambda x: x["horas"])
        extremos_tabla5 = f"{format_decimal_es(fastest['dias'], 1)} / {format_decimal_es(slowest['dias'], 1)} días"
        tiempo_promedio_detalle = (
            f"Mediana de {format_decimal_es(med_days, 1)} días; "
            f"máximo de {format_decimal_es(slowest['dias'], 1)} días "
            f"(sobre {len(closed_with_time)} {'solicitud cerrada' if len(closed_with_time) == 1 else 'solicitudes cerradas'})"
        )
    else:
        avg_days = 0.0
        tiempo_promedio_kpi = "—"
        tiempo_promedio_tabla5 = "—"
        mediana_tabla5 = "—"
        extremos_tabla5 = "—"
        tiempo_promedio_detalle = "Sin solicitudes cerradas en el período"

    fecha_corte_larga = format_spanish_date_long(fecha_corte)
    pendientes_phrase = (
        "1 queda pendiente"
        if total_pendientes == 1
        else f"{total_pendientes} quedan pendientes"
    )
    declinadas_phrase = (
        "ninguna fue declinada"
        if total_declinadas == 0
        else (
            "1 fue declinada"
            if total_declinadas == 1
            else f"{total_declinadas} fueron declinadas"
        )
    )
    resumen_narrativa = (
        f"{periodo_narrativa} se tramitaron {total_tramitadas} solicitudes relacionadas con {nombre_corto}: "
        f"{count_asegurado} como asegurado, {count_consorcio} como integrante de consorcio o unión temporal, "
        f"{count_tomador} como tomador. {total_atendidas} de ellas ({format_decimal_es(efectividad_pct, 1)} %) "
        f"ya fueron atendidas, {declinadas_phrase} y {pendientes_phrase}."
    )

    tipos_present = []
    for t in CANONICAL_TIPO_ORDER:
        if any(x["tipo"] == t for x in matched_records):
            tipos_present.append(t)
    for x in matched_records:
        if x["tipo"] not in tipos_present:
            tipos_present.append(x["tipo"])

    trazabilidad_rows: list[dict[str, str]] = []
    tiempos_por_tipo_rows: list[dict[str, str]] = []

    for t in tipos_present:
        t_recs = [x for x in matched_records if x["tipo"] == t]
        t_tram = len(t_recs)
        t_emit = sum(1 for x in t_recs if x["estado_cat"] == "emitida")
        t_nov = sum(1 for x in t_recs if x["estado_cat"] == "novedad")
        t_decl = sum(1 for x in t_recs if x["estado_cat"] == "declinada")
        t_pend = sum(1 for x in t_recs if x["estado_cat"] == "pendiente")
        t_efec = ((t_emit + t_nov) / t_tram * 100.0) if t_tram > 0 else 0.0

        t_closed = [x for x in t_recs if x["estado_cat"] != "pendiente" and x["horas"] is not None]
        if t_closed:
            t_avg_d = (sum(x["horas"] for x in t_closed) / len(t_closed)) / 24.0
            t_dias_str = format_decimal_es(t_avg_d, 1)
            tiempos_por_tipo_rows.append(
                {
                    "tipo": t,
                    "promedio": t_dias_str,
                }
            )
        else:
            t_dias_str = "—"

        trazabilidad_rows.append(
            {
                "tipo": t,
                "tramitadas": str(t_tram),
                "emitidas": str(t_emit),
                "novedades": str(t_nov),
                "declinadas": str(t_decl),
                "pendientes": str(t_pend),
                "efectividad": f"{format_decimal_es(t_efec, 1)} %",
                "dias_prom": t_dias_str,
            }
        )

    trazabilidad_total = {
        "tipo": "Total",
        "tramitadas": str(total_tramitadas),
        "emitidas": str(total_emitidas),
        "novedades": str(total_novedades),
        "declinadas": str(total_declinadas),
        "pendientes": str(total_pendientes),
        "efectividad": f"{format_decimal_es(efectividad_pct, 1)} %",
        "dias_prom": format_decimal_es(avg_days, 1) if closed_with_time else "—",
    }

    declinadas_list = [x for x in matched_records if x["estado_cat"] == "declinada"]
    declinadas_motivos_counter = Counter(x["motivo_declinada"] for x in declinadas_list)
    declinadas_resumen_rows = [
        {"motivo": m, "cantidad": str(c)}
        for m, c in declinadas_motivos_counter.most_common()
    ]

    pendientes_list = [x for x in matched_records if x["estado_cat"] == "pendiente"]
    pendientes_motivos_counter = Counter(x["motivo_pendiente"] for x in pendientes_list)
    pendientes_resumen_rows = [
        {"motivo": m, "cantidad": str(c)}
        for m, c in pendientes_motivos_counter.most_common()
    ]

    n_closed = len(closed_with_time)
    tramo_1 = sum(1 for x in closed_with_time if x["dias"] <= 3.0)
    tramo_2 = sum(1 for x in closed_with_time if 3.0 < x["dias"] <= 7.0)
    tramo_3 = sum(1 for x in closed_with_time if 7.0 < x["dias"] <= 14.0)
    tramo_4 = sum(1 for x in closed_with_time if x["dias"] > 14.0)

    def _fmt_tramo(cnt: int) -> str:
        if n_closed == 0:
            return "0 (0,0 %)"
        pct = cnt / n_closed * 100.0
        return f"{cnt} ({format_decimal_es(pct, 1)} %)"

    tramos_rows = [
        {"tramo": "Hasta 3 días", "solicitudes": _fmt_tramo(tramo_1)},
        {"tramo": "Más de 3 y hasta 7 días", "solicitudes": _fmt_tramo(tramo_2)},
        {"tramo": "Más de 7 y hasta 14 días", "solicitudes": _fmt_tramo(tramo_3)},
        {"tramo": "Más de 14 días", "solicitudes": _fmt_tramo(tramo_4)},
    ]

    return {
        "periodo": periodo_clean,
        "periodo_titulo": periodo_titulo,
        "periodo_narrativa": periodo_narrativa,
        "mes_nombre": mes_nombre,
        "anio": anio,
        "rango_fechas_texto": rango_fechas_texto,
        "universo_periodo_texto": universo_periodo_texto,
        "fecha_corte": fecha_corte,
        "fecha_corte_larga": fecha_corte_larga,
        "ramo_codigo": ramo_codigo,
        "ramo_nombre": ramo_nombre,
        "cliente_nombre_corto": nombre_corto,
        "cliente_nombre_completo": nombre_completo,
        "cliente_nit": nit_formatted,
        "kpis": {
            "tramitadas": total_tramitadas,
            "emitidas": total_emitidas,
            "novedades": total_novedades,
            "declinadas": total_declinadas,
            "pendientes": total_pendientes,
            "atendidas": total_atendidas,
            "cerradas": total_cerradas,
            "efectividad_str": f"{format_decimal_es(efectividad_pct, 1)} %",
            "efectividad_detalle": f"{total_atendidas} de {total_tramitadas} solicitudes atendidas con resultado (emitidas o novedades atendidas)",
            "gestion_str": f"{format_decimal_es(gestion_pct, 1)} %",
            "gestion_detalle": f"{total_cerradas} de {total_tramitadas} solicitudes con respuesta final (atendidas o declinadas)",
            "tiempo_promedio_str": tiempo_promedio_kpi,
            "tiempo_promedio_detalle": tiempo_promedio_detalle,
        },
        "resumen_narrativa": resumen_narrativa,
        "trazabilidad_rows": trazabilidad_rows,
        "trazabilidad_total": trazabilidad_total,
        "declinadas_resumen_rows": declinadas_resumen_rows,
        "declinadas_detalle_rows": declinadas_list,
        "pendientes_resumen_rows": pendientes_resumen_rows,
        "pendientes_detalle_rows": pendientes_list,
        "pending_pagare_cartera_count": pending_pagare_cartera_count,
        "tiempos_medidas": {
            "cerradas": str(len(closed_with_time)),
            "promedio": tiempo_promedio_tabla5,
            "mediana": mediana_tabla5,
            "extremos": extremos_tabla5,
        },
        "tiempos_tramos_rows": tramos_rows,
        "tiempos_por_tipo_rows": tiempos_por_tipo_rows,
        "detalle_solicitudes_rows": matched_records,
    }
