# -*- coding: utf-8 -*-
"""DOCX Renderer service — generates official comparative Word documents.

Dynamically loads and populates the official Word layout templates located in
assets/layouts/ according to the selected insurance category:
- Autos (assets/layouts/Autos.docx)
- Copropiedades (assets/layouts/Copropiedades.docx)
- Hogar (assets/layouts/Hogar.docx)
- Todo Riesgo Construcción (assets/layouts/Todo_Riesgo_Construccion.docx)

Integrates ConceptMapper to associate insurer concepts and technical synonyms
with the layout's canonical SFC items.
"""
from __future__ import annotations

import logging
import os
import tempfile
from pathlib import Path
from typing import Any, Optional

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from app.core.config import settings
from app.services.concept_mapper import get_concept_mapper, normalize_str, ConceptMapper

logger = logging.getLogger(__name__)

# --------------------------------------------------------------------------- #
# Corporate Styling Constants                                                 #
# --------------------------------------------------------------------------- #
AZUL = RGBColor(0x1F, 0x38, 0x64)
AZUL_HEX = "1F3864"
GRIS_HEX = "F2F2F2"
AZUL_CLARO_HEX = "D9E2F3"
FUENTE = "Arial"
RELLENO = "NO ESPECIFICA"
ROMANOS = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"]


# --------------------------------------------------------------------------- #
# Formatting and Table Helpers                                                #
# --------------------------------------------------------------------------- #

def _fmt_cop(valor: Any, decimales: bool = True) -> str:
    """Format a value as Colombian pesos or return as-is if string."""
    if valor is None or str(valor).strip() == "":
        return RELLENO
    if isinstance(valor, (int, float)):
        entero = f"{valor:,.2f}" if decimales else f"{valor:,.0f}"
        entero = entero.replace(",", "@").replace(".", ",").replace("@", ".")
        return f"$ {entero}"
    return str(valor)


def _sombrear(celda: Any, hex_color: str) -> None:
    """Apply background shading to a cell."""
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    celda._tc.get_or_add_tcPr().append(shd)


def _bordes_tabla(tabla: Any, color: str = "7F7F7F", sz: int = 4) -> None:
    """Apply borders to a table."""
    tblPr = tabla._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def _repetir_encabezado(fila: Any) -> None:
    """Mark a table row to repeat as header on page breaks."""
    trPr = fila._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def _no_partir(fila: Any) -> None:
    """Prevent a table row from splitting across pages."""
    trPr = fila._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


def _texto(
    celda: Any,
    contenido: str,
    *,
    negrita: bool = False,
    tam: int = 9,
    color: RGBColor | None = None,
    centrado: bool = True,
    fuente: str = FUENTE,
) -> None:
    """Write formatted text to a table cell, handling multi-line content."""
    celda.text = ""
    p = celda.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if centrado else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    for i, linea in enumerate(str(contenido).split("\n")):
        if i:
            p = celda.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if centrado else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
        run = p.add_run(linea)
        run.font.name = fuente
        run.font.size = Pt(tam)
        run.bold = negrita
        if color is not None:
            run.font.color.rgb = color
    celda.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _nombre_corto(ase: dict) -> str:
    """Return a clean short display name for the insurer."""
    nombre = ase.get("nombre", "")
    id_ase = ase.get("id", "").lower()
    mapeo = {
        "sura": "SURA",
        "suramericana": "SURA",
        "zurich": "Zurich",
        "chubb": "Chubb",
        "axa": "AXA Colpatria",
        "axa_colpatria": "AXA Colpatria",
        "colpatria": "AXA Colpatria",
        "mundial": "Seguros Mundial",
        "hdi": "HDI",
        "bolivar": "Bolívar",
        "estado": "Seguros del Estado",
        "berkley": "Berkley",
        "davivienda": "Seguros Davivienda",
    }
    for k, v in mapeo.items():
        if k == id_ase or k in id_ase or k in nombre.lower():
            return v
    return nombre.split(" S.A")[0].strip() or id_ase.upper()


def _logo_en_celda(
    celda: Any,
    ruta: str,
    ancho_in: float = 1.25,
    pie: str | None = None,
    texto_alternativo: str | None = None,
) -> None:
    """Insert a logo image into a table cell, falling back to text if missing."""
    celda.text = ""
    p = celda.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if ruta and os.path.exists(ruta):
        p.add_run().add_picture(ruta, width=Inches(ancho_in))
    elif texto_alternativo:
        r = p.add_run(texto_alternativo)
        r.bold = True
        r.font.size = Pt(10)
        r.font.name = FUENTE
        r.font.color.rgb = AZUL
    if pie:
        p2 = celda.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p2.add_run(pie)
        r.bold = True
        r.font.size = Pt(8)
        r.font.name = FUENTE
    celda.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _ajustar_columnas(tabla: Any, n_esperadas: int, col_width: Any = Inches(1.5)) -> None:
    """Adjust number of columns in a table to match the expected count."""
    n_actuales = len(tabla.rows[0].cells)
    if n_actuales > n_esperadas:
        for _ in range(n_actuales - n_esperadas):
            col_a_eliminar = len(tabla.rows[0].cells) - 1
            for fila in tabla.rows:
                tc = fila.cells[col_a_eliminar]._tc
                fila._tr.remove(tc)
    elif n_actuales < n_esperadas:
        for _ in range(n_esperadas - n_actuales):
            tabla.add_column(col_width)


def _buscar_valor_concepto(
    conceptos_dict: dict[str, Any],
    etiqueta: str,
    categoria: str,
    subitem: str = "",
) -> Any:
    """Look up a concept value from an insurer's dictionary using ConceptMapper."""
    if not conceptos_dict:
        return RELLENO

    # Exact key match
    if etiqueta in conceptos_dict:
        val = conceptos_dict[etiqueta]
        return val if val not in (None, "") else RELLENO

    norm_etiqueta = normalize_str(etiqueta)
    for k, v in conceptos_dict.items():
        if normalize_str(k) == norm_etiqueta:
            return v if v not in (None, "") else RELLENO

    mapper = get_concept_mapper()
    sfc_target = mapper.map_concept(etiqueta, categoria)

    for k, v in conceptos_dict.items():
        if subitem and normalize_str(subitem) not in normalize_str(k):
            continue
        sfc_k = mapper.map_concept(k, categoria)
        if sfc_k and sfc_target and sfc_k == sfc_target:
            return v if v not in (None, "") else RELLENO
        norm_k = normalize_str(k)
        if norm_k in norm_etiqueta or norm_etiqueta in norm_k:
            return v if v not in (None, "") else RELLENO

    return RELLENO


def _asegura_membrete(doc: Document, assets: str) -> None:
    """Ensure the document header contains the corporate letterhead."""
    if not doc.sections:
        return
    seccion = doc.sections[0]
    encabezado = seccion.header
    tiene_imagen = bool(encabezado._element.findall(".//" + qn("a:blip")))
    if tiene_imagen:
        return
    logo = os.path.join(assets, "logo_mrc.png")
    p = encabezado.paragraphs[0] if encabezado.paragraphs else encabezado.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(logo):
        p.add_run().add_picture(logo, width=Inches(2.2))


# --------------------------------------------------------------------------- #
# Category-Specific Layout Renderers                                          #
# --------------------------------------------------------------------------- #

def _render_autos(doc: Document, data: dict, assets: str) -> None:
    """Populate the Autos.docx layout template."""
    meta = data.get("meta", {})
    aseguradoras = data.get("aseguradoras", [])
    rec = data.get("recomendacion") or {}
    n_ase = len(aseguradoras)

    # 1. Table 1: General Info (5 rows x 4 cols)
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        t1.rows[0].cells[1].text = str(meta.get("tomador", ""))
        t1.rows[0].cells[3].text = str(meta.get("identificacion", meta.get("nit", "")))
        t1.rows[1].cells[1].text = str(meta.get("marca", ""))
        t1.rows[1].cells[3].text = str(meta.get("placa", ""))
        t1.rows[2].cells[1].text = str(meta.get("linea", ""))
        t1.rows[2].cells[3].text = str(meta.get("modelo", ""))
        t1.rows[3].cells[1].text = str(meta.get("servicio", "Particular"))
        t1.rows[3].cells[3].text = str(meta.get("zona_circulacion", meta.get("ubicacion", "")))
        t1.rows[4].cells[1].text = _fmt_cop(meta.get("valor_asegurado"))
        t1.rows[4].cells[3].text = str(meta.get("accesorios", "NO ESPECIFICA"))

        for row in t1.rows:
            for cell in row.cells:
                cell.paragraphs[0].runs[0].font.size = Pt(9) if cell.paragraphs[0].runs else None

    # 2. Table 3: Quotes table (3 rows x (1 + N) cols)
    if len(doc.tables) > 3 and n_ase > 0:
        t3 = doc.tables[3]
        _ajustar_columnas(t3, 1 + n_ase)
        for i, ase in enumerate(aseguradoras):
            col = i + 1
            logo_path = os.path.join(assets, "logos", ase.get("logo", ""))
            _logo_en_celda(t3.rows[0].cells[col], logo_path, ancho_in=1.1, texto_alternativo=_nombre_corto(ase))
            op = ase.get("opciones", [{}])[0] if ase.get("opciones") else {}
            _texto(t3.rows[1].cells[col], op.get("etiqueta", op.get("modalidad", "")), tam=9)
            _texto(t3.rows[2].cells[col], _fmt_cop(op.get("prima")), tam=9, negrita=True)

    # 3. Table 4: Recommendation (2 rows x 2 cols)
    if len(doc.tables) > 4 and rec:
        t4 = doc.tables[4]
        elegida = next((a for a in aseguradoras if a.get("id") == rec.get("aseguradora_id")), None)
        logo_path = os.path.join(assets, "logos", elegida.get("logo", "") if elegida else "")
        _logo_en_celda(
            t4.rows[1].cells[0],
            logo_path,
            ancho_in=1.3,
            pie=rec.get("opcion", ""),
            texto_alternativo=_nombre_corto(elegida) if elegida else rec.get("opcion", ""),
        )
        der = t4.rows[1].cells[1]
        der.text = ""
        for i, vin in enumerate(rec.get("vinetas", [])):
            p = der.paragraphs[0] if i == 0 else der.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run("•  " + vin)
            r.font.size = Pt(9)
            r.font.name = FUENTE

    # 4. Table 6: Coverages matrix (12 rows x (1 + N) cols)
    if len(doc.tables) > 6 and n_ase > 0:
        t6 = doc.tables[6]
        _ajustar_columnas(t6, 1 + n_ase)
        for i, ase in enumerate(aseguradoras):
            col = i + 1
            logo_path = os.path.join(assets, "logos", ase.get("logo", ""))
            _logo_en_celda(t6.rows[0].cells[col], logo_path, ancho_in=1.0, texto_alternativo=_nombre_corto(ase))

        for row in t6.rows[1:]:
            concepto_label = row.cells[0].text.strip()
            for i, ase in enumerate(aseguradoras):
                val = _buscar_valor_concepto(ase.get("coberturas", {}), concepto_label, "Autos")
                _texto(row.cells[i + 1], _fmt_cop(val) if isinstance(val, (int, float)) else str(val), tam=8)

    # 5. Table 8: Deductibles matrix (7 rows x (1 + N) cols)
    if len(doc.tables) > 8 and n_ase > 0:
        t8 = doc.tables[8]
        _ajustar_columnas(t8, 1 + n_ase)
        for i, ase in enumerate(aseguradoras):
            col = i + 1
            logo_path = os.path.join(assets, "logos", ase.get("logo", ""))
            _logo_en_celda(t8.rows[0].cells[col], logo_path, ancho_in=1.0, texto_alternativo=_nombre_corto(ase))

        for row in t8.rows[1:]:
            concepto_label = row.cells[0].text.strip()
            for i, ase in enumerate(aseguradoras):
                val = _buscar_valor_concepto(ase.get("deducibles", {}), concepto_label, "Autos")
                _texto(row.cells[i + 1], str(val), tam=8)


def _render_copropiedades(doc: Document, data: dict, assets: str) -> None:
    """Populate the Copropiedades.docx layout template."""
    meta = data.get("meta", {})
    aseguradoras = data.get("aseguradoras", [])
    rec = data.get("recomendacion") or {}
    n_ase = len(aseguradoras)

    # 1. Table 0: General Info (6 rows x 2 cols)
    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        t0.rows[0].cells[1].text = str(meta.get("fecha", ""))
        t0.rows[1].cells[1].text = str(meta.get("tipo_cobertura", "MULTIRRIESGO COPROPIEDADES"))
        t0.rows[2].cells[1].text = str(meta.get("tomador", ""))
        t0.rows[3].cells[1].text = str(meta.get("asegurado", ""))
        t0.rows[4].cells[1].text = str(meta.get("beneficiario", "Terceros afectados / Copropietarios"))
        t0.rows[5].cells[1].text = str(meta.get("ubicacion", ""))

    # 2. Table 1: Valores Asegurados (10 rows x 2 cols)
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        for row in t1.rows[1:]:
            etq = row.cells[0].text.strip()
            val = _buscar_valor_concepto(
                aseguradoras[0].get("coberturas", {}) if aseguradoras else {},
                etq,
                "Copropiedades",
            )
            if val == RELLENO and meta.get("valor_asegurado") and "EDIFICIO" in etq:
                val = meta.get("valor_asegurado")
            _texto(row.cells[1], _fmt_cop(val) if isinstance(val, (int, float)) else str(val), tam=9, centrado=False)

    # 3. Table 2: Quotes Table (COMPAÑÍA, PRIMA, MODALIDAD)
    if len(doc.tables) > 2 and n_ase > 0:
        t2 = doc.tables[2]
        # Clean existing rows after header
        while len(t2.rows) > 1:
            tr = t2.rows[-1]._tr
            t2._tbl.remove(tr)

        for ase in aseguradoras:
            opciones = ase.get("opciones", [])
            if not opciones:
                opciones = [{"etiqueta": "", "prima": 0, "modalidad": RELLENO}]
            for op in opciones:
                fila = t2.add_row()
                logo_path = os.path.join(assets, "logos", ase.get("logo", ""))
                _logo_en_celda(fila.cells[0], logo_path, ancho_in=1.1, texto_alternativo=_nombre_corto(ase))
                _texto(fila.cells[1], _fmt_cop(op.get("prima")), tam=9, negrita=True)
                _texto(fila.cells[2], f"{op.get('modalidad', '')}\nTasa: {op.get('tasa', '')}", tam=9, centrado=False)
                _no_partir(fila)

    # 4. Table 3: Recommendation (1 row x 2 cols)
    if len(doc.tables) > 3 and rec:
        t3 = doc.tables[3]
        elegida = next((a for a in aseguradoras if a.get("id") == rec.get("aseguradora_id")), None)
        logo_path = os.path.join(assets, "logos", elegida.get("logo", "") if elegida else "")
        _logo_en_celda(
            t3.rows[0].cells[0],
            logo_path,
            ancho_in=1.3,
            pie=rec.get("opcion", ""),
            texto_alternativo=_nombre_corto(elegida) if elegida else rec.get("opcion", ""),
        )
        der = t3.rows[0].cells[1]
        der.text = ""
        for i, vin in enumerate(rec.get("vinetas", [])):
            p = der.paragraphs[0] if i == 0 else der.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run("•  " + vin)
            r.font.size = Pt(9)
            r.font.name = FUENTE

    # 5. Table 4: Coverages (12 rows x (1 + N) cols)
    if len(doc.tables) > 4 and n_ase > 0:
        t4 = doc.tables[4]
        _ajustar_columnas(t4, 1 + n_ase)
        for i, ase in enumerate(aseguradoras):
            col = i + 1
            logo_path = os.path.join(assets, "logos", ase.get("logo", ""))
            _logo_en_celda(t4.rows[0].cells[col], logo_path, ancho_in=1.0, texto_alternativo=_nombre_corto(ase))

        for row in t4.rows[1:]:
            concepto_label = row.cells[0].text.strip()
            # If section band header, leave as section header across row
            if "modulo" in concepto_label.lower():
                continue
            for i, ase in enumerate(aseguradoras):
                val = _buscar_valor_concepto(ase.get("coberturas", {}), concepto_label, "Copropiedades")
                _texto(row.cells[i + 1], _fmt_cop(val) if isinstance(val, (int, float)) else str(val), tam=8)

    # 6. Table 5: Deductibles (5 rows x (1 + N) cols)
    if len(doc.tables) > 5 and n_ase > 0:
        t5 = doc.tables[5]
        _ajustar_columnas(t5, 1 + n_ase)
        for i, ase in enumerate(aseguradoras):
            col = i + 1
            logo_path = os.path.join(assets, "logos", ase.get("logo", ""))
            _logo_en_celda(t5.rows[0].cells[col], logo_path, ancho_in=1.0, texto_alternativo=_nombre_corto(ase))

        for row in t5.rows[1:]:
            concepto_label = row.cells[0].text.strip()
            for i, ase in enumerate(aseguradoras):
                val = _buscar_valor_concepto(ase.get("deducibles", {}), concepto_label, "Copropiedades")
                _texto(row.cells[i + 1], str(val), tam=8)


def _render_hogar(doc: Document, data: dict, assets: str) -> None:
    """Populate the Hogar.docx layout template."""
    meta = data.get("meta", {})
    aseguradoras = data.get("aseguradoras", [])
    rec = data.get("recomendacion") or {}
    n_ase = len(aseguradoras)

    # 1. Table 0: General Info (13 rows x 2 cols)
    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        t0.rows[1].cells[1].text = str(meta.get("tomador", ""))
        t0.rows[2].cells[1].text = str(meta.get("cedula", meta.get("identificacion", "")))
        t0.rows[3].cells[1].text = str(meta.get("direccion", meta.get("ubicacion", "")))
        t0.rows[4].cells[1].text = str(meta.get("ciudad", "Bogotá D.C."))
        t0.rows[5].cells[1].text = str(meta.get("ano_construccion", "NO ESPECIFICA"))
        t0.rows[6].cells[1].text = _fmt_cop(meta.get("valor_edificio", meta.get("valor_asegurado")))
        t0.rows[7].cells[1].text = _fmt_cop(meta.get("valor_muebles"))
        t0.rows[8].cells[1].text = _fmt_cop(meta.get("valor_equipos"))
        t0.rows[9].cells[1].text = _fmt_cop(meta.get("valor_arte"))
        t0.rows[10].cells[1].text = _fmt_cop(meta.get("valor_dinero"))
        t0.rows[11].cells[1].text = str(meta.get("asegurado_actualmente", "NO ESPECIFICA"))
        t0.rows[12].cells[1].text = str(meta.get("siniestros_previos", "NO ESPECIFICA"))

    # 2. Table 1: Quotes Table (COMPAÑÍA, PRECIO INCLUIDO IVA)
    if len(doc.tables) > 1 and n_ase > 0:
        t1 = doc.tables[1]
        while len(t1.rows) > 1:
            tr = t1.rows[-1]._tr
            t1._tbl.remove(tr)

        for ase in aseguradoras:
            fila = t1.add_row()
            logo_path = os.path.join(assets, "logos", ase.get("logo", ""))
            _logo_en_celda(fila.cells[0], logo_path, ancho_in=1.1, texto_alternativo=_nombre_corto(ase))
            op = ase.get("opciones", [{}])[0] if ase.get("opciones") else {}
            _texto(fila.cells[1], _fmt_cop(op.get("prima")), tam=9, negrita=True)
            _no_partir(fila)

    # 3. Table 2: Recommendation (1 row x 2 cols)
    if len(doc.tables) > 2 and rec:
        t2 = doc.tables[2]
        elegida = next((a for a in aseguradoras if a.get("id") == rec.get("aseguradora_id")), None)
        logo_path = os.path.join(assets, "logos", elegida.get("logo", "") if elegida else "")
        _logo_en_celda(
            t2.rows[0].cells[0],
            logo_path,
            ancho_in=1.3,
            pie=rec.get("opcion", ""),
            texto_alternativo=_nombre_corto(elegida) if elegida else rec.get("opcion", ""),
        )
        der = t2.rows[0].cells[1]
        der.text = ""
        for i, vin in enumerate(rec.get("vinetas", [])):
            p = der.paragraphs[0] if i == 0 else der.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run("•  " + vin)
            r.font.size = Pt(9)
            r.font.name = FUENTE

    # 4. Table 3: Coverages (19 rows x (2 + N) cols)
    if len(doc.tables) > 3 and n_ase > 0:
        t3 = doc.tables[3]
        _ajustar_columnas(t3, 2 + n_ase)
        for i, ase in enumerate(aseguradoras):
            col = i + 2
            logo_path = os.path.join(assets, "logos", ase.get("logo", ""))
            _logo_en_celda(t3.rows[0].cells[col], logo_path, ancho_in=1.0, texto_alternativo=_nombre_corto(ase))

        for row in t3.rows[1:]:
            ramo_label = row.cells[0].text.strip()
            item_label = row.cells[1].text.strip()
            query = f"{ramo_label} {item_label}".strip()
            for i, ase in enumerate(aseguradoras):
                val = _buscar_valor_concepto(ase.get("coberturas", {}), query, "Hogar", subitem=item_label)
                if val == RELLENO:
                    val = _buscar_valor_concepto(ase.get("coberturas", {}), item_label, "Hogar")
                _texto(row.cells[i + 2], _fmt_cop(val) if isinstance(val, (int, float)) else str(val), tam=8)

    # 5. Table 4: Deductibles (5 rows x (1 + N) cols)
    if len(doc.tables) > 4 and n_ase > 0:
        t4 = doc.tables[4]
        _ajustar_columnas(t4, 1 + n_ase)
        for i, ase in enumerate(aseguradoras):
            col = i + 1
            logo_path = os.path.join(assets, "logos", ase.get("logo", ""))
            _logo_en_celda(t4.rows[0].cells[col], logo_path, ancho_in=1.0, texto_alternativo=_nombre_corto(ase))

        for row in t4.rows[1:]:
            concepto_label = row.cells[0].text.strip()
            for i, ase in enumerate(aseguradoras):
                val = _buscar_valor_concepto(ase.get("deducibles", {}), concepto_label, "Hogar")
                _texto(row.cells[i + 1], str(val), tam=8)


def _render_todo_riesgo_construccion(doc: Document, data: dict, assets: str) -> None:
    """Populate the Todo_Riesgo_Construccion.docx layout template."""
    meta = data.get("meta", {})
    aseguradoras = data.get("aseguradoras", [])
    rec = data.get("recomendacion") or {}
    n_ase = len(aseguradoras)

    # 1. Table 0: General Info (9 rows x 2 cols)
    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        vig_str = ""
        vc = meta.get("vigencia_construccion") or {}
        vm = meta.get("vigencia_mantenimiento") or {}
        if vc:
            vig_str += f"Construcción: {vc.get('desde', '')} - {vc.get('hasta', '')}\n"
        if vm:
            vig_str += f"Mantenimiento: {vm.get('desde', '')} - {vm.get('hasta', '')}"

        t0.rows[0].cells[1].text = str(meta.get("fecha", ""))
        t0.rows[1].cells[1].text = str(meta.get("tipo_cobertura", "TODO RIESGO CONSTRUCCIÓN Y MONTAJE"))
        t0.rows[2].cells[1].text = str(meta.get("tomador", ""))
        t0.rows[3].cells[1].text = str(meta.get("asegurado", ""))
        t0.rows[4].cells[1].text = str(meta.get("beneficiario", ""))
        t0.rows[5].cells[1].text = vig_str.strip() or RELLENO
        t0.rows[6].cells[1].text = str(meta.get("ubicacion", ""))
        t0.rows[7].cells[1].text = _fmt_cop(meta.get("valor_asegurado"))
        t0.rows[8].cells[1].text = str(meta.get("descripcion_proyecto", ""))

    # 2. Table 1: Quotes Table (COMPAÑÍA DE SEGUROS, TASA, PRIMA, MODALIDAD)
    if len(doc.tables) > 1 and n_ase > 0:
        t1 = doc.tables[1]
        while len(t1.rows) > 1:
            tr = t1.rows[-1]._tr
            t1._tbl.remove(tr)

        for ase in aseguradoras:
            opciones = ase.get("opciones", [])
            if not opciones:
                opciones = [{"etiqueta": "", "tasa": "", "prima": 0, "modalidad": RELLENO}]
            for op in opciones:
                fila = t1.add_row()
                logo_path = os.path.join(assets, "logos", ase.get("logo", ""))
                _logo_en_celda(fila.cells[0], logo_path, ancho_in=1.1, texto_alternativo=_nombre_corto(ase))
                _texto(fila.cells[1], op.get("tasa", RELLENO), tam=9)
                _texto(fila.cells[2], _fmt_cop(op.get("prima")), tam=9, negrita=True)
                _texto(fila.cells[3], op.get("modalidad", RELLENO), tam=9, centrado=False)
                _no_partir(fila)

    # 3. Table 2: Recommendation (1 row x 2 cols)
    if len(doc.tables) > 2 and rec:
        t2 = doc.tables[2]
        elegida = next((a for a in aseguradoras if a.get("id") == rec.get("aseguradora_id")), None)
        logo_path = os.path.join(assets, "logos", elegida.get("logo", "") if elegida else "")
        _logo_en_celda(
            t2.rows[0].cells[0],
            logo_path,
            ancho_in=1.3,
            pie=rec.get("opcion", ""),
            texto_alternativo=_nombre_corto(elegida) if elegida else rec.get("opcion", ""),
        )
        der = t2.rows[0].cells[1]
        der.text = ""
        for i, vin in enumerate(rec.get("vinetas", [])):
            p = der.paragraphs[0] if i == 0 else der.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run("•  " + vin)
            r.font.size = Pt(9)
            r.font.name = FUENTE

    # 4. Table 3: Coverages (42 rows x (1 + N) cols)
    if len(doc.tables) > 3 and n_ase > 0:
        t3 = doc.tables[3]
        _ajustar_columnas(t3, 1 + n_ase)
        for i, ase in enumerate(aseguradoras):
            col = i + 1
            logo_path = os.path.join(assets, "logos", ase.get("logo", ""))
            _logo_en_celda(t3.rows[0].cells[col], logo_path, ancho_in=1.0, texto_alternativo=_nombre_corto(ase))

        for row in t3.rows[1:]:
            concepto_label = row.cells[0].text.strip()
            if not concepto_label or "cobertura sección" in concepto_label.lower():
                continue
            for i, ase in enumerate(aseguradoras):
                val = _buscar_valor_concepto(ase.get("coberturas", {}), concepto_label, "Todo_Riesgo_Construccion")
                _texto(row.cells[i + 1], _fmt_cop(val) if isinstance(val, (int, float)) else str(val), tam=8)

    # 5. Table 4: Deductibles (21 rows x (1 + N) cols)
    if len(doc.tables) > 4 and n_ase > 0:
        t4 = doc.tables[4]
        _ajustar_columnas(t4, 1 + n_ase)
        for i, ase in enumerate(aseguradoras):
            col = i + 1
            logo_path = os.path.join(assets, "logos", ase.get("logo", ""))
            _logo_en_celda(t4.rows[0].cells[col], logo_path, ancho_in=1.0, texto_alternativo=_nombre_corto(ase))

        for row in t4.rows[1:]:
            concepto_label = row.cells[0].text.strip()
            if not concepto_label or "cobertura sección" in concepto_label.lower():
                continue
            for i, ase in enumerate(aseguradoras):
                val = _buscar_valor_concepto(ase.get("deducibles", {}), concepto_label, "Todo_Riesgo_Construccion")
                _texto(row.cells[i + 1], str(val), tam=8)


def _anexos_generales(doc: Document, data: dict) -> None:
    """Append general technical annexes if available in data."""
    aseguradoras = data.get("aseguradoras", [])
    subj = [a for a in aseguradoras if a.get("subjetividades")]
    if not subj:
        return

    doc.add_page_break()
    p = doc.add_paragraph()
    r = p.add_run("ANEXO: SUBJETIVIDADES Y CONDICIONES PARTICULARES")
    r.bold = True
    r.underline = True
    r.font.size = Pt(12)
    r.font.name = FUENTE
    r.font.color.rgb = AZUL

    for ase in subj:
        p_ase = doc.add_paragraph()
        r_ase = p_ase.add_run(ase.get("nombre", ase.get("id", "")))
        r_ase.bold = True
        r_ase.font.size = Pt(10)
        r_ase.font.name = FUENTE
        for s in ase.get("subjetividades", []):
            p_s = doc.add_paragraph()
            p_s.paragraph_format.left_indent = Inches(0.3)
            r_s = p_s.add_run("•  " + str(s))
            r_s.font.size = Pt(9)
            r_s.font.name = FUENTE


# --------------------------------------------------------------------------- #
# Public Rendering API                                                        #
# --------------------------------------------------------------------------- #

def render_comparative(
    data: dict,
    output_path: str | None = None,
    assets_dir: str | None = None,
    categoria: str | None = None,
) -> str:
    """Generate the comparative Word document according to category layout.

    Replaces previous hardcoded rendering with dynamic template instantiation
    from assets/layouts/{Category}.docx.

    Args:
        data: Consolidated comparative data dictionary.
        output_path: Destination path for the .docx file.
        assets_dir: Path to assets directory (logos and layouts).
        categoria: Optional explicit category name ('Autos', 'Copropiedades', 'Hogar', 'Todo_Riesgo_Construccion').

    Returns:
        Absolute path to the generated .docx file.
    """
    assets = assets_dir or str(settings.ASSETS_DIR)
    layouts_dir = Path(assets) / "layouts"

    # 1. Determine and normalize category
    target_cat = categoria or data.get("meta", {}).get("categoria") or data.get("meta", {}).get("tipo_cobertura", "")
    norm_cat = ConceptMapper.normalize_category(target_cat)

    # 2. Match layout file in assets/layouts/
    template_file = layouts_dir / f"{norm_cat}.docx"
    if not template_file.exists():
        # Case-insensitive / alias search
        for f in layouts_dir.glob("*.docx"):
            if f.name.startswith("~$"):
                continue
            if ConceptMapper.normalize_category(f.stem) == norm_cat:
                template_file = f
                break

    if not template_file.exists():
        template_file = layouts_dir / "Todo_Riesgo_Construccion.docx"
        if not template_file.exists():
            template_file = Path(assets) / "plantilla_membrete.docx"

    if output_path is None:
        fecha_slug = str(data.get("meta", {}).get("fecha", "output")).replace("/", "-")
        output_path = os.path.join(
            tempfile.gettempdir(),
            f"COMPARATIVO_{norm_cat.upper()}_{fecha_slug}.docx",
        )

    logger.info("Cargando plantilla layout para '%s': %s", norm_cat, template_file)
    doc = Document(str(template_file)) if template_file.exists() else Document()

    _asegura_membrete(doc, assets)

    # 3. Dispatch to specialized category renderer
    if norm_cat == "Autos":
        _render_autos(doc, data, assets)
    elif norm_cat == "Copropiedades":
        _render_copropiedades(doc, data, assets)
    elif norm_cat == "Hogar":
        _render_hogar(doc, data, assets)
    elif norm_cat == "Todo_Riesgo_Construccion":
        _render_todo_riesgo_construccion(doc, data, assets)
    else:
        # Fallback to TRC
        _render_todo_riesgo_construccion(doc, data, assets)

    _anexos_generales(doc, data)

    doc.save(output_path)
    logger.info("Documento comparativo renderizado exitosamente: %s", output_path)
    return output_path
