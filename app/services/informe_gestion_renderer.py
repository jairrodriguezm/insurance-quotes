# -*- coding: utf-8 -*-
"""PowerPoint (.pptx) and Word (.docx) Renderer for Management Reports (Informes de Gestión).

Primary output: Widescreen 16:9 PowerPoint presentation (.pptx) styled after
'Oferta servicios Flormorado.pdf' (Multiriesgos de Colombia - MRC), preserving
100% of the analytical content, KPIs, tables, narratives, and methodological notes.
"""
from __future__ import annotations

import logging
import math
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Inches, Pt

logger = logging.getLogger(__name__)

# --------------------------------------------------------------------------- #
# Brand Color Palette from "Oferta servicios Flormorado.pdf"                  #
# --------------------------------------------------------------------------- #
PPT_NAVY = RGBColor(0x05, 0x1D, 0x40)          # #051D40 Primary Dark Navy
PPT_NAVY_PANEL = RGBColor(0x0B, 0x2B, 0x5B)    # #0B2B5B Elevated Dark Panel
PPT_ROYAL = RGBColor(0x14, 0x5D, 0xA0)         # #145DA0 Secondary Royal Blue
PPT_SKY = RGBColor(0x56, 0xAE, 0xFF)           # #56AEFF Accent Sky Blue
PPT_CYAN = RGBColor(0x18, 0xAF, 0xD6)          # #18AFD6 Accent Cyan / Turquoise
PPT_ICE_CARD = RGBColor(0xD8, 0xF3, 0xFC)      # #D8F3FC Soft Ice Cyan Card (Page 9)
PPT_LIGHT_BG = RGBColor(0xF3, 0xF4, 0xFA)      # #F3F4FA Soft Corporate Light BG
PPT_ZEBRA = RGBColor(0xE8, 0xEE, 0xF7)         # #E8EEF7 Alternating Table Row
PPT_TOTAL_BG = RGBColor(0xD2, 0xE3, 0xFC)      # #D2E3FC Table Total Row Highlight
PPT_GREEN = RGBColor(0x53, 0x9B, 0x5C)         # #539B5C Corporate Green (Page 16)
PPT_EMERALD_BG = RGBColor(0xE6, 0xF4, 0xEA)    # #E6F4EA Soft Status Badge Green
PPT_AMBER_BG = RGBColor(0xFE, 0xF3, 0xC7)      # #FEF3C7 Soft Status Badge Pending
PPT_WHITE = RGBColor(0xFF, 0xFF, 0xFF)         # #FFFFFF Pure White
PPT_MUTED_LIGHT = RGBColor(0xB8, 0xD4, 0xF8)   # #B8D4F8 Soft Subtitle on Dark
PPT_MUTED_DARK = RGBColor(0x47, 0x55, 0x69)    # #475569 Slate Text on Light
PPT_DARK_TEXT = RGBColor(0x0F, 0x17, 0x2A)     # #0F172A Deep Ink Text

FONT_HEADING = "Century Gothic"
FONT_BODY = "Arial"

SLIDE_W_PT = 1440.0
SLIDE_H_PT = 810.0


# --------------------------------------------------------------------------- #
# Low-level XML & Shape Helpers for python-pptx                               #
# --------------------------------------------------------------------------- #
def _SubElement(parent: Any, tagname: str, **kwargs: str) -> Any:
    element = OxmlElement(tagname)
    element.attrib.update(kwargs)
    parent.append(element)
    return element


def _set_cell_border(cell: Any, border_color_hex: str = "CBD5E1", width_pt: float = 0.75) -> None:
    """Apply crisp borders to all 4 sides of a PowerPoint table cell via XML."""
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        existing = tcPr.find(qn(tag))
        if existing is not None:
            tcPr.remove(existing)
    w_emu = str(int(round(width_pt * 12700)))
    for side in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = _SubElement(tcPr, side, w=w_emu, cap="flat", cmpd="sng", algn="ctr")
        solidFill = _SubElement(ln, "a:solidFill")
        _SubElement(solidFill, "a:srgbClr", val=border_color_hex)
        _SubElement(ln, "a:prstDash", val="solid")


def _set_cell_margins(
    cell: Any,
    top_pt: float = 5.0,
    bottom_pt: float = 5.0,
    left_pt: float = 8.0,
    right_pt: float = 8.0,
) -> None:
    """Set internal padding of a table cell in points."""
    cell.margin_top = Pt(top_pt)
    cell.margin_bottom = Pt(bottom_pt)
    cell.margin_left = Pt(left_pt)
    cell.margin_right = Pt(right_pt)


def _format_run(
    run: Any,
    size_pt: float,
    bold: bool = False,
    italic: bool = False,
    color: RGBColor = PPT_WHITE,
    font_name: str = FONT_BODY,
) -> None:
    """Apply font family, size, weight, style, and RGB color to a text run."""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def _add_textbox(
    slide: Any,
    x_pt: float,
    y_pt: float,
    w_pt: float,
    h_pt: float,
    text: str,
    size_pt: float = 14.0,
    bold: bool = False,
    italic: bool = False,
    color: RGBColor = PPT_WHITE,
    font_name: str = FONT_BODY,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    v_anchor: MSO_ANCHOR = MSO_ANCHOR.TOP,
    word_wrap: bool = True,
) -> Any:
    """Add a positioned text box with single paragraph styling."""
    box = slide.shapes.add_textbox(Pt(x_pt), Pt(y_pt), Pt(w_pt), Pt(h_pt))
    tf = box.text_frame
    tf.word_wrap = word_wrap
    tf.vertical_anchor = v_anchor
    tf.margin_left = Pt(0)
    tf.margin_right = Pt(0)
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    _format_run(run, size_pt=size_pt, bold=bold, italic=italic, color=color, font_name=font_name)
    return box


def _add_rect(
    slide: Any,
    x_pt: float,
    y_pt: float,
    w_pt: float,
    h_pt: float,
    fill_color: RGBColor | None = None,
    border_color: RGBColor | None = None,
    border_pt: float = 1.0,
    rounded: bool = False,
) -> Any:
    """Add a rectangle or rounded rectangle shape."""
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, Pt(x_pt), Pt(y_pt), Pt(w_pt), Pt(h_pt))
    if rounded and hasattr(shp, "adjustments") and len(shp.adjustments) > 0:
        shp.adjustments[0] = 0.04
    if fill_color is not None:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill_color
    else:
        shp.fill.background()
    if border_color is not None:
        shp.line.color.rgb = border_color
        shp.line.width = Pt(border_pt)
    else:
        shp.line.fill.background()
    return shp


def _set_solid_slide_bg(slide: Any, color: RGBColor) -> None:
    """Fill the entire slide background with a solid RGB color."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_diamond_accents(slide: Any, origin_x: float = 1280.0, origin_y: float = 20.0, color: RGBColor = PPT_ROYAL) -> None:
    """Draw the signature Flormorado diamond grid motif using vector shapes."""
    coords = [
        (origin_x, origin_y),
        (origin_x + 70, origin_y),
        (origin_x + 35, origin_y + 38),
        (origin_x + 105, origin_y + 38),
        (origin_x, origin_y + 76),
        (origin_x + 70, origin_y + 76),
    ]
    for dx, dy in coords:
        dm = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, Pt(dx), Pt(dy), Pt(18), Pt(18))
        dm.fill.solid()
        dm.fill.fore_color.rgb = color
        dm.line.fill.background()


def _add_white_slide_chrome(
    slide: Any,
    title: str,
    subtitle: str,
    page_label: str = "",
) -> None:
    """Apply clean pure-white corporate chrome (#FFFFFF bg, navy/royal left stripes, divider, footer) with zero images."""
    _set_solid_slide_bg(slide, PPT_WHITE)

    # Left geometric accent stripes
    _add_rect(slide, 0, 0, 14, SLIDE_H_PT, fill_color=PPT_NAVY)
    _add_rect(slide, 14, 0, 5, SLIDE_H_PT, fill_color=PPT_ROYAL)
    _add_rect(slide, 19, 0, 3, SLIDE_H_PT, fill_color=PPT_CYAN)

    # Top-right corporate text label (image-free)
    _add_textbox(
        slide,
        x_pt=1110,
        y_pt=32,
        w_pt=282,
        h_pt=24,
        text="MULTIRIESGOS DE COLOMBIA",
        size_pt=10.5,
        bold=True,
        color=PPT_ROYAL,
        font_name=FONT_HEADING,
        align=PP_ALIGN.RIGHT,
    )

    # Title & subtitle
    _add_textbox(
        slide,
        x_pt=54,
        y_pt=26,
        w_pt=1060,
        h_pt=42,
        text=title.upper(),
        size_pt=28.0,
        bold=True,
        color=PPT_NAVY,
        font_name=FONT_HEADING,
    )
    _add_textbox(
        slide,
        x_pt=54,
        y_pt=68,
        w_pt=1280,
        h_pt=26,
        text=subtitle,
        size_pt=13.5,
        bold=True,
        color=PPT_ROYAL,
        font_name=FONT_BODY,
    )

    # Bottom border under header
    _add_rect(slide, 54, 100, 1338, 2.0, fill_color=PPT_NAVY)

    # Footer
    footer_txt = "Multiriesgos de Colombia Ltda. · Informe de gestión de solicitudes"
    if page_label:
        footer_txt += f" · {page_label}"
    _add_textbox(
        slide,
        x_pt=54,
        y_pt=776,
        w_pt=1338,
        h_pt=22,
        text=footer_txt,
        size_pt=10.5,
        color=PPT_MUTED_DARK,
        align=PP_ALIGN.RIGHT,
    )


def _create_styled_table(
    slide: Any,
    x_pt: float,
    y_pt: float,
    col_widths_pt: list[float],
    headers: list[str],
    col_alignments: list[PP_ALIGN],
    rows_data: list[list[str]],
    header_fill: RGBColor = PPT_NAVY,
    font_size_pt: float = 12.0,
    row_height_pt: float = 34.0,
    header_height_pt: float = 38.0,
    bold_cols: set[int] | None = None,
    has_total_row: bool = False,
    border_hex: str = "CBD5E1",
) -> tuple[Any, float]:
    """Create a richly styled PowerPoint table and return (table, total_height_pt)."""
    bold_cols = bold_cols or set()
    total_w_pt = sum(col_widths_pt)
    n_rows = 1 + len(rows_data)
    n_cols = len(col_widths_pt)
    est_h_pt = header_height_pt + len(rows_data) * row_height_pt

    graphic_frame = slide.shapes.add_table(
        n_rows,
        n_cols,
        Pt(x_pt),
        Pt(y_pt),
        Pt(total_w_pt),
        Pt(est_h_pt),
    )
    table = graphic_frame.table

    for c_idx, w_pt in enumerate(col_widths_pt):
        table.columns[c_idx].width = Pt(w_pt)

    # Style Header Row
    table.rows[0].height = Pt(header_height_pt)
    for c_idx, h_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_fill
        _set_cell_border(cell, border_color_hex=border_hex, width_pt=0.75)
        _set_cell_margins(cell, top_pt=6, bottom_pt=6, left_pt=10, right_pt=10)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = ""
        p.alignment = col_alignments[c_idx] if c_idx == 0 else PP_ALIGN.CENTER
        if len(headers) <= 3 and c_idx == len(headers) - 1 and col_alignments[c_idx] == PP_ALIGN.LEFT:
            p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = h_text
        _format_run(run, size_pt=font_size_pt, bold=True, color=PPT_WHITE, font_name=FONT_BODY)

    # Style Data Rows
    actual_total_h = header_height_pt
    for r_offset, row_vals in enumerate(rows_data):
        r_idx = r_offset + 1
        is_last = r_offset == len(rows_data) - 1
        is_total = has_total_row and is_last

        # Estimate row height if multiline content exists
        max_lines = 1
        for c_idx, val_text in enumerate(row_vals):
            col_w = max(80.0, col_widths_pt[c_idx] - 18.0)
            chars_per_line = max(12, int(col_w / (font_size_pt * 0.55)))
            lines = str(val_text).split("\n")
            line_cnt = sum(max(1, math.ceil(len(line) / chars_per_line)) for line in lines)
            if line_cnt > max_lines:
                max_lines = line_cnt

        this_row_h = max(row_height_pt, max_lines * (font_size_pt + 5.5) + 10.0)
        table.rows[r_idx].height = Pt(this_row_h)
        actual_total_h += this_row_h

        if is_total:
            row_fill = PPT_TOTAL_BG
        elif r_idx % 2 == 0:
            row_fill = PPT_ZEBRA
        else:
            row_fill = PPT_WHITE

        for c_idx, val_text in enumerate(row_vals):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_fill
            _set_cell_border(cell, border_color_hex=border_hex, width_pt=0.75)
            _set_cell_margins(cell, top_pt=5, bottom_pt=5, left_pt=10, right_pt=10)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

            tf = cell.text_frame
            tf.word_wrap = True
            lines = str(val_text).split("\n")
            for l_idx, line_str in enumerate(lines):
                p = tf.paragraphs[0] if l_idx == 0 else tf.add_paragraph()
                if l_idx == 0:
                    p.text = ""
                p.alignment = col_alignments[c_idx]

                is_bold = is_total or (c_idx in bold_cols and l_idx == 0)
                if l_idx > 0 and line_str.startswith("Integrantes:"):
                    run = p.add_run()
                    run.text = line_str
                    _format_run(
                        run,
                        size_pt=max(9.0, font_size_pt - 1.5),
                        bold=False,
                        color=PPT_MUTED_DARK,
                        font_name=FONT_BODY,
                    )
                else:
                    txt_color = PPT_NAVY if (is_total or is_bold) else PPT_DARK_TEXT
                    run = p.add_run()
                    run.text = line_str
                    _format_run(
                        run,
                        size_pt=font_size_pt,
                        bold=is_bold,
                        color=txt_color,
                        font_name=FONT_BODY,
                    )

    return table, actual_total_h


# --------------------------------------------------------------------------- #
# Main PowerPoint Renderer                                                    #
# --------------------------------------------------------------------------- #
def render_informe_gestion_pptx(
    data: dict[str, Any],
    output_path: str | Path,
    assets_dir: str | Path | None = None,
) -> str:
    """Render the complete Management Report (.pptx): blue cover & closing slides, pure white content slides, zero images."""
    prs = Presentation()
    prs.slide_width = Pt(SLIDE_W_PT)
    prs.slide_height = Pt(SLIDE_H_PT)
    blank_layout = prs.slide_layouts[6]

    cliente_corto = data["cliente_nombre_corto"]
    cliente_completo = data["cliente_nombre_completo"]
    periodo_titulo = data["periodo_titulo"]
    nit_prefix = f"NIT {data['cliente_nit']} · " if data.get("cliente_nit") else ""
    subtitle_ctx = f"{cliente_completo} · Informe de gestión · {periodo_titulo}"

    # ================================================================== #
    # SLIDE 1: PORTADA / INICIO (Fondo azul corporativo, sin imágenes)   #
    # ================================================================== #
    s1 = prs.slides.add_slide(blank_layout)
    _set_solid_slide_bg(s1, PPT_NAVY)

    # Vector geometric accents (zero images)
    _add_rect(s1, 0, 0, 18, SLIDE_H_PT, fill_color=PPT_ROYAL)
    _add_rect(s1, 18, 0, 5, SLIDE_H_PT, fill_color=PPT_CYAN)
    _add_diamond_accents(s1, origin_x=1220, origin_y=52, color=RGBColor(0x0D, 0x35, 0x68))
    _add_diamond_accents(s1, origin_x=1220, origin_y=640, color=PPT_ROYAL)

    _add_textbox(
        s1,
        x_pt=96,
        y_pt=64,
        w_pt=720,
        h_pt=28,
        text="MULTIRIESGOS DE COLOMBIA LTDA.",
        size_pt=13.5,
        bold=True,
        color=PPT_CYAN,
        font_name=FONT_HEADING,
    )

    # Slogan: "Construimos el futuro / Asegurando el presente"
    _add_textbox(
        s1,
        x_pt=96,
        y_pt=144,
        w_pt=900,
        h_pt=40,
        text="Construimos el futuro",
        size_pt=27.0,
        italic=True,
        color=PPT_SKY,
        font_name=FONT_BODY,
    )
    _add_textbox(
        s1,
        x_pt=96,
        y_pt=184,
        w_pt=900,
        h_pt=40,
        text="Asegurando el presente",
        size_pt=27.0,
        italic=True,
        color=PPT_WHITE,
        font_name=FONT_BODY,
    )

    # Period badge pill
    _add_rect(s1, 96, 256, 440, 38, fill_color=PPT_ROYAL, border_color=PPT_CYAN, border_pt=1.2, rounded=True)
    _add_textbox(
        s1,
        x_pt=112,
        y_pt=259,
        w_pt=408,
        h_pt=32,
        text=f"Informe de gestión · {periodo_titulo}",
        size_pt=15.0,
        bold=True,
        color=PPT_WHITE,
        font_name=FONT_HEADING,
        v_anchor=MSO_ANCHOR.MIDDLE,
    )

    # Main Title: INFORME DE GESTIÓN MRC
    _add_textbox(
        s1,
        x_pt=96,
        y_pt=314,
        w_pt=1100,
        h_pt=134,
        text="INFORME DE\nGESTIÓN MRC",
        size_pt=54.0,
        bold=True,
        color=PPT_WHITE,
        font_name=FONT_HEADING,
    )

    # Client Name in Sky Blue (#56AEFF)
    client_font_sz = 28.0 if len(cliente_completo) > 35 else 34.0
    _add_textbox(
        s1,
        x_pt=96,
        y_pt=462,
        w_pt=1160,
        h_pt=82,
        text=cliente_completo,
        size_pt=client_font_sz,
        bold=True,
        color=PPT_SKY,
        font_name=FONT_HEADING,
    )

    # Scope & Cutoff metadata card at bottom
    _add_rect(s1, 96, 566, 1140, 96, fill_color=PPT_NAVY_PANEL, border_color=PPT_ROYAL, border_pt=1.2, rounded=True)
    _add_rect(s1, 96, 566, 8, 96, fill_color=PPT_CYAN)
    _add_textbox(
        s1,
        x_pt=120,
        y_pt=576,
        w_pt=1090,
        h_pt=36,
        text=f"{nit_prefix}Incluye solicitudes como tomador, asegurado e integrante de consorcios o uniones temporales",
        size_pt=13.0,
        bold=False,
        color=PPT_WHITE,
        font_name=FONT_BODY,
    )
    _add_textbox(
        s1,
        x_pt=120,
        y_pt=616,
        w_pt=1090,
        h_pt=36,
        text=(
            f"Ramo: {data['ramo_nombre']} · Solicitudes recibidas {data['rango_fechas_texto']} · "
            f"Estado de las solicitudes al {data['fecha_corte_larga']}"
        ),
        size_pt=12.5,
        bold=False,
        color=PPT_MUTED_LIGHT,
        font_name=FONT_BODY,
    )

    # ================================================================== #
    # SLIDE 2: CONTENIDO (Fondo blanco)                                  #
    # ================================================================== #
    s2 = prs.slides.add_slide(blank_layout)
    _add_white_slide_chrome(
        s2,
        title="Contenido",
        subtitle=subtitle_ctx,
        page_label="Índice",
    )

    toc_items = [
        ("1", "Resumen del mes\ne indicadores"),
        ("2", "Trazabilidad de\nlas solicitudes"),
        ("3", "Solicitudes declinadas\ny motivo"),
        ("4", "Solicitudes pendientes\ny motivo"),
        ("5", "Tiempos de\natención"),
        ("6", "Detalle y notas\nmetodológicas"),
    ]
    card_w = 380.0
    card_h = 240.0
    gap_x = 30.0
    gap_y = 28.0
    start_x = 123.0
    start_y = 160.0

    for idx, (num_str, label_str) in enumerate(toc_items):
        col = idx % 3
        row = idx // 3
        cx = start_x + col * (card_w + gap_x)
        cy = start_y + row * (card_h + gap_y)

        _add_rect(s2, cx, cy, card_w, card_h, fill_color=PPT_WHITE, border_color=RGBColor(0xCB, 0xD5, 0xE1), border_pt=1.5, rounded=True)
        _add_rect(s2, cx, cy, card_w, 9, fill_color=PPT_ROYAL if idx % 2 == 0 else PPT_NAVY)
        _add_textbox(
            s2,
            x_pt=cx,
            y_pt=cy + 26,
            w_pt=card_w,
            h_pt=66,
            text=num_str,
            size_pt=46.0,
            bold=True,
            color=PPT_ROYAL,
            font_name=FONT_HEADING,
            align=PP_ALIGN.CENTER,
        )
        _add_rect(s2, cx + 80, cy + 108, card_w - 160, 2.5, fill_color=PPT_CYAN)
        _add_textbox(
            s2,
            x_pt=cx + 20,
            y_pt=cy + 126,
            w_pt=card_w - 40,
            h_pt=82,
            text=label_str,
            size_pt=18.0,
            bold=True,
            color=PPT_NAVY,
            font_name=FONT_BODY,
            align=PP_ALIGN.CENTER,
        )

    # ================================================================== #
    # SLIDE 3: 1. RESUMEN DEL MES (Fondo blanco)                         #
    # ================================================================== #
    s3 = prs.slides.add_slide(blank_layout)
    _add_white_slide_chrome(
        s3,
        title="1. Resumen del mes",
        subtitle=subtitle_ctx,
        page_label="Sección 1",
    )

    kpis = data["kpis"]
    kpi_items = [
        (str(kpis["tramitadas"]), "Tramitadas", PPT_NAVY),
        (str(kpis["emitidas"]), "Emitidas", PPT_ROYAL),
        (str(kpis["novedades"]), "Novedades atendidas", PPT_CYAN),
        (str(kpis["declinadas"]), "Declinadas", PPT_ROYAL),
        (str(kpis["pendientes"]), "Pendientes", PPT_NAVY),
    ]

    kpi_w = 250.0
    kpi_h = 148.0
    kpi_gap = 22.0
    kpi_start_x = 54.0
    kpi_y = 124.0

    for c_idx, (num_str, label_str, accent_clr) in enumerate(kpi_items):
        kx = kpi_start_x + c_idx * (kpi_w + kpi_gap)
        _add_rect(s3, kx, kpi_y, kpi_w, kpi_h, fill_color=PPT_LIGHT_BG, border_color=RGBColor(0xCB, 0xD5, 0xE1), border_pt=1.2, rounded=True)
        _add_rect(s3, kx, kpi_y, kpi_w, 8, fill_color=accent_clr)

        _add_textbox(
            s3,
            x_pt=kx,
            y_pt=kpi_y + 18,
            w_pt=kpi_w,
            h_pt=72,
            text=num_str,
            size_pt=48.0,
            bold=True,
            color=PPT_NAVY,
            font_name=FONT_HEADING,
            align=PP_ALIGN.CENTER,
        )
        _add_textbox(
            s3,
            x_pt=kx + 8,
            y_pt=kpi_y + 98,
            w_pt=kpi_w - 16,
            h_pt=38,
            text=label_str,
            size_pt=14.5,
            bold=True,
            color=PPT_ROYAL,
            font_name=FONT_BODY,
            align=PP_ALIGN.CENTER,
        )

    # Narrative Callout Banner
    _add_rect(s3, 54, 294, 1338, 86, fill_color=PPT_ICE_CARD, border_color=PPT_CYAN, border_pt=1.2, rounded=True)
    _add_rect(s3, 54, 294, 8, 86, fill_color=PPT_ROYAL)
    _add_textbox(
        s3,
        x_pt=78,
        y_pt=304,
        w_pt=1296,
        h_pt=66,
        text=data["resumen_narrativa"],
        size_pt=15.5,
        bold=False,
        color=PPT_NAVY,
        font_name=FONT_BODY,
        v_anchor=MSO_ANCHOR.MIDDLE,
    )

    # Table 1: Key Indicators (Indicador | Resultado | Cómo se calcula / alcance)
    _create_styled_table(
        s3,
        x_pt=54,
        y_pt=404,
        col_widths_pt=[380.0, 210.0, 748.0],
        headers=["Indicador", "Resultado", "Cómo se calcula / alcance"],
        col_alignments=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT],
        rows_data=[
            [
                "Índice de efectividad de atención",
                kpis["efectividad_str"],
                kpis["efectividad_detalle"],
            ],
            [
                "Índice de gestión",
                kpis["gestion_str"],
                kpis["gestion_detalle"],
            ],
            [
                "Tiempo promedio de atención",
                kpis["tiempo_promedio_str"],
                kpis["tiempo_promedio_detalle"],
            ],
        ],
        header_fill=PPT_NAVY,
        font_size_pt=14.0,
        row_height_pt=68.0,
        header_height_pt=46.0,
        bold_cols={0, 1},
    )

    # ================================================================== #
    # SLIDE 4: 2. TRAZABILIDAD DE LAS SOLICITUDES (Fondo blanco)         #
    # ================================================================== #
    s4 = prs.slides.add_slide(blank_layout)
    _add_white_slide_chrome(
        s4,
        title="2. Trazabilidad de las solicitudes",
        subtitle=subtitle_ctx,
        page_label="Sección 2",
    )

    traz_rows_matrix: list[list[str]] = []
    for r in data["trazabilidad_rows"]:
        traz_rows_matrix.append(
            [
                r["tipo"],
                r["tramitadas"],
                r["emitidas"],
                r["novedades"],
                r["declinadas"],
                r["pendientes"],
                r["efectividad"],
                r["dias_prom"],
            ]
        )
    tot = data["trazabilidad_total"]
    traz_rows_matrix.append(
        [
            tot["tipo"],
            tot["tramitadas"],
            tot["emitidas"],
            tot["novedades"],
            tot["declinadas"],
            tot["pendientes"],
            tot["efectividad"],
            tot["dias_prom"],
        ]
    )

    _, traz_h = _create_styled_table(
        s4,
        x_pt=54,
        y_pt=128,
        col_widths_pt=[288.0, 145.0, 140.0, 185.0, 145.0, 145.0, 155.0, 135.0],
        headers=[
            "Tipo de solicitud",
            "Tramitadas",
            "Emitidas",
            "Novedades atendidas",
            "Declinadas",
            "Pendientes",
            "Efectividad",
            "Días prom.",
        ],
        col_alignments=[
            PP_ALIGN.LEFT,
            PP_ALIGN.CENTER,
            PP_ALIGN.CENTER,
            PP_ALIGN.CENTER,
            PP_ALIGN.CENTER,
            PP_ALIGN.CENTER,
            PP_ALIGN.CENTER,
            PP_ALIGN.CENTER,
        ],
        rows_data=traz_rows_matrix,
        header_fill=PPT_NAVY,
        font_size_pt=13.5,
        row_height_pt=46.0,
        header_height_pt=46.0,
        bold_cols={0},
        has_total_row=True,
    )

    note_y = min(685.0, 128.0 + traz_h + 24.0)
    _add_rect(s4, 54, note_y, 1338, 62, fill_color=PPT_ICE_CARD, border_color=PPT_CYAN, border_pt=1.2, rounded=True)
    _add_rect(s4, 54, note_y, 8, 62, fill_color=PPT_ROYAL)
    _add_textbox(
        s4,
        x_pt=76,
        y_pt=note_y + 8,
        w_pt=1300,
        h_pt=46,
        text=(
            "Emitidas: pólizas nuevas expedidas. Novedades atendidas: modificaciones "
            "(prórrogas, adiciones, aclaraciones) ya resueltas. Días promedio: tiempo "
            "transcurrido entre la recepción de la solicitud y su cierre."
        ),
        size_pt=13.0,
        color=PPT_NAVY,
        v_anchor=MSO_ANCHOR.MIDDLE,
    )

    # ================================================================== #
    # SLIDE 5: 3. SOLICITUDES DECLINADAS Y MOTIVO (Fondo blanco)         #
    # ================================================================== #
    declinadas_detalle = data["declinadas_detalle_rows"]
    if not declinadas_detalle:
        s5 = prs.slides.add_slide(blank_layout)
        _add_white_slide_chrome(
            s5,
            title="3. Solicitudes declinadas y motivo",
            subtitle=subtitle_ctx,
            page_label="Sección 3",
        )
        # Visual highlight card for 0 declined requests on white slide
        _add_rect(s5, 54, 160, 1338, 240, fill_color=PPT_EMERALD_BG, border_color=PPT_GREEN, border_pt=1.5, rounded=True)
        badge_circle = s5.shapes.add_shape(MSO_SHAPE.OVAL, Pt(104), Pt(218), Pt(120), Pt(120))
        badge_circle.fill.solid()
        badge_circle.fill.fore_color.rgb = PPT_GREEN
        badge_circle.line.color.rgb = PPT_WHITE
        badge_circle.line.width = Pt(2.0)
        _add_textbox(
            s5,
            x_pt=104,
            y_pt=238,
            w_pt=120,
            h_pt=80,
            text="0",
            size_pt=52.0,
            bold=True,
            color=PPT_WHITE,
            font_name=FONT_HEADING,
            align=PP_ALIGN.CENTER,
        )
        _add_textbox(
            s5,
            x_pt=258,
            y_pt=216,
            w_pt=1080,
            h_pt=50,
            text=f"{data['periodo_narrativa']} no se declinó ninguna solicitud.",
            size_pt=26.0,
            bold=True,
            color=PPT_NAVY,
            font_name=FONT_HEADING,
        )
        _add_textbox(
            s5,
            x_pt=258,
            y_pt=278,
            w_pt=1080,
            h_pt=60,
            text=(
                f"El 100 % de las solicitudes cerradas de {cliente_corto} en el período concluyeron "
                f"con póliza emitida o novedad atendida satisfactoriamente."
            ),
            size_pt=16.0,
            color=PPT_DARK_TEXT,
        )
    else:
        cnt_dec = len(declinadas_detalle)
        dec_txt = (
            f"{data['periodo_narrativa']} se declinó 1 solicitud, por los siguientes motivos:"
            if cnt_dec == 1
            else f"{data['periodo_narrativa']} se declinaron {cnt_dec} solicitudes, por los siguientes motivos:"
        )
        # Paginate declined detail rows (max 6 per slide)
        dec_chunks = [declinadas_detalle[i : i + 6] for i in range(0, cnt_dec, 6)]
        for c_idx, chunk in enumerate(dec_chunks):
            s_dec = prs.slides.add_slide(blank_layout)
            sfx = f" ({c_idx + 1}/{len(dec_chunks)})" if len(dec_chunks) > 1 else ""
            _add_white_slide_chrome(
                s_dec,
                title=f"3. Solicitudes declinadas y motivo{sfx}",
                subtitle=subtitle_ctx,
                page_label="Sección 3",
            )
            cur_y = 118.0
            if c_idx == 0:
                _add_textbox(
                    s_dec,
                    x_pt=54,
                    y_pt=cur_y,
                    w_pt=1338,
                    h_pt=32,
                    text=dec_txt,
                    size_pt=15.0,
                    bold=True,
                    color=PPT_NAVY,
                )
                cur_y += 38.0
                _, res_h = _create_styled_table(
                    s_dec,
                    x_pt=54,
                    y_pt=cur_y,
                    col_widths_pt=[980.0, 358.0],
                    headers=["Motivo", "Solicitudes"],
                    col_alignments=[PP_ALIGN.LEFT, PP_ALIGN.CENTER],
                    rows_data=[[r["motivo"], r["cantidad"]] for r in data["declinadas_resumen_rows"][:5]],
                    header_fill=PPT_ROYAL,
                    font_size_pt=12.5,
                    row_height_pt=32.0,
                    header_height_pt=36.0,
                )
                cur_y += res_h + 20.0

            _create_styled_table(
                s_dec,
                x_pt=54,
                y_pt=cur_y,
                col_widths_pt=[160.0, 150.0, 240.0, 380.0, 408.0],
                headers=["Solicitud", "Recibida", "Tipo", "Tomador / Asegurado", "Motivo"],
                col_alignments=[
                    PP_ALIGN.CENTER,
                    PP_ALIGN.CENTER,
                    PP_ALIGN.LEFT,
                    PP_ALIGN.LEFT,
                    PP_ALIGN.LEFT,
                ],
                rows_data=[
                    [
                        r["id"],
                        r["recibida"],
                        r["tipo"],
                        r["tomador_asegurado"],
                        r["motivo_declinada"],
                    ]
                    for r in chunk
                ],
                header_fill=PPT_NAVY,
                font_size_pt=11.5,
                row_height_pt=38.0,
                header_height_pt=38.0,
                bold_cols={0},
            )

    # ================================================================== #
    # SLIDE 6: 4. SOLICITUDES PENDIENTES Y MOTIVO (Fondo blanco)         #
    # ================================================================== #
    pendientes_detalle = data["pendientes_detalle_rows"]
    if not pendientes_detalle:
        s6 = prs.slides.add_slide(blank_layout)
        _add_white_slide_chrome(
            s6,
            title="4. Solicitudes pendientes y motivo",
            subtitle=subtitle_ctx,
            page_label="Sección 4",
        )
        _add_rect(s6, 54, 160, 1338, 220, fill_color=PPT_LIGHT_BG, border_color=PPT_ROYAL, border_pt=1.5, rounded=True)
        _add_textbox(
            s6,
            x_pt=90,
            y_pt=230,
            w_pt=1260,
            h_pt=60,
            text=f"Al {data['fecha_corte_larga']} no queda ninguna solicitud pendiente.",
            size_pt=24.0,
            bold=True,
            color=PPT_NAVY,
            font_name=FONT_HEADING,
            align=PP_ALIGN.CENTER,
        )
    else:
        cnt_pen = len(pendientes_detalle)
        pen_txt = (
            f"Al {data['fecha_corte_larga']} queda 1 solicitud pendiente, por los siguientes motivos:"
            if cnt_pen == 1
            else f"Al {data['fecha_corte_larga']} quedan {cnt_pen} solicitudes pendientes, por los siguientes motivos:"
        )
        pen_chunks = [pendientes_detalle[i : i + 6] for i in range(0, cnt_pen, 6)]
        for c_idx, chunk in enumerate(pen_chunks):
            s_pen = prs.slides.add_slide(blank_layout)
            sfx = f" ({c_idx + 1}/{len(pen_chunks)})" if len(pen_chunks) > 1 else ""
            _add_white_slide_chrome(
                s_pen,
                title=f"4. Solicitudes pendientes y motivo{sfx}",
                subtitle=subtitle_ctx,
                page_label="Sección 4",
            )
            cur_y = 118.0
            if c_idx == 0:
                _add_textbox(
                    s_pen,
                    x_pt=54,
                    y_pt=cur_y,
                    w_pt=1338,
                    h_pt=32,
                    text=pen_txt,
                    size_pt=15.0,
                    bold=True,
                    color=PPT_NAVY,
                )
                cur_y += 38.0
                _, res_h = _create_styled_table(
                    s_pen,
                    x_pt=54,
                    y_pt=cur_y,
                    col_widths_pt=[980.0, 358.0],
                    headers=["Motivo", "Solicitudes"],
                    col_alignments=[PP_ALIGN.LEFT, PP_ALIGN.CENTER],
                    rows_data=[[r["motivo"], r["cantidad"]] for r in data["pendientes_resumen_rows"][:5]],
                    header_fill=PPT_ROYAL,
                    font_size_pt=12.5,
                    row_height_pt=32.0,
                    header_height_pt=36.0,
                )
                cur_y += res_h + 22.0

            _create_styled_table(
                s_pen,
                x_pt=54,
                y_pt=cur_y,
                col_widths_pt=[160.0, 150.0, 240.0, 380.0, 408.0],
                headers=["Solicitud", "Recibida", "Tipo", "Tomador / Asegurado", "Motivo"],
                col_alignments=[
                    PP_ALIGN.CENTER,
                    PP_ALIGN.CENTER,
                    PP_ALIGN.LEFT,
                    PP_ALIGN.LEFT,
                    PP_ALIGN.LEFT,
                ],
                rows_data=[
                    [
                        r["id"],
                        r["recibida"],
                        r["tipo"],
                        r["tomador_asegurado"],
                        r["motivo_pendiente"],
                    ]
                    for r in chunk
                ],
                header_fill=PPT_NAVY,
                font_size_pt=11.5,
                row_height_pt=40.0,
                header_height_pt=38.0,
                bold_cols={0},
            )

    # ================================================================== #
    # SLIDE 7: 5. TIEMPOS DE ATENCIÓN (Fondo blanco)                     #
    # ================================================================== #
    s7 = prs.slides.add_slide(blank_layout)
    _add_white_slide_chrome(
        s7,
        title="5. Tiempos de atención",
        subtitle=subtitle_ctx,
        page_label="Sección 5",
    )

    # Intro paragraph
    _add_rect(s7, 54, 120, 1338, 68, fill_color=PPT_ICE_CARD, border_color=PPT_CYAN, border_pt=1.2, rounded=True)
    _add_rect(s7, 54, 120, 8, 68, fill_color=PPT_ROYAL)
    _add_textbox(
        s7,
        x_pt=76,
        y_pt=126,
        w_pt=1298,
        h_pt=56,
        text=(
            "El tiempo de atención se mide desde que la solicitud ingresa a Multiriesgos de Colombia "
            "hasta que se cierra (póliza emitida, novedad atendida o solicitud declinada). Incluye los "
            "tiempos de espera de terceros, como la respuesta de la aseguradora o la entrega de documentos "
            "por parte del cliente."
        ),
        size_pt=13.5,
        color=PPT_NAVY,
        v_anchor=MSO_ANCHOR.MIDDLE,
    )

    tm = data["tiempos_medidas"]
    # Left Column: Medidas Table + Promedio por Tipo Table
    _, med_h = _create_styled_table(
        s7,
        x_pt=54,
        y_pt=208,
        col_widths_pt=[356.0, 290.0],
        headers=["Medida", "Resultado"],
        col_alignments=[PP_ALIGN.LEFT, PP_ALIGN.CENTER],
        rows_data=[
            ["Solicitudes cerradas en el período", tm["cerradas"]],
            ["Promedio", tm["promedio"]],
            ["Mediana", tm["mediana"]],
            ["Más rápida / más lenta", tm["extremos"]],
        ],
        header_fill=PPT_NAVY,
        font_size_pt=13.0,
        row_height_pt=38.0,
        header_height_pt=40.0,
        bold_cols={0, 1},
    )

    if data["tiempos_por_tipo_rows"]:
        tipo_y = 208.0 + med_h + 24.0
        _create_styled_table(
            s7,
            x_pt=54,
            y_pt=tipo_y,
            col_widths_pt=[436.0, 210.0],
            headers=["Promedio por tipo de solicitud", "Días"],
            col_alignments=[PP_ALIGN.LEFT, PP_ALIGN.CENTER],
            rows_data=[[r["tipo"], r["promedio"]] for r in data["tiempos_por_tipo_rows"]],
            header_fill=PPT_ROYAL,
            font_size_pt=13.0,
            row_height_pt=36.0,
            header_height_pt=40.0,
            bold_cols={1},
        )

    # Right Column: Distribución por Tiempo de Atención Table + Summary Cards
    _, tramos_h = _create_styled_table(
        s7,
        x_pt=732,
        y_pt=208,
        col_widths_pt=[430.0, 230.0],
        headers=["Distribución por tiempo de atención", "Solicitudes"],
        col_alignments=[PP_ALIGN.LEFT, PP_ALIGN.CENTER],
        rows_data=[[r["tramo"], r["solicitudes"]] for r in data["tiempos_tramos_rows"]],
        header_fill=PPT_NAVY,
        font_size_pt=13.0,
        row_height_pt=42.0,
        header_height_pt=40.0,
        bold_cols={1},
    )

    # Highlight cards under Right Column
    cards_y = 208.0 + tramos_h + 28.0
    if cards_y < 620.0:
        mini_cards = [
            ("PROMEDIO", kpis["tiempo_promedio_str"], PPT_CYAN),
            ("MEDIANA", tm["mediana"], PPT_SKY),
            ("RANGO (MÍN / MÁX)", tm["extremos"], PPT_ROYAL),
        ]
        mc_w = 208.0
        mc_gap = 18.0
        for m_idx, (m_title, m_val, m_clr) in enumerate(mini_cards):
            mx = 732.0 + m_idx * (mc_w + mc_gap)
            _add_rect(s7, mx, cards_y, mc_w, 125, fill_color=PPT_LIGHT_BG, border_color=m_clr, border_pt=1.5, rounded=True)
            _add_rect(s7, mx, cards_y, mc_w, 8, fill_color=m_clr)
            _add_textbox(
                s7,
                x_pt=mx + 6,
                y_pt=cards_y + 18,
                w_pt=mc_w - 12,
                h_pt=24,
                text=m_title,
                size_pt=11.0,
                bold=True,
                color=PPT_ROYAL,
                align=PP_ALIGN.CENTER,
            )
            _add_textbox(
                s7,
                x_pt=mx + 6,
                y_pt=cards_y + 48,
                w_pt=mc_w - 12,
                h_pt=60,
                text=m_val,
                size_pt=20.0 if len(m_val) < 12 else 16.0,
                bold=True,
                color=PPT_NAVY,
                font_name=FONT_HEADING,
                align=PP_ALIGN.CENTER,
                v_anchor=MSO_ANCHOR.MIDDLE,
            )

    # ================================================================== #
    # SLIDES 8..K: 6. DETALLE DE SOLICITUDES (Paginated, Fondo blanco)   #
    # ================================================================== #
    detalle_rows = data["detalle_solicitudes_rows"]
    # Dynamic chunking based on row text length (consortium members take extra vertical space)
    det_chunks: list[list[dict[str, Any]]] = []
    current_chunk: list[dict[str, Any]] = []
    current_weight = 0.0
    MAX_SLIDE_WEIGHT = 9.5

    for r in detalle_rows:
        t_len = len(r.get("tomador_tabla8", ""))
        row_weight = 1.8 if ("\n" in r.get("tomador_tabla8", "") or t_len > 65) else 1.0
        if current_chunk and (current_weight + row_weight > MAX_SLIDE_WEIGHT):
            det_chunks.append(current_chunk)
            current_chunk = [r]
            current_weight = row_weight
        else:
            current_chunk.append(r)
            current_weight += row_weight
    if current_chunk:
        det_chunks.append(current_chunk)
    if not det_chunks:
        det_chunks = [[]]

    for page_idx, chunk in enumerate(det_chunks):
        s_det = prs.slides.add_slide(blank_layout)
        page_suffix = f" ({page_idx + 1}/{len(det_chunks)})" if len(det_chunks) > 1 else ""
        _add_white_slide_chrome(
            s_det,
            title=f"6. Detalle de solicitudes{page_suffix}",
            subtitle=subtitle_ctx,
            page_label=f"Sección 6 · Pág. {page_idx + 1}/{len(det_chunks)}",
        )

        _create_styled_table(
            s_det,
            x_pt=54,
            y_pt=116,
            col_widths_pt=[125.0, 120.0, 210.0, 365.0, 288.0, 120.0, 110.0],
            headers=["Solicitud", "Recibida", "Tipo", "Tomador", "Asegurado", "Estado", "Días"],
            col_alignments=[
                PP_ALIGN.CENTER,
                PP_ALIGN.CENTER,
                PP_ALIGN.LEFT,
                PP_ALIGN.LEFT,
                PP_ALIGN.LEFT,
                PP_ALIGN.CENTER,
                PP_ALIGN.CENTER,
            ],
            rows_data=[
                [
                    r["id"],
                    r["recibida"],
                    r["tipo_detalle"],
                    r["tomador_tabla8"],
                    r["asegurado_fmt"],
                    r["estado_label"],
                    r["dias_fmt"],
                ]
                for r in chunk
            ],
            header_fill=PPT_NAVY,
            font_size_pt=10.5,
            row_height_pt=34.0,
            header_height_pt=36.0,
            bold_cols={0, 5},
        )

        _add_textbox(
            s_det,
            x_pt=54,
            y_pt=742,
            w_pt=1338,
            h_pt=30,
            text=(
                "El nombre de las columnas “Tomador” y “Asegurado” corresponde al remitente y a la entidad "
                "registrados en cada solicitud. Cuando el cliente actúa como integrante de un consorcio o "
                "unión temporal, se indican los integrantes registrados."
            ),
            size_pt=10.5,
            color=PPT_MUTED_DARK,
        )

    # ================================================================== #
    # SLIDE K+1: 7. NOTAS METODOLÓGICAS (Fondo blanco, sin imágenes)     #
    # ================================================================== #
    s_notes = prs.slides.add_slide(blank_layout)
    _add_white_slide_chrome(
        s_notes,
        title="7. Notas metodológicas",
        subtitle="Criterios técnicos y alcance de las métricas del informe",
        page_label="Sección 7",
    )

    pagare_cnt = data.get("pending_pagare_cartera_count", 0)
    pagare_suffix = (
        "en este período no hubo solicitudes en esa condición."
        if pagare_cnt == 0
        else f"en este período hubo {pagare_cnt} {'solicitud' if pagare_cnt == 1 else 'solicitudes'} en esa condición."
    )

    notes_list = [
        (
            "Universo del informe",
            f"Universo: solicitudes {data['universo_periodo_texto']} en las que el cliente aparece como "
            f"tomador, como asegurado o como integrante de un consorcio o unión temporal.",
            PPT_GREEN,
        ),
        (
            "Fecha de corte",
            f"Corte: el estado de cada solicitud corresponde al {data['fecha_corte_larga']}; las "
            f"solicitudes pendientes pueden cerrarse después de esa fecha.",
            PPT_ROYAL,
        ),
        (
            "Efectividad de atención",
            "Índice de efectividad de atención = (pólizas emitidas + novedades atendidas) ÷ solicitudes tramitadas.",
            PPT_CYAN,
        ),
        (
            "Índice de gestión",
            "Índice de gestión = (pólizas emitidas + novedades atendidas + solicitudes declinadas) ÷ solicitudes tramitadas.",
            PPT_GREEN,
        ),
        (
            "Pagarés y estado de cartera",
            f"Los pagarés pendientes y los estados de cartera se informan como motivo de pendiente cuando "
            f"la solicitud los requiere; {pagare_suffix}",
            PPT_ROYAL,
        ),
    ]

    note_start_y = 132.0
    note_step_y = 124.0
    for n_idx, (n_title, n_body, badge_clr) in enumerate(notes_list):
        ny = note_start_y + n_idx * note_step_y
        _add_rect(s_notes, 54, ny, 1338, 108, fill_color=PPT_LIGHT_BG, border_color=RGBColor(0xE2, 0xE8, 0xF0), border_pt=1.0, rounded=True)
        circle = s_notes.shapes.add_shape(MSO_SHAPE.OVAL, Pt(74), Pt(ny + 22), Pt(64), Pt(64))
        circle.fill.solid()
        circle.fill.fore_color.rgb = badge_clr
        circle.line.color.rgb = PPT_WHITE
        circle.line.width = Pt(1.5)
        _add_textbox(
            s_notes,
            x_pt=74,
            y_pt=ny + 34,
            w_pt=64,
            h_pt=40,
            text=f"0{n_idx + 1}",
            size_pt=20.0,
            bold=True,
            color=PPT_WHITE,
            font_name=FONT_HEADING,
            align=PP_ALIGN.CENTER,
        )
        _add_textbox(
            s_notes,
            x_pt=158,
            y_pt=ny + 14,
            w_pt=1200,
            h_pt=28,
            text=n_title,
            size_pt=17.0,
            bold=True,
            color=PPT_NAVY,
            font_name=FONT_HEADING,
        )
        _add_textbox(
            s_notes,
            x_pt=158,
            y_pt=ny + 46,
            w_pt=1200,
            h_pt=50,
            text=n_body,
            size_pt=14.0,
            bold=False,
            color=PPT_DARK_TEXT,
            font_name=FONT_BODY,
        )

    # ================================================================== #
    # SLIDE K+2: GRACIAS / FIN (Fondo azul corporativo, sin imágenes)    #
    # ================================================================== #
    s_end = prs.slides.add_slide(blank_layout)
    _set_solid_slide_bg(s_end, PPT_NAVY)

    # Vector geometric accents (zero images)
    _add_rect(s_end, 0, 0, 18, SLIDE_H_PT, fill_color=PPT_ROYAL)
    _add_rect(s_end, 18, 0, 5, SLIDE_H_PT, fill_color=PPT_CYAN)
    _add_diamond_accents(s_end, origin_x=1220, origin_y=52, color=RGBColor(0x0D, 0x35, 0x68))
    _add_diamond_accents(s_end, origin_x=1220, origin_y=640, color=PPT_ROYAL)

    _add_textbox(
        s_end,
        x_pt=96,
        y_pt=96,
        w_pt=700,
        h_pt=40,
        text="Construimos el futuro",
        size_pt=26.0,
        italic=True,
        color=PPT_SKY,
        font_name=FONT_BODY,
    )
    _add_textbox(
        s_end,
        x_pt=96,
        y_pt=134,
        w_pt=700,
        h_pt=40,
        text="Asegurando el presente",
        size_pt=26.0,
        italic=True,
        color=PPT_WHITE,
        font_name=FONT_BODY,
    )

    _add_textbox(
        s_end,
        x_pt=96,
        y_pt=248,
        w_pt=800,
        h_pt=115,
        text="GRACIAS",
        size_pt=72.0,
        bold=True,
        color=PPT_WHITE,
        font_name=FONT_HEADING,
    )
    _add_rect(s_end, 96, 376, 180, 4, fill_color=PPT_CYAN)

    _add_textbox(
        s_end,
        x_pt=96,
        y_pt=402,
        w_pt=900,
        h_pt=46,
        text="Si tiene alguna pregunta póngase en contacto con nosotros:",
        size_pt=22.0,
        bold=False,
        color=PPT_ICE_CARD,
        font_name=FONT_BODY,
    )

    _add_rect(s_end, 96, 480, 720, 152, fill_color=PPT_NAVY_PANEL, border_color=PPT_ROYAL, border_pt=1.2, rounded=True)
    _add_rect(s_end, 96, 480, 8, 152, fill_color=PPT_CYAN)
    _add_textbox(
        s_end,
        x_pt=126,
        y_pt=498,
        w_pt=660,
        h_pt=38,
        text="Teléfono: 320 4982460",
        size_pt=22.0,
        bold=True,
        color=PPT_WHITE,
        font_name=FONT_BODY,
        v_anchor=MSO_ANCHOR.MIDDLE,
    )
    _add_textbox(
        s_end,
        x_pt=126,
        y_pt=542,
        w_pt=660,
        h_pt=38,
        text="Correo: santiagocardona@mrc.com.co",
        size_pt=22.0,
        bold=False,
        color=PPT_SKY,
        font_name=FONT_BODY,
        v_anchor=MSO_ANCHOR.MIDDLE,
    )
    _add_textbox(
        s_end,
        x_pt=126,
        y_pt=586,
        w_pt=660,
        h_pt=32,
        text="Multiriesgos de Colombia Ltda.",
        size_pt=14.0,
        bold=False,
        color=PPT_MUTED_LIGHT,
        font_name=FONT_BODY,
        v_anchor=MSO_ANCHOR.MIDDLE,
    )

    output_path_str = str(output_path)
    prs.save(output_path_str)
    logger.info("Presentación de Informe de Gestión (.pptx) generada exitosamente en: %s", output_path_str)
    return output_path_str


# --------------------------------------------------------------------------- #
# Word (.docx) Renderer matching 'Informe_Gestion_SONACOL_2026-09.docx'       #
# --------------------------------------------------------------------------- #
from docx import Document as DocxDocument
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement as DocxOxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn as docx_qn
from docx.shared import Inches as DocxInches, Pt as DocxPt, RGBColor as DocxRGBColor

HEX_NAVY = "1F3864"
HEX_GRAY_TEXT = "595959"
HEX_FOOTER_GRAY = "7F7F7F"
HEX_BORDER_GRAY = "BFBFBF"
HEX_ZEBRA_FILL = "F2F2F2"
HEX_TOTAL_FILL = "D9E2F3"
HEX_WHITE = "FFFFFF"

RGB_NAVY = DocxRGBColor(0x1F, 0x38, 0x64)
RGB_GRAY_TEXT = DocxRGBColor(0x59, 0x59, 0x59)
RGB_FOOTER_GRAY = DocxRGBColor(0x7F, 0x7F, 0x7F)
RGB_WHITE = DocxRGBColor(0xFF, 0xFF, 0xFF)
RGB_BLACK = DocxRGBColor(0x00, 0x00, 0x00)


def _docx_set_run_font(
    run: Any,
    size_pt: float,
    bold: bool = False,
    italic: bool = False,
    color: DocxRGBColor | None = None,
    font_name: str = "Arial",
) -> None:
    """Apply consistent Arial typography to a Word run."""
    run.font.name = font_name
    run.font.size = DocxPt(size_pt)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(docx_qn("w:rFonts"))
    if rFonts is None:
        rFonts = DocxOxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(docx_qn("w:ascii"), font_name)
    rFonts.set(docx_qn("w:hAnsi"), font_name)
    rFonts.set(docx_qn("w:cs"), font_name)
    rFonts.set(docx_qn("w:eastAsia"), font_name)


def _docx_set_paragraph_spacing(
    p: Any,
    before_twips: int | None = None,
    after_twips: int | None = None,
    line_twips: int | None = None,
) -> None:
    """Set exact twip spacing on a Word paragraph (<w:spacing>)."""
    pPr = p._p.get_or_add_pPr()
    spacing = pPr.find(docx_qn("w:spacing"))
    if spacing is None:
        spacing = DocxOxmlElement("w:spacing")
        pPr.append(spacing)
    if before_twips is not None:
        spacing.set(docx_qn("w:before"), str(before_twips))
    if after_twips is not None:
        spacing.set(docx_qn("w:after"), str(after_twips))
    if line_twips is not None:
        spacing.set(docx_qn("w:line"), str(line_twips))


def _docx_set_paragraph_bottom_border(p: Any, color_hex: str = HEX_NAVY, sz: int = 8, space: int = 2) -> None:
    """Add a bottom border to a Word paragraph (used for Heading 1 and Header)."""
    pPr = p._p.get_or_add_pPr()
    pBdr = pPr.find(docx_qn("w:pBdr"))
    if pBdr is None:
        pBdr = DocxOxmlElement("w:pBdr")
        pPr.append(pBdr)
    bottom = DocxOxmlElement("w:bottom")
    bottom.set(docx_qn("w:val"), "single")
    bottom.set(docx_qn("w:sz"), str(sz))
    bottom.set(docx_qn("w:space"), str(space))
    bottom.set(docx_qn("w:color"), color_hex)
    pBdr.append(bottom)


def _docx_add_section_heading(doc: Any, title: str) -> Any:
    """Add a numbered section heading matching Heading 1 in reference docx."""
    p = doc.add_paragraph()
    try:
        p.style = doc.styles["Heading 1"]
    except KeyError:
        pass
    p.paragraph_format.keep_with_next = True
    _docx_set_paragraph_bottom_border(p, color_hex=HEX_NAVY, sz=8, space=2)
    _docx_set_paragraph_spacing(p, before_twips=300, after_twips=120)
    run = p.add_run(title)
    _docx_set_run_font(run, size_pt=13.0, bold=True, color=RGB_NAVY)
    return p


def _docx_add_spacer_paragraph(doc: Any, after_twips: int = 80) -> Any:
    """Add a compact spacer paragraph between tables."""
    p = doc.add_paragraph()
    _docx_set_paragraph_spacing(p, before_twips=0, after_twips=after_twips, line_twips=276)
    return p


def _docx_configure_header_and_footer(doc: Any, assets_dir: str | Path | None = None) -> None:
    """Configure the corporate header (logo + right title + navy rule) and footer (page X of Y)."""
    section = doc.sections[0]
    section.top_margin = DocxInches(1.0417)
    section.bottom_margin = DocxInches(0.8333)
    section.left_margin = DocxInches(1.0)
    section.right_margin = DocxInches(1.0)
    section.page_width = DocxInches(8.5)
    section.page_height = DocxInches(11.0)

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    hp.text = ""
    _docx_set_paragraph_bottom_border(hp, color_hex=HEX_NAVY, sz=6, space=4)
    _docx_set_paragraph_spacing(hp, after_twips=80)

    pPr = hp._p.get_or_add_pPr()
    tabs = DocxOxmlElement("w:tabs")
    tab = DocxOxmlElement("w:tab")
    tab.set(docx_qn("w:val"), "right")
    tab.set(docx_qn("w:pos"), "9360")
    tabs.append(tab)
    pPr.append(tabs)

    logo_path: Path | None = None
    if assets_dir:
        candidate = Path(assets_dir) / "logo_mrc.png"
        if candidate.exists():
            logo_path = candidate

    if logo_path and logo_path.exists():
        r_logo = hp.add_run()
        r_logo.add_picture(str(logo_path), width=DocxInches(1.354), height=DocxInches(0.479))

    r_title = hp.add_run("\tInforme de gestión de solicitudes")
    _docx_set_run_font(r_title, size_pt=9.0, bold=False, color=RGB_GRAY_TEXT)

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fp.text = ""
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    r1 = fp.add_run("Multiriesgos de Colombia Ltda. · Página ")
    _docx_set_run_font(r1, size_pt=8.0, color=RGB_FOOTER_GRAY)

    def _add_xml_field(paragraph: Any, field_name: str) -> None:
        r_fld = paragraph.add_run()
        _docx_set_run_font(r_fld, size_pt=8.0, color=RGB_FOOTER_GRAY)
        fld_begin = DocxOxmlElement("w:fldChar")
        fld_begin.set(docx_qn("w:fldCharType"), "begin")
        instr = DocxOxmlElement("w:instrText")
        instr.set(docx_qn("xml:space"), "preserve")
        instr.text = field_name
        fld_sep = DocxOxmlElement("w:fldChar")
        fld_sep.set(docx_qn("w:fldCharType"), "separate")
        fld_end = DocxOxmlElement("w:fldChar")
        fld_end.set(docx_qn("w:fldCharType"), "end")
        r_fld._r.append(fld_begin)
        r_fld._r.append(instr)
        r_fld._r.append(fld_sep)
        r_fld._r.append(fld_end)

    _add_xml_field(fp, "PAGE")
    r2 = fp.add_run(" de ")
    _docx_set_run_font(r2, size_pt=8.0, color=RGB_FOOTER_GRAY)
    _add_xml_field(fp, "NUMPAGES")


def _docx_set_table_grid_and_width(table: Any, col_widths_dxa: list[int]) -> None:
    """Configure exact dxa column widths on tblPr and tblGrid."""
    total_w = sum(col_widths_dxa)
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblW = tblPr.find(docx_qn("w:tblW"))
    if tblW is None:
        tblW = DocxOxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(docx_qn("w:type"), "dxa")
    tblW.set(docx_qn("w:w"), str(total_w))

    tblGrid = tbl.find(docx_qn("w:tblGrid"))
    if tblGrid is not None:
        tbl.remove(tblGrid)
    tblGrid = DocxOxmlElement("w:tblGrid")
    for w in col_widths_dxa:
        gridCol = DocxOxmlElement("w:gridCol")
        gridCol.set(docx_qn("w:w"), str(w))
        tblGrid.append(gridCol)
    tbl.insert(tbl.index(tblPr) + 1, tblGrid)


def _docx_style_cell_box(
    cell: Any,
    width_dxa: int,
    fill_hex: str | None = None,
    border_color: str = HEX_BORDER_GRAY,
    border_sz: int = 4,
    pad_top: int = 50,
    pad_bottom: int = 50,
    pad_left: int = 90,
    pad_right: int = 90,
    is_kpi_card: bool = False,
) -> None:
    """Apply exact borders, background shading, width, margins, and vertical centering to a Word cell."""
    tcPr = cell._tc.get_or_add_tcPr()

    tcW = tcPr.find(docx_qn("w:tcW"))
    if tcW is None:
        tcW = DocxOxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(docx_qn("w:type"), "dxa")
    tcW.set(docx_qn("w:w"), str(width_dxa))

    tcBorders = tcPr.find(docx_qn("w:tcBorders"))
    if tcBorders is not None:
        tcPr.remove(tcBorders)
    if is_kpi_card:
        borders_xml = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="none" w:color="FFFFFF" w:sz="0"/>'
            f'<w:left w:val="single" w:color="FFFFFF" w:sz="24"/>'
            f'<w:bottom w:val="none" w:color="FFFFFF" w:sz="0"/>'
            f'<w:right w:val="single" w:color="FFFFFF" w:sz="24"/>'
            f"</w:tcBorders>"
        )
    else:
        borders_xml = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:color="{border_color}" w:sz="{border_sz}"/>'
            f'<w:left w:val="single" w:color="{border_color}" w:sz="{border_sz}"/>'
            f'<w:bottom w:val="single" w:color="{border_color}" w:sz="{border_sz}"/>'
            f'<w:right w:val="single" w:color="{border_color}" w:sz="{border_sz}"/>'
            f"</w:tcBorders>"
        )
    tcPr.append(borders_xml)

    shd = tcPr.find(docx_qn("w:shd"))
    if shd is not None:
        tcPr.remove(shd)
    if fill_hex:
        shd_xml = parse_xml(
            f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill_hex}"/>'
        )
        tcPr.append(shd_xml)

    tcMar = tcPr.find(docx_qn("w:tcMar"))
    if tcMar is not None:
        tcPr.remove(tcMar)
    mar_xml = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:type="dxa" w:w="{pad_top}"/>'
        f'<w:left w:type="dxa" w:w="{pad_left}"/>'
        f'<w:bottom w:type="dxa" w:w="{pad_bottom}"/>'
        f'<w:right w:type="dxa" w:w="{pad_right}"/>'
        f"</w:tcMar>"
    )
    tcPr.append(mar_xml)

    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _docx_populate_standard_table(
    doc: Any,
    col_widths_dxa: list[int],
    headers: list[str],
    col_alignments: list[WD_ALIGN_PARAGRAPH],
    rows_data: list[list[str]],
    font_size_pt: float = 8.5,
    bold_cols: set[int] | None = None,
    has_total_row: bool = False,
) -> Any:
    """Create and populate a standard corporate Word table matching the SONACOL report."""
    bold_cols = bold_cols or set()
    total_rows = 1 + len(rows_data)
    table = doc.add_table(rows=total_rows, cols=len(col_widths_dxa))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _docx_set_table_grid_and_width(table, col_widths_dxa)

    hdr_cells = table.rows[0].cells
    for c_idx, h_text in enumerate(headers):
        cell = hdr_cells[c_idx]
        _docx_style_cell_box(
            cell,
            width_dxa=col_widths_dxa[c_idx],
            fill_hex=HEX_NAVY,
        )
        p = cell.paragraphs[0]
        p.text = ""
        p.alignment = col_alignments[c_idx] if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
        if len(headers) <= 3 and c_idx == len(headers) - 1 and col_alignments[c_idx] == WD_ALIGN_PARAGRAPH.LEFT:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        _docx_set_paragraph_spacing(p, before_twips=30, after_twips=30)
        run = p.add_run(h_text)
        _docx_set_run_font(run, size_pt=font_size_pt, bold=True, color=RGB_WHITE)

    for r_offset, row_vals in enumerate(rows_data):
        r_idx = r_offset + 1
        is_last = r_offset == len(rows_data) - 1
        is_total = has_total_row and is_last

        if is_total:
            fill_hex = HEX_TOTAL_FILL
        elif r_idx % 2 == 0:
            fill_hex = HEX_ZEBRA_FILL
        else:
            fill_hex = None

        row_cells = table.rows[r_idx].cells
        for c_idx, val_text in enumerate(row_vals):
            cell = row_cells[c_idx]
            _docx_style_cell_box(
                cell,
                width_dxa=col_widths_dxa[c_idx],
                fill_hex=fill_hex,
            )
            lines = str(val_text).split("\n")
            for l_idx, line_str in enumerate(lines):
                p = cell.paragraphs[0] if l_idx == 0 else cell.add_paragraph()
                if l_idx == 0:
                    p.text = ""
                p.alignment = col_alignments[c_idx]
                _docx_set_paragraph_spacing(p, before_twips=30, after_twips=30)

                is_bold = is_total or (c_idx in bold_cols and l_idx == 0)
                if l_idx > 0 and line_str.startswith("Integrantes:"):
                    run = p.add_run(line_str)
                    _docx_set_run_font(
                        run,
                        size_pt=max(7.0, font_size_pt - 0.5),
                        bold=False,
                        color=RGB_GRAY_TEXT,
                    )
                else:
                    run = p.add_run(line_str)
                    _docx_set_run_font(
                        run,
                        size_pt=font_size_pt,
                        bold=is_bold,
                        color=RGB_BLACK,
                    )

    return table


def _docx_add_bullet_note(doc: Any, text: str) -> Any:
    """Add a methodological bullet point paragraph matching Section 7 (10pt Arial)."""
    p = doc.add_paragraph()
    try:
        p.style = doc.styles["List Bullet"]
    except KeyError:
        try:
            p.style = doc.styles["List Paragraph"]
        except KeyError:
            pass
    _docx_set_paragraph_spacing(p, after_twips=60, line_twips=276)
    r_txt = p.add_run(text)
    _docx_set_run_font(r_txt, size_pt=10.0, bold=False, color=RGB_BLACK)
    return p


def render_informe_gestion_docx(
    data: dict[str, Any],
    output_path: str | Path,
    assets_dir: str | Path | None = None,
) -> str:
    """Render the complete Management Report (.docx) matching Informe_Gestion_SONACOL_2026-09.docx."""
    doc = DocxDocument()
    _docx_configure_header_and_footer(doc, assets_dir=assets_dir)

    # Title Block (Paragraphs 0..3)
    p0 = doc.add_paragraph()
    _docx_set_paragraph_spacing(p0, before_twips=120, after_twips=40)
    r0 = p0.add_run(f"Informe de gestión · {data['periodo_titulo']}")
    _docx_set_run_font(r0, size_pt=11.0, bold=False, color=RGB_GRAY_TEXT)

    p1 = doc.add_paragraph()
    _docx_set_paragraph_spacing(p1, after_twips=40)
    r1 = p1.add_run(data["cliente_nombre_completo"])
    _docx_set_run_font(r1, size_pt=20.0, bold=True, color=RGB_NAVY)

    p2 = doc.add_paragraph()
    _docx_set_paragraph_spacing(p2, before_twips=0, after_twips=60, line_twips=276)
    nit_prefix = f"NIT {data['cliente_nit']} · " if data.get("cliente_nit") else ""
    r2 = p2.add_run(
        f"{nit_prefix}Incluye solicitudes como tomador, asegurado e integrante de consorcios o uniones temporales"
    )
    _docx_set_run_font(r2, size_pt=9.5, bold=False, color=RGB_GRAY_TEXT)

    p3 = doc.add_paragraph()
    _docx_set_paragraph_spacing(p3, before_twips=0, after_twips=200, line_twips=276)
    r3 = p3.add_run(
        f"Ramo: {data['ramo_nombre']} · Solicitudes recibidas {data['rango_fechas_texto']} · "
        f"Estado de las solicitudes al {data['fecha_corte_larga']}"
    )
    _docx_set_run_font(r3, size_pt=9.0, bold=False, color=RGB_GRAY_TEXT)

    # 1. Resumen del mes
    _docx_add_section_heading(doc, "1. Resumen del mes")

    kpis = data["kpis"]
    kpi_items = [
        (str(kpis["tramitadas"]), "Tramitadas"),
        (str(kpis["emitidas"]), "Emitidas"),
        (str(kpis["novedades"]), "Novedades atendidas"),
        (str(kpis["declinadas"]), "Declinadas"),
        (str(kpis["pendientes"]), "Pendientes"),
    ]
    kpi_widths = [1872, 1872, 1872, 1872, 1872]
    tbl_kpi = doc.add_table(rows=1, cols=5)
    tbl_kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
    _docx_set_table_grid_and_width(tbl_kpi, kpi_widths)
    for c_idx, (num_str, label_str) in enumerate(kpi_items):
        cell = tbl_kpi.rows[0].cells[c_idx]
        _docx_style_cell_box(
            cell,
            width_dxa=kpi_widths[c_idx],
            fill_hex=HEX_NAVY,
            pad_top=120,
            pad_bottom=120,
            pad_left=80,
            pad_right=80,
            is_kpi_card=True,
        )
        cp0 = cell.paragraphs[0]
        cp0.text = ""
        cp0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _docx_set_paragraph_spacing(cp0, after_twips=0)
        rn = cp0.add_run(num_str)
        _docx_set_run_font(rn, size_pt=26.0, bold=True, color=RGB_WHITE)

        cp1 = cell.add_paragraph()
        cp1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _docx_set_paragraph_spacing(cp1, after_twips=0)
        rl = cp1.add_run(label_str)
        _docx_set_run_font(rl, size_pt=8.5, bold=False, color=RGB_WHITE)

    _docx_add_spacer_paragraph(doc, after_twips=80)

    _docx_populate_standard_table(
        doc,
        col_widths_dxa=[3000, 1500, 4860],
        headers=["Indicador", "Resultado", "Cómo se calcula / alcance"],
        col_alignments=[
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
        ],
        rows_data=[
            [
                "Índice de efectividad de atención",
                kpis["efectividad_str"],
                kpis["efectividad_detalle"],
            ],
            [
                "Índice de gestión",
                kpis["gestion_str"],
                kpis["gestion_detalle"],
            ],
            [
                "Tiempo promedio de atención",
                kpis["tiempo_promedio_str"],
                kpis["tiempo_promedio_detalle"],
            ],
        ],
        font_size_pt=9.0,
        bold_cols={0, 1},
    )

    _docx_add_spacer_paragraph(doc, after_twips=60)

    p_res = doc.add_paragraph()
    _docx_set_paragraph_spacing(p_res, before_twips=60, after_twips=120, line_twips=276)
    r_res = p_res.add_run(data["resumen_narrativa"])
    _docx_set_run_font(r_res, size_pt=10.0, bold=False, color=RGB_BLACK)

    # 2. Trazabilidad de las solicitudes
    _docx_add_section_heading(doc, "2. Trazabilidad de las solicitudes")

    traz_rows_matrix: list[list[str]] = []
    for r in data["trazabilidad_rows"]:
        traz_rows_matrix.append(
            [
                r["tipo"],
                r["tramitadas"],
                r["emitidas"],
                r["novedades"],
                r["declinadas"],
                r["pendientes"],
                r["efectividad"],
                r["dias_prom"],
            ]
        )
    tot = data["trazabilidad_total"]
    traz_rows_matrix.append(
        [
            tot["tipo"],
            tot["tramitadas"],
            tot["emitidas"],
            tot["novedades"],
            tot["declinadas"],
            tot["pendientes"],
            tot["efectividad"],
            tot["dias_prom"],
        ]
    )

    _docx_populate_standard_table(
        doc,
        col_widths_dxa=[1900, 1050, 1000, 1200, 1050, 1050, 1160, 950],
        headers=[
            "Tipo de solicitud",
            "Tramitadas",
            "Emitidas",
            "Novedades atendidas",
            "Declinadas",
            "Pendientes",
            "Efectividad",
            "Días prom.",
        ],
        col_alignments=[
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
        ],
        rows_data=traz_rows_matrix,
        font_size_pt=8.0,
        has_total_row=True,
    )

    p_traz_note = doc.add_paragraph()
    _docx_set_paragraph_spacing(p_traz_note, before_twips=80, after_twips=120, line_twips=276)
    r_traz_note = p_traz_note.add_run(
        "Emitidas: pólizas nuevas expedidas. Novedades atendidas: modificaciones "
        "(prórrogas, adiciones, aclaraciones) ya resueltas. Días promedio: tiempo "
        "transcurrido entre la recepción de la solicitud y su cierre."
    )
    _docx_set_run_font(r_traz_note, size_pt=8.0, bold=False, color=RGB_GRAY_TEXT)

    # 3. Solicitudes declinadas y motivo
    _docx_add_section_heading(doc, "3. Solicitudes declinadas y motivo")

    declinadas_detalle = data["declinadas_detalle_rows"]
    if not declinadas_detalle:
        p_dec = doc.add_paragraph()
        _docx_set_paragraph_spacing(p_dec, before_twips=0, after_twips=120, line_twips=276)
        r_dec = p_dec.add_run(
            f"{data['periodo_narrativa']} no se declinó ninguna solicitud."
        )
        _docx_set_run_font(r_dec, size_pt=10.0, bold=False, color=RGB_BLACK)
    else:
        cnt_dec = len(declinadas_detalle)
        p_dec = doc.add_paragraph()
        _docx_set_paragraph_spacing(p_dec, before_twips=0, after_twips=120, line_twips=276)
        dec_txt = (
            f"{data['periodo_narrativa']} se declinó 1 solicitud, por los siguientes motivos:"
            if cnt_dec == 1
            else f"{data['periodo_narrativa']} se declinaron {cnt_dec} solicitudes, por los siguientes motivos:"
        )
        r_dec = p_dec.add_run(dec_txt)
        _docx_set_run_font(r_dec, size_pt=10.0, bold=False, color=RGB_BLACK)

        _docx_populate_standard_table(
            doc,
            col_widths_dxa=[6360, 3000],
            headers=["Motivo", "Solicitudes"],
            col_alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
            rows_data=[[r["motivo"], r["cantidad"]] for r in data["declinadas_resumen_rows"]],
            font_size_pt=8.5,
        )
        _docx_add_spacer_paragraph(doc, after_twips=80)
        _docx_populate_standard_table(
            doc,
            col_widths_dxa=[1250, 1100, 1800, 2400, 2810],
            headers=["Solicitud", "Recibida", "Tipo", "Tomador / Asegurado", "Motivo"],
            col_alignments=[
                WD_ALIGN_PARAGRAPH.CENTER,
                WD_ALIGN_PARAGRAPH.CENTER,
                WD_ALIGN_PARAGRAPH.LEFT,
                WD_ALIGN_PARAGRAPH.LEFT,
                WD_ALIGN_PARAGRAPH.LEFT,
            ],
            rows_data=[
                [
                    r["id"],
                    r["recibida"],
                    r["tipo"],
                    r["tomador_asegurado"],
                    r["motivo_declinada"],
                ]
                for r in declinadas_detalle
            ],
            font_size_pt=8.0,
        )

    # 4. Solicitudes pendientes y motivo
    _docx_add_section_heading(doc, "4. Solicitudes pendientes y motivo")

    pendientes_detalle = data["pendientes_detalle_rows"]
    if not pendientes_detalle:
        p_pen = doc.add_paragraph()
        _docx_set_paragraph_spacing(p_pen, before_twips=0, after_twips=120, line_twips=276)
        r_pen = p_pen.add_run(
            f"Al {data['fecha_corte_larga']} no queda ninguna solicitud pendiente."
        )
        _docx_set_run_font(r_pen, size_pt=10.0, bold=False, color=RGB_BLACK)
    else:
        cnt_pen = len(pendientes_detalle)
        p_pen = doc.add_paragraph()
        _docx_set_paragraph_spacing(p_pen, before_twips=0, after_twips=120, line_twips=276)
        pen_txt = (
            f"Al {data['fecha_corte_larga']} queda 1 solicitud pendiente, por los siguientes motivos:"
            if cnt_pen == 1
            else f"Al {data['fecha_corte_larga']} quedan {cnt_pen} solicitudes pendientes, por los siguientes motivos:"
        )
        r_pen = p_pen.add_run(pen_txt)
        _docx_set_run_font(r_pen, size_pt=10.0, bold=False, color=RGB_BLACK)

        _docx_populate_standard_table(
            doc,
            col_widths_dxa=[6360, 3000],
            headers=["Motivo", "Solicitudes"],
            col_alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
            rows_data=[[r["motivo"], r["cantidad"]] for r in data["pendientes_resumen_rows"]],
            font_size_pt=8.5,
        )
        _docx_add_spacer_paragraph(doc, after_twips=80)
        _docx_populate_standard_table(
            doc,
            col_widths_dxa=[1250, 1100, 1800, 2400, 2810],
            headers=["Solicitud", "Recibida", "Tipo", "Tomador / Asegurado", "Motivo"],
            col_alignments=[
                WD_ALIGN_PARAGRAPH.CENTER,
                WD_ALIGN_PARAGRAPH.CENTER,
                WD_ALIGN_PARAGRAPH.LEFT,
                WD_ALIGN_PARAGRAPH.LEFT,
                WD_ALIGN_PARAGRAPH.LEFT,
            ],
            rows_data=[
                [
                    r["id"],
                    r["recibida"],
                    r["tipo"],
                    r["tomador_asegurado"],
                    r["motivo_pendiente"],
                ]
                for r in pendientes_detalle
            ],
            font_size_pt=8.0,
        )

    # 5. Tiempos de atención
    _docx_add_section_heading(doc, "5. Tiempos de atención")

    p_t = doc.add_paragraph()
    _docx_set_paragraph_spacing(p_t, before_twips=0, after_twips=120, line_twips=276)
    r_t = p_t.add_run(
        "El tiempo de atención se mide desde que la solicitud ingresa a Multiriesgos de Colombia "
        "hasta que se cierra (póliza emitida, novedad atendida o solicitud declinada). Incluye los "
        "tiempos de espera de terceros, como la respuesta de la aseguradora o la entrega de documentos "
        "por parte del cliente."
    )
    _docx_set_run_font(r_t, size_pt=10.0, bold=False, color=RGB_BLACK)

    tm = data["tiempos_medidas"]
    _docx_populate_standard_table(
        doc,
        col_widths_dxa=[4680, 4680],
        headers=["Medida", "Resultado"],
        col_alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
        rows_data=[
            ["Solicitudes cerradas en el período", tm["cerradas"]],
            ["Promedio", tm["promedio"]],
            ["Mediana", tm["mediana"]],
            ["Más rápida / más lenta", tm["extremos"]],
        ],
        font_size_pt=8.5,
        bold_cols={0},
    )

    _docx_add_spacer_paragraph(doc, after_twips=80)

    _docx_populate_standard_table(
        doc,
        col_widths_dxa=[6360, 3000],
        headers=["Distribución por tiempo de atención", "Solicitudes"],
        col_alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
        rows_data=[[r["tramo"], r["solicitudes"]] for r in data["tiempos_tramos_rows"]],
        font_size_pt=8.5,
    )

    if data["tiempos_por_tipo_rows"]:
        _docx_add_spacer_paragraph(doc, after_twips=80)
        _docx_populate_standard_table(
            doc,
            col_widths_dxa=[6360, 3000],
            headers=["Promedio por tipo de solicitud", "Días"],
            col_alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
            rows_data=[[r["tipo"], r["promedio"]] for r in data["tiempos_por_tipo_rows"]],
            font_size_pt=8.5,
        )

    # 6. Detalle de solicitudes
    _docx_add_section_heading(doc, "6. Detalle de solicitudes")

    _docx_populate_standard_table(
        doc,
        col_widths_dxa=[1150, 1000, 1450, 2200, 1850, 900, 810],
        headers=["Solicitud", "Recibida", "Tipo", "Tomador", "Asegurado", "Estado", "Días"],
        col_alignments=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
        ],
        rows_data=[
            [
                r["id"],
                r["recibida"],
                r["tipo_detalle"],
                r["tomador_tabla8"],
                r["asegurado_fmt"],
                r["estado_label"],
                r["dias_fmt"],
            ]
            for r in data["detalle_solicitudes_rows"]
        ],
        font_size_pt=7.5,
    )

    p_det_note = doc.add_paragraph()
    _docx_set_paragraph_spacing(p_det_note, before_twips=80, after_twips=120, line_twips=276)
    r_det_note = p_det_note.add_run(
        "El nombre de las columnas “Tomador” y “Asegurado” corresponde al remitente y a la entidad "
        "registrados en cada solicitud. Cuando el cliente actúa como integrante de un consorcio o "
        "unión temporal, se indican los integrantes registrados."
    )
    _docx_set_run_font(r_det_note, size_pt=8.0, bold=False, color=RGB_GRAY_TEXT)

    # 7. Notas metodológicas
    _docx_add_section_heading(doc, "7. Notas metodológicas")

    _docx_add_bullet_note(
        doc,
        f"Universo: solicitudes {data['universo_periodo_texto']} en las que el cliente aparece como "
        f"tomador, como asegurado o como integrante de un consorcio o unión temporal.",
    )
    _docx_add_bullet_note(
        doc,
        f"Corte: el estado de cada solicitud corresponde al {data['fecha_corte_larga']}; las "
        f"solicitudes pendientes pueden cerrarse después de esa fecha.",
    )
    _docx_add_bullet_note(
        doc,
        "Índice de efectividad de atención = (pólizas emitidas + novedades atendidas) ÷ solicitudes tramitadas.",
    )
    _docx_add_bullet_note(
        doc,
        "Índice de gestión = (pólizas emitidas + novedades atendidas + solicitudes declinadas) ÷ solicitudes tramitadas.",
    )
    pagare_cnt = data.get("pending_pagare_cartera_count", 0)
    pagare_suffix = (
        "en este período no hubo solicitudes en esa condición."
        if pagare_cnt == 0
        else f"en este período hubo {pagare_cnt} {'solicitud' if pagare_cnt == 1 else 'solicitudes'} en esa condición."
    )
    _docx_add_bullet_note(
        doc,
        f"Los pagarés pendientes y los estados de cartera se informan como motivo de pendiente cuando "
        f"la solicitud los requiere; {pagare_suffix}",
    )

    output_path_str = str(output_path)
    doc.save(output_path_str)
    logger.info("Documento de Informe de Gestión (.docx) generado exitosamente en: %s", output_path_str)
    return output_path_str


def render_informe_gestion(
    data: dict[str, Any],
    output_path: str | Path,
    assets_dir: str | Path | None = None,
    formato: str = "pptx",
) -> str:
    """Render the Management Report in either PowerPoint (.pptx) or Word (.docx) format."""
    fmt_clean = (formato or "").strip().lower()
    path_str = str(output_path).lower()
    if fmt_clean in ("docx", "documento", "word", ".docx") or path_str.endswith(".docx"):
        return render_informe_gestion_docx(data=data, output_path=output_path, assets_dir=assets_dir)
    return render_informe_gestion_pptx(data=data, output_path=output_path, assets_dir=assets_dir)
