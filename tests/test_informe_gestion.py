# -*- coding: utf-8 -*-
"""Tests for the Management Reports (Informes de Gestión) module (.pptx and .docx)."""
from __future__ import annotations

from pathlib import Path
import pytest
from docx import Document
from pptx import Presentation
from pptx.util import Pt
from fastapi.testclient import TestClient

from app.main import app
from app.services.informe_gestion_parser import (
    format_colombian_nit,
    inspect_flokzu_xlsx,
    process_flokzu_xlsx,
)
from app.services.informe_gestion_renderer import render_informe_gestion

EXAMPLE_DIR = Path(r"c:\jair\duban\API\ejemplo informe gestion")
XLSX_PATH = EXAMPLE_DIR / "multiriesgos-CUMP-2026-10-22.xlsx"
XLSX_NOV_PATH = EXAMPLE_DIR / "multiriesgos-CUMP-2026-11-17.xlsx"
REF_DOCX_PATH = EXAMPLE_DIR / "Informe_Gestion_SONACOL_2026-09.docx"


def test_format_colombian_nit():
    assert format_colombian_nit("8301292898") == "830.129.289-8"
    assert format_colombian_nit("830129289") == "830.129.289-8"
    assert format_colombian_nit("830.129.289-8") == "830.129.289-8"


@pytest.mark.skipif(not XLSX_PATH.exists(), reason="Reference Excel file not present")
def test_sonacol_sept_2026_pptx_and_docx_generation(tmp_path: Path):
    file_bytes = XLSX_PATH.read_bytes()

    # 1. Inspect
    insp = inspect_flokzu_xlsx(file_bytes)
    assert insp.ramo == "Seguros de cumplimiento"
    assert insp.fecha_corte == "2026-10-22"
    assert insp.total_filas == 1409
    assert [p.periodo for p in insp.periodos] == ["2026-10", "2026-09", "2026-08", "2026-07"]
    sonacol_sug = next(c for c in insp.clientes_sugeridos if c.id == "SONACOL")
    assert sonacol_sug.solicitudes_por_periodo["2026-09"] == 13

    # 2. Process
    data = process_flokzu_xlsx(
        file_bytes=file_bytes,
        periodo="2026-09",
        cliente_nombre_corto="SONACOL",
        cliente_nombre_completo="SONACOL – Solarte Nacional de Construcciones S.A.S.",
        cliente_nit="830.129.289-8",
    )
    kpis = data["kpis"]
    assert kpis["tramitadas"] == 13
    assert kpis["emitidas"] == 6
    assert kpis["novedades"] == 6
    assert kpis["declinadas"] == 0
    assert kpis["pendientes"] == 1
    assert kpis["efectividad_str"] == "92,3 %"
    assert kpis["gestion_str"] == "92,3 %"
    assert kpis["tiempo_promedio_str"] == "11,2 días"
    assert data["tiempos_medidas"]["promedio"] == "11,2 días (270 horas)"
    assert data["tiempos_medidas"]["mediana"] == "11,0 días"
    assert data["tiempos_medidas"]["extremos"] == "3,6 / 21,8 días"

    assets_dir = Path(__file__).resolve().parent.parent / "assets"

    # 3A. Render PPTX & verify slides, dimensions, tables, and content
    out_pptx = tmp_path / "Informe_Gestion_SONACOL_2026-09_test.pptx"
    render_informe_gestion(data, out_pptx, assets_dir=assets_dir, formato="pptx")
    assert out_pptx.exists()

    prs = Presentation(str(out_pptx))
    assert prs.slide_width == Pt(1440.0)
    assert prs.slide_height == Pt(810.0)
    assert len(prs.slides) >= 10

    all_text = []
    tables_count = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                all_text.append(shape.text_frame.text)
            if shape.has_table:
                tables_count += 1
                tbl = shape.table
                for row in tbl.rows:
                    for cell in row.cells:
                        all_text.append(cell.text.strip())

    joined = "\n".join(all_text)
    assert "INFORME DE\nGESTIÓN MRC" in joined
    assert "SONACOL – Solarte Nacional de Construcciones S.A.S." in joined
    assert "CONTENIDO" in joined
    assert "1. RESUMEN DEL MES" in joined
    assert "2. TRAZABILIDAD DE LAS SOLICITUDES" in joined
    assert "3. SOLICITUDES DECLINADAS Y MOTIVO" in joined
    assert "4. SOLICITUDES PENDIENTES Y MOTIVO" in joined
    assert "5. TIEMPOS DE ATENCIÓN" in joined
    assert "6. DETALLE DE SOLICITUDES" in joined
    assert "7. NOTAS METODOLÓGICAS" in joined
    assert "GRACIAS" in joined
    assert "CUMP-17215" in joined
    assert "92,3 %" in joined
    assert tables_count >= 8

    # 3B. Render DOCX & verify against reference Informe_Gestion_SONACOL_2026-09.docx
    out_docx = tmp_path / "Informe_Gestion_SONACOL_2026-09_test.docx"
    render_informe_gestion(data, out_docx, assets_dir=assets_dir, formato="docx")
    assert out_docx.exists()

    doc = Document(str(out_docx))
    assert len(doc.tables) == 9
    para_texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    assert "Informe de gestión · Septiembre de 2026" in para_texts
    assert "SONACOL – Solarte Nacional de Construcciones S.A.S." in para_texts
    assert "1. Resumen del mes" in para_texts
    assert "2. Trazabilidad de las solicitudes" in para_texts
    assert "3. Solicitudes declinadas y motivo" in para_texts
    assert "4. Solicitudes pendientes y motivo" in para_texts
    assert "5. Tiempos de atención" in para_texts
    assert "6. Detalle de solicitudes" in para_texts
    assert "7. Notas metodológicas" in para_texts

    if REF_DOCX_PATH.exists():
        ref_doc = Document(str(REF_DOCX_PATH))
        ref_paras = [p.text.strip() for p in ref_doc.paragraphs if p.text.strip()]
        assert para_texts == ref_paras
        for t_idx in range(len(ref_doc.tables)):
            gen_rows = [[c.text.strip() for c in r.cells] for r in doc.tables[t_idx].rows]
            ref_rows = [[c.text.strip() for c in r.cells] for r in ref_doc.tables[t_idx].rows]
            assert gen_rows == ref_rows


@pytest.mark.skipif(not XLSX_NOV_PATH.exists(), reason="November Excel file not present")
def test_november_2026_xlsx_source_validation(tmp_path: Path):
    file_bytes = XLSX_NOV_PATH.read_bytes()
    insp = inspect_flokzu_xlsx(file_bytes)
    assert insp.ramo == "Seguros de cumplimiento"
    assert insp.fecha_corte == "2026-11-17"
    assert insp.total_filas == 635
    assert [p.periodo for p in insp.periodos] == ["2026-11", "2026-10"]

    for per in ["2026-11", "2026-10"]:
        data = process_flokzu_xlsx(
            file_bytes=file_bytes,
            periodo=per,
            cliente_nombre_corto="SONACOL",
            cliente_nombre_completo="SONACOL – Solarte Nacional de Construcciones S.A.S.",
            cliente_nit="830.129.289-8",
        )
        assert data["kpis"]["tramitadas"] == 10
        out_pptx = tmp_path / f"SONACOL_{per}.pptx"
        out_docx = tmp_path / f"SONACOL_{per}.docx"
        render_informe_gestion(data, out_pptx, formato="pptx")
        render_informe_gestion(data, out_docx, formato="docx")
        assert out_pptx.exists() and out_pptx.stat().st_size > 20000
        assert out_docx.exists() and out_docx.stat().st_size > 15000


@pytest.mark.skipif(not XLSX_PATH.exists(), reason="Reference Excel file not present")
def test_informes_api_endpoints():
    client = TestClient(app)
    file_bytes = XLSX_PATH.read_bytes()

    # 1. Inspect endpoint
    resp_insp = client.post(
        "/api/v1/informes/inspect",
        files={"file": ("multiriesgos-CUMP-2026-10-22.xlsx", file_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert resp_insp.status_code == 200
    insp_json = resp_insp.json()
    assert insp_json["fecha_corte"] == "2026-10-22"
    assert [p["periodo"] for p in insp_json["periodos"]] == ["2026-10", "2026-09", "2026-08", "2026-07"]

    # 2A. Create job endpoint for Diapositivas (.pptx)
    resp_job_pptx = client.post(
        "/api/v1/informes/jobs",
        data={
            "periodo": "2026-09",
            "cliente_nombre_corto": "SONACOL",
            "cliente_nombre_completo": "SONACOL – Solarte Nacional de Construcciones S.A.S.",
            "cliente_nit": "830.129.289-8",
            "tipo_informe": "pptx",
        },
        files={"file": ("multiriesgos-CUMP-2026-10-22.xlsx", file_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert resp_job_pptx.status_code == 202
    job_id_pptx = resp_job_pptx.json()["job_id"]

    resp_status_pptx = client.get(f"/api/v1/informes/jobs/{job_id_pptx}/status")
    assert resp_status_pptx.status_code == 200
    st_pptx = resp_status_pptx.json()
    assert st_pptx["status"] == "completed"
    assert st_pptx["resumen"]["formato"] == "pptx"
    assert st_pptx["resumen"]["kpis"]["tramitadas"] == 13

    resp_dl_pptx = client.get(f"/api/v1/informes/jobs/{job_id_pptx}/download")
    assert resp_dl_pptx.status_code == 200
    assert "presentationml.presentation" in resp_dl_pptx.headers.get("content-type", "")
    assert len(resp_dl_pptx.content) > 50000

    # 2B. Create job endpoint for Documento (.docx)
    resp_job_docx = client.post(
        "/api/v1/informes/jobs",
        data={
            "periodo": "2026-09",
            "cliente_nombre_corto": "SONACOL",
            "cliente_nombre_completo": "SONACOL – Solarte Nacional de Construcciones S.A.S.",
            "cliente_nit": "830.129.289-8",
            "tipo_informe": "docx",
        },
        files={"file": ("multiriesgos-CUMP-2026-10-22.xlsx", file_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert resp_job_docx.status_code == 202
    job_id_docx = resp_job_docx.json()["job_id"]

    resp_status_docx = client.get(f"/api/v1/informes/jobs/{job_id_docx}/status")
    assert resp_status_docx.status_code == 200
    st_docx = resp_status_docx.json()
    assert st_docx["status"] == "completed"
    assert st_docx["resumen"]["formato"] == "docx"
    assert st_docx["resumen"]["kpis"]["tramitadas"] == 13

    resp_dl_docx = client.get(f"/api/v1/informes/jobs/{job_id_docx}/download")
    assert resp_dl_docx.status_code == 200
    assert "wordprocessingml.document" in resp_dl_docx.headers.get("content-type", "")
    assert len(resp_dl_docx.content) > 15000
