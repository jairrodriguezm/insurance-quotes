# -*- coding: utf-8 -*-
"""Tests for the Management Reports (Informes de Gestión) module (.pptx)."""
from __future__ import annotations

from pathlib import Path
import pytest
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


def test_format_colombian_nit():
    assert format_colombian_nit("8301292898") == "830.129.289-8"
    assert format_colombian_nit("830129289") == "830.129.289-8"
    assert format_colombian_nit("830.129.289-8") == "830.129.289-8"


@pytest.mark.skipif(not XLSX_PATH.exists(), reason="Reference Excel file not present")
def test_sonacol_sept_2026_pptx_generation(tmp_path: Path):
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

    # 3. Render PPTX & verify slides, dimensions, tables, and content
    out_pptx = tmp_path / "Informe_Gestion_SONACOL_2026-09_test.pptx"
    assets_dir = Path(__file__).resolve().parent.parent / "assets"
    render_informe_gestion(data, out_pptx, assets_dir=assets_dir)
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

    # 2. Create job endpoint
    resp_job = client.post(
        "/api/v1/informes/jobs",
        data={
            "periodo": "2026-09",
            "cliente_nombre_corto": "SONACOL",
            "cliente_nombre_completo": "SONACOL – Solarte Nacional de Construcciones S.A.S.",
            "cliente_nit": "830.129.289-8",
        },
        files={"file": ("multiriesgos-CUMP-2026-10-22.xlsx", file_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert resp_job.status_code == 202
    job_id = resp_job.json()["job_id"]

    # 3. Check status
    resp_status = client.get(f"/api/v1/informes/jobs/{job_id}/status")
    assert resp_status.status_code == 200
    st_json = resp_status.json()
    assert st_json["status"] == "completed"
    assert st_json["resumen"]["kpis"]["tramitadas"] == 13

    # 4. Direct download (.pptx)
    resp_dl = client.get(f"/api/v1/informes/jobs/{job_id}/download")
    assert resp_dl.status_code == 200
    assert "presentationml.presentation" in resp_dl.headers.get("content-type", "")
    assert len(resp_dl.content) > 50000
