# -*- coding: utf-8 -*-
"""Unit tests for DOCX Renderer service across category layouts."""
from __future__ import annotations

import json
import os
import tempfile
import unittest
from pathlib import Path

from docx import Document

from app.schemas.comparative import Recommendation
from app.schemas.quote import ExtractedQuote, ProjectMeta, QuoteOption
from app.services.consolidator import consolidate_quotes
from app.services.docx_renderer import render_comparative


class TestDocxRendererLayouts(unittest.TestCase):
    """Test suite for Word document rendering using assets/layouts templates."""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.sample_data = {
            "meta": {
                "tomador": "Test Corp S.A.S.",
                "identificacion": "900.555.444-3",
                "asegurado": "Test Corp S.A.S.",
                "beneficiario": "Terceros Afectados",
                "marca": "CHEVROLET",
                "placa": "XYZ123",
                "linea": "TRACKER",
                "modelo": "2024",
                "fecha": "28/09/2026",
                "valor_asegurado": 85000000,
                "ubicacion": "Bogotá D.C.",
                "descripcion_proyecto": "Riesgo estándar",
            },
            "aseguradoras": [
                {
                    "id": "sura",
                    "nombre": "Seguros SURA",
                    "logo": "sura.png",
                    "opciones": [
                        {
                            "etiqueta": "Opción 1",
                            "tasa": "1,2 ‰",
                            "prima": 2500000.0,
                            "modalidad": "Todo Riesgo",
                        }
                    ],
                    "coberturas": {
                        "RCE Daños a bienes de Terceros": 1000000000,
                        "Perdida parcial y total Daños": "100%",
                        "Incendio y/o rayo": "100%",
                    },
                    "deducibles": {
                        "RCE": "10% min 1 SMMLV",
                        "Incendio y/o rayo": "10% min 2 SMMLV",
                    },
                    "subjetividades": ["Inspección técnica"],
                },
                {
                    "id": "zurich",
                    "nombre": "Zurich Colombia",
                    "logo": "zurich.png",
                    "opciones": [
                        {
                            "etiqueta": "Opción 1",
                            "tasa": "1,1 ‰",
                            "prima": 2350000.0,
                            "modalidad": "Todo Riesgo",
                        }
                    ],
                    "coberturas": {
                        "RCE Daños a bienes de Terceros": 1200000000,
                        "Perdida parcial y total Daños": "100%",
                        "Incendio y/o rayo": "100%",
                    },
                    "deducibles": {
                        "RCE": "10% min 1 SMMLV",
                        "Incendio y/o rayo": "10% min 2 SMMLV",
                    },
                    "subjetividades": ["SARLAFT"],
                },
            ],
            "recomendacion": {
                "aseguradora_id": "zurich",
                "opcion": "Opción 1",
                "vinetas": ["Mayor cobertura económica", "Tarifa más competitiva"],
            },
        }

    def test_render_autos_layout(self):
        """Test rendering Autos layout."""
        out = os.path.join(self.temp_dir, "comparativo_autos.docx")
        path = render_comparative(self.sample_data, output_path=out, categoria="Autos")
        self.assertTrue(os.path.exists(path))
        doc = Document(path)
        self.assertGreater(len(doc.tables), 5)
        # Verify tomador filled in General Info table
        t1_text = " ".join(c.text for row in doc.tables[1].rows for c in row.cells)
        self.assertIn("Test Corp S.A.S.", t1_text)

    def test_render_copropiedades_layout(self):
        """Test rendering Copropiedades layout."""
        out = os.path.join(self.temp_dir, "comparativo_copropiedades.docx")
        path = render_comparative(self.sample_data, output_path=out, categoria="Copropiedades")
        self.assertTrue(os.path.exists(path))
        doc = Document(path)
        self.assertEqual(len(doc.tables), 8)

    def test_render_copropiedades_full_almeria_structure(self):
        """Verify end-to-end Copropiedades consolidation and rendering against official standard."""
        q_estado = ExtractedQuote(
            nombre_compania="Seguros del Estado S.A.",
            id_compania="estado",
            opciones=[QuoteOption(etiqueta="Opción 1", prima=11015324.0, modalidad="Deducible terremoto sobre\nValor Asegurable")],
            coberturas={"incendio_rayo_explosion_agua": 5386845442, "manejo": 10000000, "rce_basico": 400000000, "rce_vehiculos": "N/A"},
            deducibles={
                "ded_incendio": "A) Riesgo de incendio y/o rayo: SIN DEDUCIBLE. B) Demás eventos: 5% de la pérdida, mínimo 1 SMMLV",
                "ded_terremoto": "2% del valor asegurable, mínimo 2 SMMLV",
            },
            asistencias={
                "cerrajero_comunes": "30 SMDLV\n8 eventos vigencia",
                "cerrajero_privadas": "No presta servicio a las áreas privadas",
            },
        )
        # Two separate PDFs for Zurich (Option 1 VA & Option 2 Valor de la Pérdida)
        q_zurich_1 = ExtractedQuote(
            nombre_compania="Zurich Colombia Seguros S.A.",
            id_compania="zurich",
            opciones=[QuoteOption(etiqueta="Opción 1", prima=14129835.0, modalidad="Deducible terremoto sobre\nValor Asegurable")],
            coberturas={"incendio_rayo_explosion_agua": 5386845442, "manejo": 30000000, "rce_basico": 300000000, "rce_vehiculos": "40% evento/vigencia"},
            deducibles={"ded_terremoto": "2% del valor asegurable del artículo afectado, mínimo 1 SMMLV"},
            asistencias={
                "cerrajero_comunes": "30 SMDLV\nMax.15 eventos vigencia",
                "reparacion_propiedades_vecinas": "30 SMDLV\n2 eventos vigencia",
                "cerrajero_privadas": "15 eventos vigencia entre Comunes y Privadas",
            },
        )
        q_zurich_2 = ExtractedQuote(
            nombre_compania="Zurich Colombia Seguros S.A.",
            id_compania="zurich",
            opciones=[QuoteOption(etiqueta="Opción 1", prima=15276625.0, modalidad="Deducible terremoto sobre\nValor de la Pérdida")],
            coberturas={"incendio_rayo_explosion_agua": 5386845442, "manejo": 30000000, "rce_basico": 300000000, "rce_vehiculos": "40% evento/vigencia"},
            deducibles={"ded_terremoto": "3% de la pérdida, mínimo 1 SMMLV"},
            asistencias={
                "cerrajero_comunes": "30 SMDLV\nMax.15 eventos vigencia",
                "reparacion_propiedades_vecinas": "30 SMDLV\n2 eventos vigencia",
                "cerrajero_privadas": "15 eventos vigencia entre Comunes y Privadas",
            },
        )
        q_mapfre = ExtractedQuote(
            nombre_compania="Mapfre Seguros Generales de Colombia S.A.",
            id_compania="mapfre",
            opciones=[QuoteOption(etiqueta="Opción 1", prima=18926907.0, modalidad="Deducible terremoto sobre\nValor Asegurable")],
            coberturas={"incendio_rayo_explosion_agua": 5386845442, "manejo": 10000000, "rce_basico": 200000000, "rce_vehiculos": "30% evento/vigencia"},
            deducibles={"ded_incendio": "10% PERD Min 1 (SMMLV)", "ded_terremoto": "2% ABLE Min 2 (SMMLV)"},
            asistencias={"cerrajero_comunes": "20 eventos vigencia entre Comunes y Privadas"},
        )

        meta = ProjectMeta(
            tipo_cobertura="MULTIRRIESGO COPROPIEDADES",
            tomador="EDIFICIO ALMERIA P.H.",
            asegurado="EDIFICIO ALMERIA P.H. - NIT 8300705739",
            beneficiario="EDIFICIO ALMERIA P.H. / Copropietarios / Terceros Afectados",
            direccion="CL 81 # 8 - 51",
            ciudad="Bogotá",
            ano_construccion="1990",
            valor_edificio=5149392347,
            valor_cimentacion="(incluida en ítem anterior)",
            valor_maquinaria=206598117,
            valor_muebles=9845000,
            valor_equipos=21009978,
            valor_equipos_moviles=21009978,
            valor_rce=400000000,
            valor_dno=200000000,
            valor_asegurado=5386845442,
        )
        rec = Recommendation(
            aseguradora_id="zurich",
            opcion="Opción 2",
            vinetas=[
                "Mejor deducible en Terremoto: 3% del valor de la pérdida, mínimo 1 SMMLV",
                "Mejores deducibles en HMACC, AMIT, Rotura de maquinaria, equipo electrónico.",
            ],
        )

        consolidated = consolidate_quotes(
            quotes=[q_mapfre, q_zurich_2, q_estado, q_zurich_1],
            project_meta=meta,
            recommendation=rec,
            categoria="Copropiedades",
        )
        # 1. Zurich PDFs merged into 1 entry with 2 sorted options; insurers sorted by ascending min premium
        self.assertEqual([a.id for a in consolidated.aseguradoras], ["estado", "zurich", "mapfre"])
        self.assertEqual(len(consolidated.aseguradoras[1].opciones), 2)
        self.assertEqual(consolidated.aseguradoras[1].opciones[0]["etiqueta"], "Opción 1")
        self.assertEqual(consolidated.aseguradoras[1].opciones[1]["etiqueta"], "Opción 2")
        self.assertIn("Opción 1:", consolidated.aseguradoras[1].deducibles["ded_terremoto"])
        self.assertIn("Opción 2:", consolidated.aseguradoras[1].deducibles["ded_terremoto"])

        out = os.path.join(self.temp_dir, "comparativo_almeria_test.docx")
        data_dict = json.loads(consolidated.model_dump_json())
        render_comparative(data_dict, output_path=out, categoria="Copropiedades")

        doc = Document(out)
        # Exact 8 tables (no TRC Annexes appended)
        self.assertEqual(len(doc.tables), 8)
        all_paras = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("II. VALORES ASEGURADOS", all_paras)
        self.assertNotIn("ANEXO 1", all_paras)

        # Table 1: EDIFICIO shows 5.149.392.347 (not total 5.386.845.442), MANEJO shows mode '$ 10.000.000'
        self.assertEqual(doc.tables[1].rows[1].cells[1].text.strip(), "$ 5.149.392.347")
        self.assertEqual(doc.tables[1].rows[9].cells[1].text.strip(), "$ 10.000.000")

        # Table 2: 1 header + 4 option rows (Estado, Zurich Op1, Zurich Op2, Mapfre)
        self.assertEqual(len(doc.tables[2].rows), 5)
        self.assertIn("Valor Asegurable", doc.tables[2].rows[2].cells[2].text)
        self.assertIn("Valor de la Pérdida", doc.tables[2].rows[3].cells[2].text)

        # Matrix exact row counts (29 coverages, 17 deductibles, 25 asistencias)
        self.assertEqual(len(doc.tables[4].rows), 29)
        self.assertEqual(len(doc.tables[5].rows), 17)
        self.assertEqual(len(doc.tables[6].rows), 25)

        # Table 4 basic fire coverage uses total sum insured ($ 5.386.845.442)
        self.assertEqual(doc.tables[4].rows[2].cells[1].text.strip(), "$ 5.386.845.442")

        # Table 5 Explosión/Agua deductible strips Clause A fire text, Mapfre PERD/ABLE expanded
        self.assertEqual(doc.tables[5].rows[3].cells[1].text.strip(), "5% de la pérdida, mínimo 1 SMMLV")
        self.assertEqual(doc.tables[5].rows[2].cells[3].text.strip(), "10% de la pérdida, mínimo 1 SMMLV")
        self.assertEqual(doc.tables[5].rows[5].cells[3].text.strip(), "2% del valor asegurable, mínimo 2 SMMLV")

        # Table 6 Asistencias: no cross-section leakage between cerrajero_comunes and cerrajero_privadas
        self.assertIn("30 SMDLV", doc.tables[6].rows[2].cells[1].text)
        self.assertEqual(doc.tables[6].rows[18].cells[1].text.strip(), "No presta servicio a las áreas privadas")

    def test_project_meta_enrichment_and_fallback(self):
        """Verify ProjectMeta coercion, enrichment from quotes, and deterministic fallback."""
        from app.services.llm_extractor import _build_fallback_project_meta, _enrich_meta_from_quotes

        q_axa = ExtractedQuote(
            nombre_compania="AXA Colpatria Seguros S.A.",
            id_compania="axa_colpatria",
            opciones=[QuoteOption(etiqueta="Opción 1", tasa="", prima=14314667.0, modalidad="Deducible terremoto sobre\nValor Asegurable")],
            coberturas={
                "incendio_rayo_explosion_agua": 5386845442,
                "rotura_maquinaria": 206598117,
                "muebles_dano_interno": 9845000,
                "equipo_electronico_dano_interno": 21009978,
                "rce_basico": 400000000,
                "dno_directores": 200000000,
                "manejo": 10000000,
            },
        )
        raw_texts = [
            "TOMADOR: EDIFICIO ALMERIA P.H.\nNIT: 830.070.573-9\nDirección del Riesgo: CL 81 # 8 - 51\nCiudad: BOGOTA\nAño de Construcción: 1.990"
        ]
        fb_meta = _build_fallback_project_meta(raw_texts, "Copropiedades", [q_axa])
        self.assertEqual(fb_meta.tomador, "EDIFICIO ALMERIA P.H.")
        self.assertEqual(fb_meta.ano_construccion, "1990")
        self.assertEqual(fb_meta.valor_asegurado, 5386845442.0)
        self.assertEqual(fb_meta.valor_edificio, 5149392347.0)
        self.assertEqual(fb_meta.valor_manejo, 10000000.0)

        # Also test coercion when LLM passes nulls, ints for strings, or formatted COP strings
        partial = ProjectMeta.model_validate(
            {
                "tipo_cobertura": None,
                "tomador": "EDIFICIO ALMERIA P.H.",
                "ano_construccion": 1990,
                "valor_asegurado": "$ 5.386.845.442",
                "valor_edificio": 5386845442,
                "numero_torres": None,
            }
        )
        enriched = _enrich_meta_from_quotes(partial, [q_axa], raw_texts, "Copropiedades")
        self.assertEqual(enriched.valor_edificio, 5149392347.0)
        self.assertEqual(enriched.ano_construccion, "1990")
        self.assertEqual(enriched.numero_torres, "")

    def test_render_hogar_layout(self):
        """Test rendering Hogar layout."""
        out = os.path.join(self.temp_dir, "comparativo_hogar.docx")
        path = render_comparative(self.sample_data, output_path=out, categoria="Hogar")
        self.assertTrue(os.path.exists(path))
        doc = Document(path)
        self.assertGreater(len(doc.tables), 4)

    def test_render_trc_layout(self):
        """Test rendering Todo Riesgo Construcción layout."""
        out = os.path.join(self.temp_dir, "comparativo_trc.docx")
        path = render_comparative(self.sample_data, output_path=out, categoria="Todo_Riesgo_Construccion")
        self.assertTrue(os.path.exists(path))
        doc = Document(path)
        self.assertGreater(len(doc.tables), 4)


if __name__ == "__main__":
    unittest.main()
