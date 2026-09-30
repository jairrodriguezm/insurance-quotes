# -*- coding: utf-8 -*-
"""Tests for LLM Extractor service."""
from __future__ import annotations

import json
import unittest
from unittest.mock import AsyncMock, MagicMock, patch

from app.schemas.quote import ExtractedQuote, ProjectMeta
from app.schemas.comparative import Recommendation
from app.services.llm_extractor import (
    MASTER_COMPARATIVE_PROMPT,
    EXTRACTION_SYSTEM_PROMPT,
    META_EXTRACTION_PROMPT,
    RECOMMENDATION_PROMPT,
    WORST_MARKERS_PROMPT,
    extract_quote,
    extract_project_meta,
    generate_recommendation,
    generate_worst_markers,
    generate_comparative_report,
    generate_comparative_report_from_data,
)


class TestLLMExtractor(unittest.IsolatedAsyncioTestCase):
    """Unit tests for llm_extractor module."""

    def test_prompts_content(self):
        """Verify that master and specialized prompts contain required sections and underwriting rules."""
        self.assertIn("Analista Técnico Senior de Suscripción de Multiriesgos de Colombia Ltda.", MASTER_COMPARATIVE_PROMPT)
        self.assertIn("I. INFORMACIÓN GENERAL", MASTER_COMPARATIVE_PROMPT)
        self.assertIn("II. COTIZACIONES REALIZADAS", MASTER_COMPARATIVE_PROMPT)
        self.assertIn("III. RECOMENDACIÓN", MASTER_COMPARATIVE_PROMPT)
        self.assertIn("IV. COBERTURAS", MASTER_COMPARATIVE_PROMPT)
        self.assertIn("V. DEDUCIBLES", MASTER_COMPARATIVE_PROMPT)
        self.assertIn("VI. PREVALENCIA DE LOS TÉRMINOS COTIZADOS", MASTER_COMPARATIVE_PROMPT)
        self.assertIn("ANEXO 1: SUBJETIVIDADES Y CONDICIONES PARTICULARES", MASTER_COMPARATIVE_PROMPT)
        self.assertIn("ANEXO 2: OBSERVACIONES Y PREGUNTAS DEL CLIENTE", MASTER_COMPARATIVE_PROMPT)
        self.assertIn("CERO ALUCINACIONES", EXTRACTION_SYSTEM_PROMPT)
        self.assertIn("NO ESPECIFICA", EXTRACTION_SYSTEM_PROMPT)
        self.assertIn("No cotizado", EXTRACTION_SYSTEM_PROMPT)
        self.assertIn("N/A", EXTRACTION_SYSTEM_PROMPT)

    @patch("app.services.llm_extractor._get_client")
    async def test_extract_quote_success(self, mock_get_client):
        mock_client = MagicMock()
        mock_response = MagicMock()
        mock_response.text = json.dumps({
            "nombre_compania": "Zurich Colombia Seguros S.A.",
            "id_compania": "zurich",
            "opciones": [
                {
                    "etiqueta": "Opción 1",
                    "tasa": "0,949 ‰",
                    "prima": 40508863.0,
                    "modalidad": "Claims Made. Prima Neta: $ 40.508.863,00, IVA: $ 7.696.683,97, Total: $ 48.205.546,97."
                }
            ],
            "coberturas": {"danos_materiales": 2000000000},
            "deducibles": {"ded_incendio": "10% mínimo 5 SMMLV"},
            "subjetividades": ["Ambulancia medicalizada en sitio"],
            "tasa_prorroga": "Prima única NO ESPECIFICA",
            "validez_oferta": "15 días calendario"
        })
        mock_client.aio.models.generate_content = AsyncMock(return_value=mock_response)
        mock_get_client.return_value = mock_client

        quote = await extract_quote("Contenido de prueba de cotización")
        self.assertIsInstance(quote, ExtractedQuote)
        self.assertEqual(quote.nombre_compania, "Zurich Colombia Seguros S.A.")
        self.assertEqual(quote.id_compania, "zurich")
        self.assertEqual(len(quote.opciones), 1)
        self.assertEqual(quote.opciones[0].prima, 40508863.0)

    @patch("app.services.llm_extractor._get_client")
    async def test_extract_project_meta_success(self, mock_get_client):
        mock_client = MagicMock()
        mock_response = MagicMock()
        mock_response.text = json.dumps({
            "tipo_cobertura": "RESPONSABILIDAD CIVIL EXTRACONTRACTUAL EVENTOS",
            "asegurado": "Eventos Especiales S.A.S. - NIT 900.888.777-1",
            "beneficiario": "Asistentes al evento",
            "vigencia_construccion": {"desde": "15/10/2026", "hasta": "17/10/2026"},
            "ubicacion": "Corferias, Bogotá D.C.",
            "descripcion_proyecto": "Feria comercial con aforo de 30 personas.",
            "valor_asegurado": None
        })
        mock_client.aio.models.generate_content = AsyncMock(return_value=mock_response)
        mock_get_client.return_value = mock_client

        meta = await extract_project_meta(["texto doc 1", "texto doc 2"])
        self.assertIsInstance(meta, ProjectMeta)
        self.assertEqual(meta.tipo_cobertura, "RESPONSABILIDAD CIVIL EXTRACONTRACTUAL EVENTOS")
        self.assertEqual(meta.ubicacion, "Corferias, Bogotá D.C.")

    @patch("app.services.llm_extractor._get_client")
    async def test_generate_recommendation_success(self, mock_get_client):
        mock_client = MagicMock()
        mock_response = MagicMock()
        mock_response.text = json.dumps({
            "aseguradora_id": "zurich",
            "opcion": "Opción 1",
            "vinetas": [
                "Comparativo Económico: Ahorro del 15% frente a HDI.",
                "Capacidad: Aforo cubierto según requerimiento de 30 personas."
            ]
        })
        mock_client.aio.models.generate_content = AsyncMock(return_value=mock_response)
        mock_get_client.return_value = mock_client

        rec = await generate_recommendation({"aseguradoras": []})
        self.assertIsInstance(rec, Recommendation)
        self.assertEqual(rec.aseguradora_id, "zurich")
        self.assertEqual(len(rec.vinetas), 2)

    @patch("app.services.llm_extractor._get_client")
    async def test_generate_worst_markers_success(self, mock_get_client):
        mock_client = MagicMock()
        mock_response = MagicMock()
        mock_response.text = json.dumps({"ded_rce": ["chubb"]})
        mock_client.aio.models.generate_content = AsyncMock(return_value=mock_response)
        mock_get_client.return_value = mock_client

        markers = await generate_worst_markers({"aseguradoras": []})
        self.assertEqual(markers, {"ded_rce": ["chubb"]})

    @patch("app.services.llm_extractor._get_client")
    async def test_generate_comparative_report_success(self, mock_get_client):
        mock_client = MagicMock()
        mock_response = MagicMock()
        mock_response.text = "# mrc - multiriesgos de colombia\n\n### I. INFORMACIÓN GENERAL\n..."
        mock_client.aio.models.generate_content = AsyncMock(return_value=mock_response)
        mock_get_client.return_value = mock_client

        report = await generate_comparative_report(["texto 1", "texto 2"])
        self.assertTrue(report.startswith("# mrc - multiriesgos de colombia"))

    @patch("app.services.llm_extractor._get_client")
    async def test_generate_comparative_report_from_data_success(self, mock_get_client):
        mock_client = MagicMock()
        mock_response = MagicMock()
        mock_response.text = "# mrc - multiriesgos de colombia\n\n### I. INFORMACIÓN GENERAL\n..."
        mock_client.aio.models.generate_content = AsyncMock(return_value=mock_response)
        mock_get_client.return_value = mock_client

        report = await generate_comparative_report_from_data({"meta": {}, "aseguradoras": []})
        self.assertTrue(report.startswith("# mrc - multiriesgos de colombia"))

    def test_category_prompts_loading(self):
        """Verify that category-specific prompt files exist and load correctly."""
        from app.services.llm_extractor import load_prompt_for_category, get_prompt_path_for_category
        from app.services.concept_mapper import normalize_str

        for cat in ["Autos", "Copropiedades", "Hogar", "Todo_Riesgo_Construccion", "Pyme", "Maquinaria_Equipo"]:
            path = get_prompt_path_for_category(cat)
            self.assertTrue(path.exists(), f"Prompt file for {cat} must exist at {path}")
            content = load_prompt_for_category(cat)
            self.assertIn("multiriesgos de colombia", normalize_str(content))
            self.assertIn(normalize_str(cat), normalize_str(content))

    @patch("app.services.llm_extractor._get_client")
    async def test_extract_quote_with_category_and_custom_prompt(self, mock_get_client):
        """Test extract_quote using a specific category and custom prompt_file."""
        import tempfile
        from pathlib import Path

        mock_client = MagicMock()
        mock_response = MagicMock()
        mock_response.text = json.dumps({
            "nombre_compania": "Seguros SURA",
            "id_compania": "sura",
            "opciones": [
                {
                    "etiqueta": "Plan Autos Global",
                    "tasa": "Prima fija",
                    "prima": 2500000.0,
                    "modalidad": "Todo Riesgo Autos"
                }
            ],
            "coberturas": {"rce": 2000000000},
            "deducibles": {"ded_rce": "10% min 1 SMMLV"},
            "subjetividades": ["Inspección previa"],
            "tasa_prorroga": "NO ESPECIFICA",
            "validez_oferta": "30 días"
        })
        mock_client.aio.models.generate_content = AsyncMock(return_value=mock_response)
        mock_get_client.return_value = mock_client

        # 1. Test with categoria="Autos"
        quote_autos = await extract_quote("Texto cotización autos", categoria="Autos")
        self.assertEqual(quote_autos.nombre_compania, "Seguros SURA")

        # 2. Test with custom prompt_file
        with tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False, encoding="utf-8") as tmp:
            tmp.write("Prompt personalizado de prueba para suscripción.")
            tmp_path = Path(tmp.name)

        try:
            quote_custom = await extract_quote("Texto", categoria="Autos", prompt_file=tmp_path)
            self.assertEqual(quote_custom.id_compania, "sura")
        finally:
            if tmp_path.exists():
                tmp_path.unlink()


if __name__ == "__main__":
    unittest.main()
