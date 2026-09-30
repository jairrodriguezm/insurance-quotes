# -*- coding: utf-8 -*-
"""Unit tests for ConceptMapper service."""
from __future__ import annotations

import unittest
from pathlib import Path

from app.services.concept_mapper import (
    ConceptMapper,
    get_concept_mapper,
    map_concept_to_sfc,
    normalize_str,
)


class TestConceptMapper(unittest.TestCase):
    """Test suite for concept_matrix dictionary lookup and normalization."""

    @classmethod
    def setUpClass(cls):
        cls.mapper = get_concept_mapper()

    def test_matrix_loaded_successfully(self):
        """Verify that concept_matrix.xlsx is loaded with all expected columns and rows."""
        self.assertGreater(len(self.mapper.entries), 100)
        self.assertIn("sura", self.mapper.insurers)
        self.assertIn("zurich", self.mapper.insurers)
        self.assertIn("chubb", self.mapper.insurers)

    def test_category_normalization(self):
        """Test normalization of user-supplied category names."""
        self.assertEqual(ConceptMapper.normalize_category("autos"), "Autos")
        self.assertEqual(ConceptMapper.normalize_category("vehiculos"), "Autos")
        self.assertEqual(ConceptMapper.normalize_category("copropiedades"), "Copropiedades")
        self.assertEqual(ConceptMapper.normalize_category("propiedad horizontal"), "Copropiedades")
        self.assertEqual(ConceptMapper.normalize_category("hogar"), "Hogar")
        self.assertEqual(ConceptMapper.normalize_category("TRC"), "Todo_Riesgo_Construccion")
        self.assertEqual(ConceptMapper.normalize_category("todo riesgo construccion"), "Todo_Riesgo_Construccion")
        self.assertEqual(ConceptMapper.normalize_category("pyme"), "Pyme")
        self.assertEqual(ConceptMapper.normalize_category("maquinaria"), "Maquinaria_Equipo")

    def test_autos_concept_mapping(self):
        """Test concept mapping for Autos insurer synonyms."""
        # Synonym for Conductor Elegido in Zurich: 'Chofer de Reemplazo'
        self.assertEqual(
            map_concept_to_sfc("Chofer de Reemplazo", "Autos"),
            "Conductor Elegido",
        )
        # Synonym for Asistencia Jurídica in Seguros del Estado: 'Defensa Judicial'
        self.assertEqual(
            map_concept_to_sfc("Defensa Judicial", "Autos"),
            "Asistencia Jurídica",
        )
        # Synonym for Gastos de Transporte in Berkley: 'Gastos de Movilización'
        self.assertEqual(
            map_concept_to_sfc("Gastos de Movilización", "Autos"),
            "Gastos de Transporte Por Pérdidas Totales",
        )
        # Synonym for RCE Daños a bienes de terceros: 'Daños a Terceros'
        self.assertEqual(
            map_concept_to_sfc("Daños a Terceros", "Autos"),
            "RCE Daños a bienes de Terceros",
        )

    def test_hogar_concept_mapping(self):
        """Test concept mapping for Hogar insurer synonyms."""
        # 'Contenidos Hogar' -> 'Muebles y enseres domésticos del hogar'
        self.assertEqual(
            map_concept_to_sfc("Contenidos Hogar", "Hogar"),
            "Muebles y enseres domésticos del hogar",
        )
        # 'Objetos de Valor y Arte' -> 'Obras de arte'
        self.assertEqual(
            map_concept_to_sfc("Objetos de Valor y Arte", "Hogar"),
            "Obras de arte",
        )

    def test_copropiedades_concept_mapping(self):
        """Test concept mapping for Copropiedades insurer synonyms."""
        # 'Rotura de Maquinaria' -> 'Rotura de maquinaria (daño interno)'
        self.assertEqual(
            map_concept_to_sfc("Rotura de Maquinaria", "Copropiedades"),
            "Rotura de maquinaria (daño interno)",
        )

    def test_trc_concept_mapping(self):
        """Test concept mapping for Todo Riesgo Construcción insurer synonyms."""
        # 'Pilotaje y Tablestacado' -> 'Cimentación por pilotaje y tablestacados para fosas de obras'
        self.assertEqual(
            map_concept_to_sfc("Pilotaje y Tablestacado", "Todo_Riesgo_Construccion"),
            "Cimentación por pilotaje y tablestacados para fosas de obras",
        )

    def test_global_todas_category_fallback(self):
        """Verify that concepts in category 'Todas' (like Deducibles) map across any category."""
        self.assertEqual(
            map_concept_to_sfc("Deducible / Coaseguro", "Autos"),
            "Participación del Asegurado en Siniestros",
        )

    def test_normalize_extracted_dict(self):
        """Test mapping an entire dictionary of extracted keys."""
        raw = {
            "Chofer de Reemplazo": "6 eventos",
            "Defensa Judicial": "Ilimitada",
            "Daños a Terceros": 1000000000,
        }
        normalized = self.mapper.normalize_extracted_dict(raw, "Autos")
        self.assertIn("Conductor Elegido", normalized)
        self.assertIn("Asistencia Jurídica", normalized)
        self.assertIn("RCE Daños a bienes de Terceros", normalized)
        self.assertEqual(normalized["Conductor Elegido"], "6 eventos")


if __name__ == "__main__":
    unittest.main()
