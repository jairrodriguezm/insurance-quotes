# -*- coding: utf-8 -*-
"""LLM Extractor service — calls Gemini with structured outputs and domain-specific underwriting prompts.

Implements the technical underwriting standard of Multiriesgos de Colombia Ltda.:
- Map pattern: each quote document is processed individually (extract_quote) producing an ExtractedQuote.
- Metadata extraction: project and risk metadata inferred across all quotes (extract_project_meta).
- Recommendation: objective technical recommendation based on 5 core underwriting criteria (generate_recommendation).
- Worst markers: identification of clearly inferior terms per coverage/deductible (generate_worst_markers).
- Official comparative report: direct generation of the official 6-section + 2-annex Markdown slip (generate_comparative_report).

Determinism: Temperature is set to 0.0 to maximize fidelity and prevent hallucinations.
"""
from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Optional

from google import genai
from google.genai import types

from app.core.config import settings
from app.core.exceptions import LLMExtractionError
from app.schemas.quote import ExtractedQuote, ProjectMeta
from app.schemas.comparative import Recommendation
from app.services.concept_mapper import ConceptMapper

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Category Prompt Files Registry
# ---------------------------------------------------------------------------
PROMPTS_DIR = Path(__file__).resolve().parent.parent / "prompts"

CATEGORY_PROMPT_FILES: dict[str, Path] = {
    "Autos": PROMPTS_DIR / "autos.txt",
    "Copropiedades": PROMPTS_DIR / "copropiedades.txt",
    "Hogar": PROMPTS_DIR / "hogar.txt",
    "Todo_Riesgo_Construccion": PROMPTS_DIR / "todo_riesgo_construccion.txt",
    "Pyme": PROMPTS_DIR / "pyme.txt",
    "Maquinaria_Equipo": PROMPTS_DIR / "maquinaria_equipo.txt",
    "General": PROMPTS_DIR / "general.txt",
}


def get_prompt_path_for_category(categoria: str) -> Path:
    """Resolve prompt file path for a given category."""
    norm = ConceptMapper.normalize_category(categoria)
    return CATEGORY_PROMPT_FILES.get(
        norm,
        CATEGORY_PROMPT_FILES.get("General", PROMPTS_DIR / "todo_riesgo_construccion.txt"),
    )


def load_prompt_for_category(
    categoria: str = "Todo_Riesgo_Construccion",
    prompt_file: str | Path | None = None,
) -> str:
    """Load system prompt text for a specific insurance category or custom file."""
    if prompt_file:
        p = Path(prompt_file)
        if p.exists():
            return p.read_text(encoding="utf-8")
        logger.warning("Archivo de prompt personalizado no encontrado: %s", prompt_file)

    target_path = get_prompt_path_for_category(categoria)
    if target_path.exists():
        return target_path.read_text(encoding="utf-8")

    return EXTRACTION_SYSTEM_PROMPT

# ---------------------------------------------------------------------------
# Master System Prompt: Standard Oficial Multiriesgos de Colombia Ltda.
# ---------------------------------------------------------------------------

MASTER_COMPARATIVE_PROMPT = """# PROMPT DEL SISTEMA: COMPARATIVO TÉCNICO DE COTIZACIONES DE SEGUROS (MULTIRIESGOS DE COLOMBIA)

Eres el **Analista Técnico Senior de Suscripción de Multiriesgos de Colombia Ltda.**, corredor de seguros profesional. Tu objetivo es procesar las cotizaciones de seguros adjuntas (slips, cartas o propuestas técnicas de compañías aseguradoras como Zurich, AXA Colpatria, HDI, Chubb, Seguros Mundial, SURA, Allianz, Bolívar, etc.) y generar un **Informe Comparativo Oficial (Slip Comparativo de Coaseguro/Mercado)** con el más alto rigor técnico, fidelidad documental y criterio asegurador colombiano.

El informe debe adaptarse con precisión al ramo cotizado (Copropiedades / Multirriesgo, RCE Eventos / Operacional, Todo Riesgo Construcción / TRC, etc.), generando exactamente las **6 Secciones Principales** y **2 Anexos Técnicos** del estándar oficial de la firma.

---

## REGLAS DE EXTRACCIÓN Y SUSCRIPCIÓN (CERO ALUCINACIONES)

1. **Fidelidad Absoluta y Transcripción Literal**:
   - Transcribe fielmente valores numéricos, porcentajes y condiciones textuales. Nunca redondees, deduzcas ni inventes coberturas.
   - Formato monetario: Expresa cifras en pesos colombianos formateadas con separador de miles (ej: `$ 40.508.863,00`, `$ 966.280,00` o `$ 200.000.000`).
   - Tasas: Exprésalas en por mil (`‰`), porcentaje, valor por millón o modalidad de tarifa (ej: `0,949 ‰`, `Prima fija por evento` o `Tabla de tarifación por millón/día`). Si la propuesta no especifica la tasa, coloca estrictamente: `Prima única NO ESPECIFICA`.
   - Desglose Financiero de Primas: Discrimina con claridad en las notas y modalidad la **Prima Neta**, **Gastos de Emisión/Expedición**, **IVA (19%)** y el **Total a pagar**. En la columna "PRIMA" de la Sección II consigna siempre la **Prima Neta antes de IVA** (o el valor liquidado neto antes de tributos).

2. **Convenciones Estrictas de Silencio o Exclusión**:
   - `NO ESPECIFICA`: Úsalo EXCLUSIVAMENTE cuando el documento fuente guarde silencio absoluto sobre el concepto, amparo o deducible.
   - `N/A`: Úsalo cuando la aseguradora exprese que la cobertura no aplica, no contrata o esté formalmente excluida.
   - `No cotizado`: Úsalo cuando la cotización incluya el campo en su esquema pero asigne $0, período de indemnización 0, o consigne la leyenda explícita "No cotizado".

3. **Manejo de Opciones**:
   - Si una aseguradora ofrece múltiples opciones (ej: Zurich Opción 1 vs Opción 2), desglosa cada una en filas independientes en la Sección II e identifica sus diferencias en deducibles y costos.

4. **Auditoría Técnica y Detección de Discrepancias (Obligatoria)**:
   - **Vigencias y Desfases Temporales**: Compara fechas de emisión y periodos amparados. Si una cotización tiene vigencia futura (ej: 2026-2027) y otras datan de años anteriores (ej: 2025), regístralo como alerta crítica.
   - **Aforos y Parámetros Operativos**: En pólizas de eventos o actividades con aforo, audita las discrepancias de número de asistentes cotizados (ej: 23 personas en Chubb vs 30 personas en Mundial) y alerta sobre riesgos de inoperancia por exceso de aforo.
   - **Valores Asegurables y Cimientos**: Compara si los bienes declarados coinciden entre aseguradoras o si alguna excluye cimientos o presenta inconsistencias aritméticas.
   - **Garantías y Requisitos Operativos**: Extrae obligaciones de seguridad física y humana exigidas por cada compañía (ej. paramédico en staff, ambulancia medicalizada en sitio, extintores bajo norma técnica, supresores de voltaje, vigilancia física 24h, protocolos de trabajos en caliente, SARLAFT y restricciones OFAC/internacionales).
   - **Condiciones de Pago**: Identifica si se exige pago de prima inmediato o previo al inicio del evento como condición de validez de la cobertura.

---

## ADAPTACIÓN DINÁMICA SEGÚN EL RAMO

Adapta las tablas de las **Secciones IV (Coberturas)** y **V (Deducibles)** según el ramo analizado:

- **A. Si es COPROPIEDAD / MULTIRRIESGO COMERCIAL**:
  - *Valores asegurables*: Edificio (áreas comunes y/o privadas), Contenidos (muebles y enseres), Maquinaria y equipo, Total valor asegurable.
  - *Daños Materiales*: Todo riesgo daño material; Terremoto, temblor y erupción; HMACC / AMIT / Terrorismo; Rotura de vidrios; Hurto calificado; Hurto simple; Daño interno maquinaria; Daño interno equipo eléctrico; Pérdida de ingresos / cuotas de administración.
  - *Sublímites*: Remoción de escombros; Gastos de extinción; Honorarios profesionales; Archivos / portadores de datos; Amparo automático de nuevos bienes; Bienes bajo cuidado o de terceros; Pérdida de arrendamiento.
  - *RCE Copropiedad*: Límite general (LUC); Contratistas y subcontratistas; RC patronal; RC cruzada; Vehículos propios y no propios; Bienes bajo cuidado; Gastos médicos; Gastos de defensa.
  - *Manejo*: Fraude / Apropiación indebida (límite asegurado).
  - *D&O*: Directores y Administradores (límite máximo de responsabilidad).

- **B. Si es RCE EVENTOS / TRC / INGENIERÍA (o formato institucional estándar)**:
  - *Cobertura Sección 1 (Todo Riesgo Construcción)*: Daños materiales (Cobertura A); Terremoto (B); Tormenta e inundación (C); Mantenimiento amplio (D); Remoción de escombros (G); Huelga y asonada; Hurto calificado; Coberturas adicionales de ingeniería.
    *(Nota técnica: Si el riesgo evaluado es un evento puntual o solo RCE y no una obra civil, lista las filas de la Sección 1 marcadas como `NO ESPECIFICA` en todas las columnas para conservar la integridad del catálogo del corredor, y explica la salvedad en el Anexo 2).*
  - *Cobertura Sección 2 (Responsabilidad Civil Extracontractual)*: Límite general por evento y vigencia; Contratistas y subcontratistas; Civil patronal; Civil cruzada; Vehículos propios y no propios; Bienes bajo cuidado, tenencia y control; Contaminación súbita e imprevista; Gastos médicos; Gastos de defensa; Coberturas específicas del evento (parqueaderos, alimentos y bebidas, etc.).
  - *Cobertura Sección 3*: Terrorismo / Amparos adicionales.

- **C. Si es OTRO RAMO**:
  - Organiza las tablas reflejando fielmente los amparos básicos, coberturas adicionales, sublímites y deducibles de las pólizas presentadas.

---

## ESTRUCTURA DEL INFORME DE SALIDA (ESTRICTAMENTE EN MARKDOWN)

Genera la respuesta con el siguiente orden y encabezados:

# mrc - multiriesgos de colombia

### I. INFORMACIÓN GENERAL
| CAMPO | DETALLE |
| :--- | :--- |
| **FECHA** | [Fecha de elaboración del comparativo, DD/MM/AAAA] |
| **TIPO DE COBERTURA** | [Ramo o modalidad exacta de la póliza] |
| **TOMADOR** | [Razón social del Tomador] - NIT [Número de NIT] |
| **ASEGURADO** | [Razón social del Asegurado(s) y NIT(s)] |
| **BENEFICIARIO** | [Terceros afectados / Copropietarios / Asistentes según corresponda] |
| **VIGENCIA** | [Periodo de vigencia: Desde DD/MM/AAAA Hasta DD/MM/AAAA] |
| **UBICACIÓN** | [Lugar exacto del riesgo, Municipio, Departamento] |
| **VALOR ASEGURADO** | [Monto total declarado o 'NO ESPECIFICA' si es póliza a primer riesgo / RCE sin suma asegurable global] |
| **DESCRIPCIÓN DEL PROYECTO / RIESGO** | [Descripción de la actividad o predio, aforo cotizado por compañía, modalidad de cobertura, jurisdicción] |

---

### II. COTIZACIONES REALIZADAS
| COMPAÑÍA DE SEGUROS | TASA | PRIMA | MODALIDAD DE ASEGURAMIENTO |
| :--- | :---: | :---: | :--- |
(Registra cada aseguradora y opciones. En PRIMA consigna la prima neta. En MODALIDAD resume la forma de aseguramiento [por ocurrencia / claims made], desglose de prima neta, gastos de expedición, IVA, total a pagar y condición de pago de la prima).

---

### III. RECOMENDACIÓN
**[NOMBRE DE LA ASEGURADORA Y RAMO / OPCIÓN SELECCIONADA]**
- **Comparativo Económico**: Comparación de prima total a pagar con IVA y porcentaje o valor absoluto de ahorro entre las alternativas.
- **Capacidad y Parámetros Operativos**: Comparación técnica de aforos cubiertos, vigencias de cobertura o valores asegurables.
- **Amplitud de Cobertura y Sublímites**: Coberturas adicionales incluidas vs excluidas (ej. presencia de D&O/Manejo, o amparo de alimentos/bebidas, parqueaderos, etc.).
- **Deducibles y Contrapartidas Técnicas**: Comparación de deducibles mínimos (en SMMLV o porcentaje) y sublímites en amparos críticos (patronal, cruzada, vehículos).
- **Perfil de Riesgo y Conclusión**: Balance técnico-económico indicando cuál oferta representa la mejor alternativa según las prioridades del cliente.

---

### IV. COBERTURAS
Tabla comparativa tabular estructurada:
| CONCEPTO | [ASEGURADORA 1] | [ASEGURADORA 2] | [ASEGURADORA 3] |

(Estructura las filas agrupando por las secciones del ramo aplicable [Daños Materiales, RCE, Manejo, D&O, o Sección 1 TRC, Sección 2 RCE, etc.] con títulos en negrita en la columna CONCEPTO).

---

### V. DEDUCIBLES
Tabla comparativa tabular estructurada:
| CONCEPTO | [ASEGURADORA 1] | [ASEGURADORA 2] | [ASEGURADORA 3] |

(Estructura los deducibles correspondientes a cada amparo o sección evaluada).

---

### VI. PREVALENCIA DE LOS TÉRMINOS COTIZADOS
Transcribe exactamente el siguiente texto institucional:

> "La información sobre el listado de Coberturas, Deducibles, Valores Asegurados, Cláusulas Adicionales, Sublímites, Primas, Tasas, Conclusiones, etc., es meramente ilustrativa y fue tomada de la cotización original y en firme presentada por el mercado asegurador. Los ítems mencionados se toman en forma parcial y a manera de ejemplo para facilitar la comparación, análisis y conclusión de lo que podría constituirse en la cotización más favorable al riesgo que se pretende trasladar.
>
> Multiriesgos de Colombia Ltda. expresamente hace extensivo al cliente los términos y condiciones oficiales de cada una de las Compañías de Seguros que demostraron su interés en el riesgo planteado y que presentaron su propuesta. En consecuencia, se deja establecido y expresamente pactado que, para todos los efectos técnicos y legales, prevalecen las Condiciones Técnicas y Económicas presentadas oficialmente por cada una de las Compañías de Seguros colombianas."

---

### ANEXO 1: SUBJETIVIDADES Y CONDICIONES PARTICULARES POR ASEGURADORA
Detalla individualmente por aseguradora mediante viñetas:
- **Garantías y Requisitos de Seguridad**: Medidas obligatorias de prevención física y humana (extintores, mantenimiento, ambulancia en sitio, paramédicos, apoyo de bomberos/policía, control de aforo).
- **Documentos Requeridos**: SARLAFT, RUT, Cámara de Comercio, estados financieros, etc.
- **Condiciones Operativas y Sublímites Específicos**: Sublímites de gastos médicos, alimentos/bebidas, parqueaderos, etc.
- **Validez de la Oferta y Condiciones Precedentes**: Días de vigencia de la cotización y condicionamientos previos al desembolso o inicio de vigencia.
- **Restricciones Internacionales o Cláusulas Especiales**: Sanciones OFAC/ONU/UE, cláusulas de infraseguro, etc.

---

### ANEXO 2: OBSERVACIONES Y PREGUNTAS DEL CLIENTE
- **Alertas Técnicas y Discrepancias**:
  - Diferencias en aforos cotizados (ej. 23 vs 30 personas), desfases de fechas de vigencia o variaciones en la tasación del riesgo.
  - Aclaraciones sobre el alcance de la póliza (ej. precisión de que se trata de RCE para evento y no obra civil, por lo que filas de TRC no aplican).
- **Exclusiones Críticas Relevantes**: Detalle de exclusiones sensibles compartidas o particulares (sustancias psicoactivas, embriaguez, COVID-19/enfermedades transmisibles, terrorismo, RC profesional, pérdidas financieras puras).
- **Preguntas Técnicas y Acciones para el Cliente**: Confirmaciones necesarias antes de la expedición formal (confirmación de aforo real, pago oportuno de la prima, envío de documentos habilitantes)."""


# ---------------------------------------------------------------------------
# Specialized Extraction Prompts (Structured JSON Schemas)
# ---------------------------------------------------------------------------

EXTRACTION_SYSTEM_PROMPT = """Eres el Analista Técnico Senior de Suscripción de Multiriesgos de Colombia Ltda., corredor de seguros profesional. Tu objetivo es extraer con fidelidad absoluta los datos técnicos y económicos de una cotización de seguros colombiana.

## REGLAS DE SUSCRIPCIÓN (CERO ALUCINACIONES)
1. **Fidelidad Absoluta y Transcripción Literal**:
   - Transcribe fielmente valores numéricos, porcentajes y condiciones textuales. Nunca redondees, deduzcas ni inventes coberturas.
   - El nombre de la compañía (`nombre_compania`) debe ser su razón social oficial (ej: "Zurich Colombia Seguros S.A.", "AXA Colpatria Seguros S.A.", "Chubb Seguros Colombia S.A.", "Compañía Suramericana de Seguros S.A.", "Seguros Mundial", "HDI Seguros S.A.", "Seguros del Estado S.A.", "Seguros Bolívar S.A.").
   - El ID de la compañía (`id_compania`) debe ser normalizado en minúsculas: 'sura', 'chubb', 'axa_colpatria', 'mundial', 'hdi', 'zurich', 'bolivar', 'estado', 'allianz', 'mapfre', 'solidaria', 'previsora', 'berkley'.
2. **Finanzas y Primas (Desglose Estricto)**:
   - Campo `prima` en `opciones`: Consigna SIEMPRE el valor numérico en pesos colombianos de la **Prima Neta antes de IVA** (o prima antes de tributos). Float sin separadores.
   - Campo `modalidad` en `opciones`: Registra la forma de aseguramiento (ej: "Ocurrencia", "Claims Made", "Primer Riesgo", "LEG 2/96"), seguido del desglose financiero explícito:
     * Prima Neta: $...
     * Gastos de Emisión/Expedición: $...
     * IVA (19%): $...
     * Total a Pagar: $...
     * Condición de Pago: (ej: Pago inmediato previo al inicio del evento / 30 días).
   - Campo `tasa` en `opciones`: Exprésala en por mil (`0,949 ‰`), porcentaje (`0,095 %`), por millón o modalidad tarifaria (`Prima fija por evento`, `Tabla por millón/día`). Si no se especifica, usa estrictamente: "Prima única NO ESPECIFICA".
3. **Manejo de Opciones**:
   - Si la aseguradora ofrece múltiples opciones (ej: Opción 1 vs Opción 2, o con/sin cláusulas LEG), crea una entrada independiente en la lista `opciones` para cada una, identificando su etiqueta, tasa, prima neta y modalidad con desglose.
4. **Convenciones Estrictas de Silencio o Exclusión**:
   - `NO ESPECIFICA`: Úsalo en coberturas/deducibles cuando el documento guarde silencio absoluto.
   - `N/A`: Úsalo cuando la aseguradora exprese formalmente que la cobertura no aplica, no contrata o esté excluida.
   - `No cotizado`: Úsalo cuando la cotización incluya el campo en su plantilla pero asigne $0, período de indemnización 0, o exprese textualmente "No cotizado".
5. **Deducibles**:
   - SIEMPRE texto (string). Transcribe porcentaje y mínimo tal cual aparecen en la cotización (ej: "10% del siniestro, mínimo 5 SMMLV", "10% valor de la pérdida, mínimo $5.000.000 COP"). NO conviertas SMMLV a pesos.
6. **Garantías y Subjetividades (`subjetividades`)**:
   - Extrae rigurosamente: requisitos de seguridad humana y física (ambulancia medicalizada en sitio, paramédicos en staff, apoyo de bomberos/policía, extintores bajo norma técnica, supresores de voltaje, vigilancia física 24h, protocolos de trabajo en caliente), documentos habilitantes (SARLAFT, RUT, Cámara de Comercio), restricciones internacionales (OFAC/ONU), condición de pago previo y validez de la oferta.
7. **Coaseguro**:
   - Si la cotización es en coaseguro (ej: Seguros Mundial con Berkley), lista cada compañía participante con su porcentaje exacto de participación.

## CATÁLOGO EXTENDIDO DE COBERTURAS (claves en "coberturas")
Adapta las claves según el ramo cotizado:
- **Copropiedades / Multirriesgo Comercial**:
  danos_materiales, terremoto, hmacc_amit, rotura_vidrios, hurto_calificado, hurto_simple, dano_interno_maquinaria, dano_interno_electrico, perdida_ingresos_cuotas, remocion_escombros, gastos_extincion, honorarios_profesionales, archivos_datos, nuevos_bienes, bienes_cuidado_terceros, perdida_arrendamiento, rce, rce_contratistas, rce_patronal, rce_cruzada, rce_vehiculos, rce_cuidado_control, rce_gastos_medicos, rce_gastos_defensa, manejo_fraude, directores_administradores.
- **RCE Eventos / TRC / Ingeniería**:
  danos_materiales, terremoto, tormenta_inundacion, mantenimiento_amplio, remocion_escombros, hmacc_amit, hurto_calificado, cronograma_avance, gastos_horas_extra, obras_zona_sismica, bienes_fuera_sitio, prueba_maquinaria, campamentos_almacenes, medidas_inundacion, proteccion_incendio, transportes_nacionales, siniestros_serie, obras_civiles_operacion, cimentacion_pilotaje, hundimiento_subsuelo, error_diseno, honorarios_profesionales, planos_documentos, gastos_extincion, condiciones_remocion, actos_autoridad, preservacion_bienes, rce, rce_contratistas, rce_patronal, rce_cruzada, rce_vehiculos, rce_cuidado_control, rce_contaminacion, rce_vibracion, rce_subterraneas, terrorismo, rce_alimentos_bebidas, rce_parqueaderos, rce_gastos_medicos, rce_gastos_defensa.
- **Otros Ramos**:
  Utiliza las claves que mejor representen el amparo básico, coberturas adicionales y sublímites del ramo.

## CATÁLOGO EXTENDIDO DE DEDUCIBLES (claves en "deducibles")
ded_incendio, ded_terremoto, ded_tormenta, ded_mantenimiento, ded_remocion, ded_hmacc, ded_hurto, ded_error_diseno, ded_cables, ded_adyacentes, ded_campamentos, ded_hundimiento, ded_aeronaves, ded_impericia, ded_corto_circuito, ded_rce, ded_rce_cruzada, ded_propiedades_existentes, ded_vidrios, ded_dano_interno, ded_manejo, ded_dno, ded_rotura_maquinaria.

Extrae toda la información con fidelidad documental absoluta y devuelve el JSON estructurado."""


META_EXTRACTION_PROMPT = """Eres el Analista Técnico Senior de Suscripción de Multiriesgos de Colombia Ltda. Analiza las siguientes cotizaciones de seguros colombianas y extrae la metadata general del proyecto, asegurado y riesgo.

## REGLAS DE EXTRACCIÓN (SECCIÓN I - INFORMACIÓN GENERAL)
1. **Tipo de Cobertura**: Identifica el ramo y modalidad exacta (ej: "TODO RIESGO CONSTRUCCIÓN Y MONTAJE", "MULTIRRIESGO COPROPIEDADES / COMERCIAL", "RESPONSABILIDAD CIVIL EXTRACONTRACTUAL EVENTOS", "RCE OPERACIONES").
2. **Tomador**: Razón social del Tomador incluyendo NIT si aparece en los documentos (ej: "Empresa S.A.S. - NIT 900.123.456-7"). Si no aparece, deja cadena vacía.
3. **Asegurado**: Razón social del Asegurado(s) y sus NITs.
4. **Beneficiario**: Asigna según el ramo (ej: "Terceros afectados", "Copropietarios", "Entidades financieras / Acreedores hipotecarios", "Asistentes al evento").
5. **Vigencias**:
   - `vigencia_construccion`: {'desde': 'DD/MM/AAAA', 'hasta': 'DD/MM/AAAA'}. Si hay discrepancia de fechas entre aseguradoras (ej: vigencias 2026 vs 2025), regístralo con fidelidad.
   - `vigencia_mantenimiento`: si aplica, duración y periodo.
6. **Ubicación**: Lugar exacto del riesgo, Municipio, Departamento.
7. **Valor Asegurado**: Monto total declarado sin separadores (número float/int). Si la póliza es a primer riesgo o RCE sin suma global asegurable de bienes, deja null.
8. **Descripción del Proyecto / Riesgo**: Detalle de la actividad, alcance, predio asegurado, aforos cotizados por aseguradora (ej. "Aforo de 23 a 30 personas según propuesta"), modalidad y jurisdicción aplicable.

Devuelve el JSON estructurado según el esquema."""


RECOMMENDATION_PROMPT = """Eres el Analista Técnico Senior de Suscripción de Multiriesgos de Colombia Ltda. Analiza los datos comparativos consolidados de las propuestas presentadas por las aseguradoras y genera una RECOMENDACIÓN TÉCNICA OFICIAL, estructurada y fundamentada para el cliente.

## CRITERIOS DE EVALUACIÓN (SECCIÓN III DEL ESTÁNDAR)
1. **Comparativo Económico**:
   - Compara las primas totales a pagar (incluyendo IVA y gastos de emisión).
   - Calcula el valor absoluto en COP y porcentaje de ahorro entre la opción más competitiva y las demás alternativas.
2. **Capacidad y Parámetros Operativos**:
   - Evalúa aforos cubiertos, vigencias amparadas (detectando posibles desfases) y suficiencia de los valores asegurables.
3. **Amplitud de Cobertura y Sublímites**:
   - Identifica amparos críticos incluidos vs omitidos o excluidos (ej. amparo de alimentos y bebidas, parqueaderos, D&O, Manejo, coberturas de contratistas o patronal).
4. **Deducibles y Contrapartidas Técnicas**:
   - Compara los deducibles mínimos en SMMLV o en pesos y los porcentajes sobre la pérdida en eventos críticos (patronal, cruzada, terremoto).
5. **Perfil de Riesgo y Conclusión**:
   - Emite una conclusión objetiva que pondere costo-beneficio, solvencia de la aseguradora / coaseguro y cumplimiento de las condiciones requeridas por el cliente.

## REGLAS
- Identifica claramente el `aseguradora_id` y la `opcion` recomendada.
- Genera en `vinetas` argumentos sólidos, objetivos y verificables agrupados por los 5 criterios anteriores.
- Cero alucinaciones: NO inventes coberturas ni valores no presentes en los datos.
- Incluye alertas sobre requisitos indispensables de previo cumplimiento (pago antes del evento, SARLAFT, paramédicos, etc.).

Datos comparativos consolidados:"""


WORST_MARKERS_PROMPT = """Eres el Analista Técnico Senior de Suscripción de Multiriesgos de Colombia Ltda. Analiza los datos consolidados de las cotizaciones y determina qué aseguradoras presentan la PEOR oferta técnica o económica en cada fila de cobertura y deducible.

## REGLAS ESTRICTAS (CERO ALUCINACIONES)
1. Marca como "peor" ÚNICAMENTE cuando la diferencia sea CLARAMENTE inferior, desfavorable y técnicamente defendible ante el cliente (ej: deducible mínimo de 15 SMMLV vs 5 SMMLV, exclusión expresa "N/A" vs cobertura amplia, o sublímite significativamente recortado).
2. Si todas las aseguradoras ofrecen condiciones idénticas o similares en un rubro, NO incluyas esa clave.
3. Usa los IDs normalizados de aseguradora (ej: "sura", "chubb", "mundial", "axa_colpatria", "hdi", "zurich").
4. Devuelve un diccionario donde la clave es la clave de cobertura o deducible y el valor es la lista de IDs de aseguradoras con la peor condición.
5. REGLA FUNDAMENTAL DE SILENCIO: NO marques rubros donde una aseguradora tenga 'NO ESPECIFICA' (el silencio no equivale a peor oferta declarada). Solo marca cuando haya un valor explícito desfavorable o 'N/A' (exclusión formal frente a inclusión).

Datos consolidados:"""


# ---------------------------------------------------------------------------
# Client Management (Singleton / Cached Pattern)
# ---------------------------------------------------------------------------

_client: genai.Client | None = None


def _get_client() -> genai.Client:
    """Get or initialize the cached Gemini API client.

    Reuses the client instance across calls to avoid connection churn and overhead.
    """
    global _client
    if _client is None:
        if not settings.GEMINI_API_KEY:
            raise LLMExtractionError(
                message="GEMINI_API_KEY no está configurada en las variables de entorno."
            )
        _client = genai.Client(api_key=settings.GEMINI_API_KEY)
    return _client


def _reset_client() -> None:
    """Reset the cached client instance (useful for testing or key rotation)."""
    global _client
    _client = None


# ---------------------------------------------------------------------------
# Core Extraction Functions (Map Phase)
# ---------------------------------------------------------------------------

async def extract_quote(
    document_text: str = "",
    file_bytes: bytes | None = None,
    filename: str = "",
    categoria: str = "Todo_Riesgo_Construccion",
    prompt_file: str | Path | None = None,
) -> ExtractedQuote:
    """Extract structured quote data from a document using Gemini.

    This is the MAP step of the Map-Reduce pattern. Uses structured output schema
    and temperature 0.0 to ensure deterministic, hallucination-free extraction.
    Supports both multimodal PDF bytes (for scans and vector layouts) and plain text.

    Args:
        document_text: Plain text extracted from a quote document (if any).
        file_bytes: Raw binary bytes of the file for multimodal processing.
        filename: Name of the file being processed.
        categoria: Insurance category (e.g. 'Autos', 'Copropiedades', 'Hogar', 'Todo_Riesgo_Construccion').
        prompt_file: Optional explicit path to custom prompt file.

    Returns:
        Structured extraction of the quote adhering to ExtractedQuote schema.

    Raises:
        LLMExtractionError: If extraction or schema validation fails.
    """
    try:
        client = _get_client()
        system_instruction = load_prompt_for_category(
            categoria=categoria, prompt_file=prompt_file
        )

        contents_list: list[Any] = []
        is_pdf = (filename.lower().endswith(".pdf")) or (file_bytes is not None and file_bytes[:4] == b"%PDF")
        if file_bytes and is_pdf:
            try:
                pdf_part = types.Part.from_bytes(data=file_bytes, mime_type="application/pdf")
                contents_list.append(pdf_part)
            except Exception as e:
                logger.warning("No se pudo adjuntar PDF binario como Part: %s", e)

        prompt_text = f"## COTIZACIÓN DE SEGUROS ({categoria.upper()}) A ANALIZAR:\n\n"
        if document_text and document_text.strip() and not document_text.startswith("[PDF"):
            prompt_text += document_text[:35000]
        else:
            prompt_text += "Analiza el documento adjunto y extrae fielmente todos los datos de cotización solicitados."
        contents_list.append(prompt_text)

        response = await client.aio.models.generate_content(
            model=settings.GEMINI_MODEL,
            contents=contents_list,
            config=types.GenerateContentConfig(
                system_instruction=system_instruction,
                temperature=0.0,
                response_mime_type="application/json",
                response_json_schema=ExtractedQuote.model_json_schema(),
            ),
        )

        if not response.text:
            raise LLMExtractionError(
                message="Gemini no retornó texto en la respuesta de extracción de cotización"
            )

        result = ExtractedQuote.model_validate_json(response.text)
        logger.info(
            "Cotización extraída exitosamente: %s (ID: %s, Ramo/Cat: %s, %d opciones, %d coberturas, %d deducibles)",
            result.nombre_compania,
            result.id_compania,
            categoria,
            len(result.opciones),
            len(result.coberturas),
            len(result.deducibles),
        )
        return result

    except LLMExtractionError:
        raise
    except Exception as e:
        logger.exception("Fallo al extraer cotización con Gemini: %s", e)
        raise LLMExtractionError(
            message=f"Error al extraer cotización con Gemini: {e}",
            detail=str(e),
        ) from e


async def extract_project_meta(
    document_texts: list[str],
    categoria: str = "Todo_Riesgo_Construccion",
    quotes: list[ExtractedQuote] | None = None,
    prompt_file: str | Path | None = None,
) -> ProjectMeta:
    """Extract project and risk metadata across all quote documents.

    Consolidates general risk information (Insured, Policyholder, Location,
    Policy Type, Insured Value, Vigencias) from multiple proposals and category-specific fields.

    Args:
        document_texts: List of plain text from all quote documents.
        categoria: Insurance category (e.g. 'Autos', 'Copropiedades', 'Hogar', 'Todo_Riesgo_Construccion').
        quotes: Optional list of ExtractedQuote objects previously extracted.
        prompt_file: Optional explicit path to custom prompt file.

    Returns:
        Inferred project metadata adhering to ProjectMeta schema.

    Raises:
        LLMExtractionError: If metadata extraction fails.
    """
    try:
        client = _get_client()
        cat_prompt = load_prompt_for_category(
            categoria=categoria, prompt_file=prompt_file
        )
        meta_instruction = (
            f"{META_EXTRACTION_PROMPT}\n\n"
            f"REGLAS Y CAMPOS ESPECÍFICOS PARA LA CATEGORÍA '{categoria}':\n"
            f"{cat_prompt[:3000]}"
        )

        content_sections: list[str] = []
        if quotes:
            q_info = []
            for q in quotes:
                cobs_sample = {k: v for k, v in list(q.coberturas.items())[:8]}
                q_info.append(
                    f"Aseguradora: {q.nombre_compania} (ID: {q.id_compania})\n"
                    f"Validez: {q.validez_oferta}\n"
                    f"Coberturas clave: {cobs_sample}\n"
                )
                if q.subjetividades:
                    q_info.append(f"Subjetividades/Garantías: {q.subjetividades[:3]}\n")
            content_sections.append("COTIZACIONES PREVIAMENTE EXTRAÍDAS:\n" + "\n".join(q_info))

        valid_texts = [t[:4000] for t in document_texts if t.strip() and not t.startswith("[PDF")]
        if valid_texts:
            combined = "\n\n---\n\n".join(
                f"TEXTO DOCUMENTO {i + 1}:\n{text}"
                for i, text in enumerate(valid_texts)
            )
            content_sections.append(combined)

        if not content_sections:
            content_sections.append(f"Extrae la metadata del riesgo para la categoría {categoria}.")

        response = await client.aio.models.generate_content(
            model=settings.GEMINI_MODEL,
            contents=["\n\n".join(content_sections)],
            config=types.GenerateContentConfig(
                system_instruction=meta_instruction,
                temperature=0.0,
                response_mime_type="application/json",
                response_json_schema=ProjectMeta.model_json_schema(),
            ),
        )

        if not response.text:
            raise LLMExtractionError(
                message="Gemini no retornó metadata del proyecto"
            )

        meta = ProjectMeta.model_validate_json(response.text)
        logger.info(
            "Metadata del proyecto extraída: Ramo='%s', Asegurado='%s', Ubicación='%s'",
            meta.tipo_cobertura,
            meta.asegurado,
            meta.ubicacion,
        )
        return meta

    except LLMExtractionError:
        raise
    except Exception as e:
        logger.exception("Fallo al extraer metadata del proyecto: %s", e)
        raise LLMExtractionError(
            message=f"Error al extraer metadata del proyecto: {e}",
            detail=str(e),
        ) from e


# ---------------------------------------------------------------------------
# Underwriting Analysis Functions (Reduce Phase)
# ---------------------------------------------------------------------------

async def generate_recommendation(consolidated_data: dict) -> Recommendation:
    """Generate an AI recommendation based on consolidated comparative data.

    Evaluates proposals using Multiriesgos de Colombia's 5 core underwriting criteria:
    1. Financial comparison (premiums with VAT and savings)
    2. Capacity & operational limits (capacity/attendance, timelines)
    3. Coverage breadth & sublimits
    4. Deductibles (percentages & SMMLV minimums)
    5. Risk profile & final underwriting conclusion

    Args:
        consolidated_data: The full consolidated data dict.

    Returns:
        Recommendation with insurer ID, option, and justification bullets.

    Raises:
        LLMExtractionError: If recommendation generation fails.
    """
    try:
        client = _get_client()

        summary = json.dumps(consolidated_data, ensure_ascii=False, indent=2)

        response = await client.aio.models.generate_content(
            model=settings.GEMINI_MODEL,
            contents=[
                f"DATOS COMPARATIVOS CONSOLIDADOS DEL RIESGO:\n\n{summary}",
            ],
            config=types.GenerateContentConfig(
                system_instruction=RECOMMENDATION_PROMPT,
                temperature=0.0,
                response_mime_type="application/json",
                response_json_schema=Recommendation.model_json_schema(),
            ),
        )

        if not response.text:
            raise LLMExtractionError(
                message="Gemini no retornó recomendación para el comparativo"
            )

        rec = Recommendation.model_validate_json(response.text)
        logger.info(
            "Recomendación técnica generada: Aseguradora='%s', Opción='%s' (%d argumentos)",
            rec.aseguradora_id,
            rec.opcion,
            len(rec.vinetas),
        )
        return rec

    except LLMExtractionError:
        raise
    except Exception as e:
        logger.exception("Fallo al generar recomendación técnica: %s", e)
        raise LLMExtractionError(
            message=f"Error al generar recomendación técnica: {e}",
            detail=str(e),
        ) from e


async def generate_worst_markers(consolidated_data: dict) -> dict[str, list[str]]:
    """Identify the worst offerings per coverage and deductible row.

    Flags insurers whose terms are clearly inferior and defensible, respecting
    the rule that silence (NO ESPECIFICA) is never flagged as worst.

    Args:
        consolidated_data: The full consolidated data dict.

    Returns:
        Dict mapping coverage/deductible keys to lists of insurer IDs.
    """
    try:
        client = _get_client()

        summary = json.dumps(consolidated_data, ensure_ascii=False, indent=2)

        response = await client.aio.models.generate_content(
            model=settings.GEMINI_MODEL,
            contents=[
                f"DATOS CONSOLIDADOS PARA ANÁLISIS DE PEORES CONDICIONES:\n\n{summary}",
            ],
            config=types.GenerateContentConfig(
                system_instruction=WORST_MARKERS_PROMPT,
                temperature=0.0,
                response_mime_type="application/json",
            ),
        )

        if not response.text:
            return {}

        markers = json.loads(response.text)
        if not isinstance(markers, dict):
            return {}

        logger.info("Marcadores de peores condiciones identificados: %d claves", len(markers))
        return markers

    except Exception as e:
        logger.warning("Advertencia al generar marcadores peor (no crítico): %s", e)
        return {}


# ---------------------------------------------------------------------------
# Official Markdown Comparative Slip Generator
# ---------------------------------------------------------------------------

async def generate_comparative_report(
    document_texts: list[str],
    metadata: dict[str, Any] | None = None,
    categoria: str = "Todo_Riesgo_Construccion",
    prompt_file: str | Path | None = None,
) -> str:
    """Generate the official Technical Comparative Report (Slip Comparativo) in Markdown.

    Produces the exact official 6 sections and 2 technical annexes established by
    Multiriesgos de Colombia Ltda., adapted to the selected insurance category.

    Args:
        document_texts: Plain text of all insurance quotes to compare.
        metadata: Optional contextual metadata (e.g., tomador override, fecha).
        categoria: Insurance category (e.g. 'Autos', 'Copropiedades', 'Hogar', 'Todo_Riesgo_Construccion').
        prompt_file: Optional explicit path to custom prompt file.

    Returns:
        The complete comparative report strictly formatted in Markdown.

    Raises:
        LLMExtractionError: If report generation fails.
    """
    try:
        client = _get_client()
        system_instruction = load_prompt_for_category(
            categoria=categoria, prompt_file=prompt_file
        )

        combined_docs = "\n\n" + "=" * 60 + "\n\n".join(
            f"COTIZACIÓN ASEGURADORA {i + 1}:\n{text}"
            for i, text in enumerate(document_texts)
        )

        context_info = ""
        if metadata:
            context_info = f"\nMETADATA ADICIONAL DEL CORREDOR:\n{json.dumps(metadata, ensure_ascii=False, indent=2)}\n\n"

        prompt_input = (
            f"{context_info}"
            f"Procesa las siguientes cotizaciones adjuntas de la categoría '{categoria}' y elabora el Informe Comparativo Oficial "
            f"(Slip Comparativo de Coaseguro/Mercado) con las 6 Secciones Principales y 2 Anexos Técnicos "
            f"estrictamente en formato Markdown conforme a tus instrucciones:\n\n{combined_docs}"
        )

        response = await client.aio.models.generate_content(
            model=settings.GEMINI_MODEL,
            contents=[prompt_input],
            config=types.GenerateContentConfig(
                system_instruction=system_instruction,
                temperature=0.0,
            ),
        )

        if not response.text:
            raise LLMExtractionError(
                message="Gemini no retornó texto para el informe comparativo en Markdown"
            )

        logger.info(
            "Informe comparativo en Markdown generado exitosamente (%d caracteres, categoría: %s)",
            len(response.text),
            categoria,
        )
        return response.text

    except LLMExtractionError:
        raise
    except Exception as e:
        logger.exception("Fallo al generar el informe comparativo en Markdown: %s", e)
        raise LLMExtractionError(
            message=f"Error al generar informe comparativo con Gemini: {e}",
            detail=str(e),
        ) from e


async def generate_comparative_report_from_data(
    consolidated_data: dict[str, Any],
    categoria: str = "Todo_Riesgo_Construccion",
    prompt_file: str | Path | None = None,
) -> str:
    """Generate the official Technical Comparative Report in Markdown from consolidated data.

    Allows producing the standardized Markdown slip directly from a pre-consolidated
    ConsolidatedData dictionary (equivalent to datos.json).

    Args:
        consolidated_data: Consolidated comparative data dictionary.
        categoria: Insurance category (e.g. 'Autos', 'Copropiedades', 'Hogar', 'Todo_Riesgo_Construccion').
        prompt_file: Optional explicit path to custom prompt file.

    Returns:
        The official comparative report formatted strictly in Markdown.

    Raises:
        LLMExtractionError: If report generation fails.
    """
    try:
        client = _get_client()
        system_instruction = load_prompt_for_category(
            categoria=categoria, prompt_file=prompt_file
        )

        summary = json.dumps(consolidated_data, ensure_ascii=False, indent=2)

        prompt_input = (
            f"Genera el Informe Comparativo Oficial (Slip Comparativo de Coaseguro/Mercado) "
            f"para la categoría '{categoria}' con las 6 Secciones Principales y 2 Anexos Técnicos "
            f"estrictamente en formato Markdown, utilizando los siguientes datos técnicos y económicos ya consolidados:\n\n{summary}"
        )

        response = await client.aio.models.generate_content(
            model=settings.GEMINI_MODEL,
            contents=[prompt_input],
            config=types.GenerateContentConfig(
                system_instruction=system_instruction,
                temperature=0.0,
            ),
        )

        if not response.text:
            raise LLMExtractionError(
                message="Gemini no retornó contenido para el informe comparativo desde datos consolidados"
            )

        logger.info(
            "Informe comparativo generado desde datos consolidados (%d caracteres)",
            len(response.text),
        )
        return response.text

    except LLMExtractionError:
        raise
    except Exception as e:
        logger.exception("Fallo al generar informe comparativo desde datos consolidados: %s", e)
        raise LLMExtractionError(
            message=f"Error al generar informe comparativo desde datos consolidados: {e}",
            detail=str(e),
        ) from e
