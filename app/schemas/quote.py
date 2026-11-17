# -*- coding: utf-8 -*-
"""Pydantic v2 schemas for individual quote extraction and project metadata."""
from __future__ import annotations

import re
from typing import Any, Optional, Union
from pydantic import BaseModel, Field, field_validator


def _parse_optional_float(val: Any) -> Optional[float]:
    """Safely coerce numeric values or formatted COP strings into float, or None."""
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        s = val.strip()
        if not s or s.upper() in {"NO ESPECIFICA", "N/A", "NONE", "NULL", "-"}:
            return None
        # Strip currency symbols and thousand separators
        cleaned = re.sub(r"[^\d,\.]", "", s)
        if not cleaned:
            return None
        # If multiple dots exist (e.g. 5.149.392.347), dots are thousand separators
        if cleaned.count(".") > 1:
            cleaned = cleaned.replace(".", "").replace(",", ".")
        elif "." in cleaned and "," in cleaned:
            if cleaned.rfind(",") > cleaned.rfind("."):
                cleaned = cleaned.replace(".", "").replace(",", ".")
            else:
                cleaned = cleaned.replace(",", "")
        elif "." in cleaned:
            parts = cleaned.split(".")
            if len(parts[-1]) == 3 and len(parts[0]) >= 1:
                cleaned = cleaned.replace(".", "")
        elif "," in cleaned:
            parts = cleaned.split(",")
            if len(parts[-1]) == 3 and len(parts[0]) >= 1:
                cleaned = cleaned.replace(",", "")
            else:
                cleaned = cleaned.replace(",", ".")
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


class CoaseguroEntry(BaseModel):
    """Entry for a co-insurance participant."""

    nombre: str = Field(default="", description="Nombre de la compañía participante en el coaseguro")
    participacion: str = Field(default="", description="Porcentaje de participación, ej: '35%'")
    logo: Optional[str] = Field(default=None, description="Nombre del archivo de logo, ej: 'mundial.png'")

    @field_validator("nombre", "participacion", mode="before")
    @classmethod
    def _coerce_str(cls, v: Any) -> str:
        return "" if v is None else str(v).strip()


class QuoteOption(BaseModel):
    """A pricing option within a quote."""

    etiqueta: str = Field(default="Opción 1", description="Identificador de la opción, ej: 'Opción 1', 'Opción 2'")
    tasa: str = Field(default="", description="Tasa en por mil (‰), ej: '1,10 ‰' (vacío si no aplica al ramo)")
    prima: float = Field(
        default=0.0,
        description=(
            "Valor numérico de la prima en pesos colombianos, sin separadores de miles. "
            "En Copropiedades y Hogar registrar SIEMPRE el TOTAL A PAGAR INCLUIDO IVA."
        ),
    )
    modalidad: str = Field(
        default="",
        description=(
            "Modalidad de aseguramiento según el ramo. "
            "En Copropiedades: 'Deducible terremoto sobre\\nValor Asegurable' o "
            "'Deducible terremoto sobre\\nValor de la Pérdida' según corresponda a esta opción. "
            "En TRC: ej. 'LEG 2/96 100% del valor asegurado'."
        ),
    )

    @field_validator("etiqueta", "tasa", "modalidad", mode="before")
    @classmethod
    def _coerce_opt_str(cls, v: Any) -> str:
        return "" if v is None else str(v)

    @field_validator("prima", mode="before")
    @classmethod
    def _coerce_prima(cls, v: Any) -> float:
        parsed = _parse_optional_float(v)
        return parsed if parsed is not None else 0.0


class ExtractedQuote(BaseModel):
    """Structured extraction of a single insurance quote from a document.

    This schema is used as the response_json_schema for Gemini structured output.
    Field descriptions guide the LLM extraction.
    """

    nombre_compania: str = Field(
        default="",
        description="Nombre completo oficial de la compañía aseguradora tal como aparece en la cotización",
    )
    id_compania: str = Field(
        default="",
        description="ID normalizado en minúsculas sin espacios ni tildes. Ejemplos: 'sura', 'chubb', 'axa_colpatria', 'mundial', 'hdi', 'zurich', 'bolivar', 'estado', 'mapfre', 'allianz'",
    )
    opciones: list[QuoteOption] = Field(
        default_factory=list,
        description="Lista de opciones de cotización. Cada opción tiene su prima y modalidad específica sin mezclar textos de otras opciones.",
    )
    coberturas: dict[str, Optional[Union[int, float, str]]] = Field(
        default_factory=dict,
        description=(
            "Diccionario de coberturas extraídas usando las claves exactas del catálogo del ramo indicado en las instrucciones del sistema. "
            "Claves Copropiedades: incendio_rayo_explosion_agua, rotura_vidrios, cuotas_administracion, "
            "rotura_maquinaria, sustraccion_violencia, sustraccion_sin_violencia, "
            "equipo_electronico_dano_interno, equipo_electronico_sustraccion, equipo_electronico_sustraccion_sin_violencia, "
            "manejo, muebles_dano_interno, muebles_sustraccion, muebles_sustraccion_sin_violencia, "
            "dno_directores, rce_basico, rce_patronal, rce_contratistas, rce_cruzada, "
            "rce_gastos_medicos, rce_parqueaderos, rce_vehiculos. "
            "Claves TRC: danos_materiales, terremoto, tormenta_inundacion, mantenimiento_amplio, remocion_escombros, "
            "hmacc_amit, hurto_calificado, cronograma_avance, gastos_horas_extra, obras_zona_sismica, "
            "bienes_fuera_sitio, prueba_maquinaria, campamentos_almacenes, medidas_inundacion, "
            "proteccion_incendio, transportes_nacionales, siniestros_serie, obras_civiles_operacion, "
            "cimentacion_pilotaje, hundimiento_subsuelo, error_diseno, honorarios_profesionales, "
            "planos_documentos, gastos_extincion, condiciones_remocion, actos_autoridad, "
            "preservacion_bienes, rce, rce_contratistas, rce_patronal, rce_cruzada, rce_vehiculos, "
            "rce_cuidado_control, rce_contaminacion, rce_vibracion, rce_subterraneas, terrorismo. "
            "Los valores pueden ser numéricos (int/float para montos puros en COP) o texto (str para sublímites/porcentajes/meses). "
            "Si se excluye o no aplica expresamente, usar 'N/A'."
        ),
    )
    deducibles: dict[str, Optional[str]] = Field(
        default_factory=dict,
        description=(
            "Diccionario de deducibles extraídos usando las claves exactas del ramo. "
            "Claves Copropiedades: ded_incendio, ded_explosion_agua, ded_hmacc_amit, ded_terremoto, "
            "ded_rotura_maquinaria, ded_sustraccion_violencia, ded_equipo_electronico_rotura, "
            "ded_equipo_electronico_sustraccion, ded_equipo_electronico_sustraccion_sin_violencia, ded_dno, ded_rce. "
            "Claves TRC: ded_incendio, ded_terremoto, ded_tormenta, ded_mantenimiento, ded_remocion, ded_hmacc, "
            "ded_hurto, ded_error_diseno, ded_cables, ded_adyacentes, ded_campamentos, ded_hundimiento, "
            "ded_aeronaves, ded_impericia, ded_corto_circuito, ded_rce, ded_rce_cruzada, ded_propiedades_existentes. "
            "Los valores siempre son texto: transcribir porcentaje, base (de la pérdida / del Valor Asegurable) y mínimo en SMMLV."
        ),
    )
    asistencias: dict[str, Optional[str]] = Field(
        default_factory=dict,
        description=(
            "Diccionario de servicios de asistencia (obligatorio en Copropiedades, revisando hasta las últimas páginas del PDF). "
            "Claves Copropiedades: cerrajero_comunes, vidriero_comunes, electricista_comunes, plomero_comunes, "
            "traslado_bienes_comunes, celador_sustituto, reparacion_propiedades_vecinas, remocion_escombros_asistencia, "
            "jardineria, reparacion_antena_colectiva, auxiliar_servicios_generales, "
            "orientacion_juridica, emision_conceptos, elaboracion_documentos, "
            "cerrajero_privadas, plomero_privadas, electricista_privadas, vidriero_privadas, "
            "reparacion_filtraciones, complementarios_telefonicos, referencias_coordinacion."
        ),
    )
    subjetividades: list[str] = Field(
        default_factory=list,
        description="Lista de subjetividades, condiciones pendientes, garantías o documentos requeridos por la aseguradora.",
    )
    tasa_prorroga: str = Field(
        default="NO ESPECIFICA",
        description="Tasa de prórroga tal cual la especifica la aseguradora. Si no la menciona, usar 'NO ESPECIFICA'.",
    )
    coaseguro: Optional[list[CoaseguroEntry]] = Field(
        default=None,
        description="Si la cotización es en coaseguro (varias compañías), listar las participantes con su porcentaje.",
    )
    validez_oferta: str = Field(
        default="NO ESPECIFICA",
        description="Plazo de validez de la oferta, ej: '15/05/2026' o '15 días calendario'",
    )

    @field_validator("nombre_compania", "id_compania", "tasa_prorroga", "validez_oferta", mode="before")
    @classmethod
    def _coerce_quote_str(cls, v: Any) -> str:
        return "NO ESPECIFICA" if v is None else str(v)

    @field_validator("deducibles", "asistencias", mode="before")
    @classmethod
    def _coerce_str_dicts(cls, v: Any) -> dict[str, Optional[str]]:
        if not isinstance(v, dict):
            return {}
        return {str(k): (None if val is None else str(val)) for k, val in v.items()}

    @field_validator("coberturas", mode="before")
    @classmethod
    def _coerce_coberturas_dict(cls, v: Any) -> dict[str, Any]:
        if not isinstance(v, dict):
            return {}
        return v

    @field_validator("subjetividades", mode="before")
    @classmethod
    def _coerce_subjetividades(cls, v: Any) -> list[str]:
        if not isinstance(v, list):
            return []
        return [str(x) for x in v if x is not None]


class ProjectMeta(BaseModel):
    """Metadata about the insured project or risk, inferred from quotes."""

    tipo_cobertura: str = Field(
        default="TODO RIESGO CONSTRUCCIÓN Y MONTAJE",
        description="Tipo de póliza, ej: 'TODO RIESGO CONSTRUCCIÓN Y MONTAJE', 'MULTIRRIESGO COPROPIEDADES', 'SEGURO DE AUTOMÓVILES', 'SEGURO DE HOGAR'",
    )
    tomador: str = Field(
        default="",
        description="Nombre o razón social del tomador y No. de Identificación/Cédula/NIT si está disponible",
    )
    asegurado: str = Field(
        default="",
        description="Nombre del asegurado / propietario del proyecto o predio",
    )
    beneficiario: str = Field(
        default="",
        description="Beneficiario de la póliza",
    )
    ubicacion: str = Field(default="", description="Ubicación del riesgo, municipio o dirección")
    valor_asegurado: Optional[float] = Field(
        default=None,
        description="Valor total asegurado en pesos colombianos, sin separadores",
    )

    # TRC / Ingeniería
    vigencia_construccion: Optional[dict[str, str]] = Field(
        default=None,
        description="{'desde': 'dd/mm/aaaa', 'hasta': 'dd/mm/aaaa'}",
    )
    vigencia_mantenimiento: Optional[dict[str, str]] = Field(
        default=None,
        description="{'desde': '...', 'hasta': '...', 'tipo': 'Amplio', 'duracion': '12 meses'}",
    )
    descripcion_proyecto: str = Field(default="", description="Descripción breve del proyecto u obra a asegurar")

    # Autos
    identificacion: str = Field(default="", description="No. Identificación / Cédula / NIT del tomador")
    marca: str = Field(default="", description="Marca del vehículo (ej: MAZDA, TOYOTA, CHEVROLET)")
    placa: str = Field(default="", description="Placa del vehículo (ej: KHY654)")
    linea: str = Field(default="", description="Línea o referencia del vehículo (ej: CX-30, COROLLA)")
    modelo: str = Field(default="", description="Modelo / Año del vehículo (ej: 2024, 2025)")
    servicio: str = Field(default="Particular", description="Servicio del vehículo (Particular / Público / Especial)")
    zona_circulacion: str = Field(default="", description="Zona o ciudad de circulación")
    accesorios: str = Field(default="NO ESPECIFICA", description="Valor o amparo de accesorios")

    # Hogar
    cedula: str = Field(default="", description="Cédula del tomador")
    direccion: str = Field(default="", description="Dirección exacta del inmueble")
    ciudad: str = Field(default="", description="Ciudad del inmueble")
    ano_construccion: str = Field(default="NO ESPECIFICA", description="Año o antigüedad de construcción")
    valor_edificio: Optional[float] = Field(default=None, description="Valor asegurado Edificio")
    valor_muebles: Optional[float] = Field(default=None, description="Valor asegurado Muebles y Enseres")
    valor_equipos: Optional[float] = Field(default=None, description="Valor asegurado Equipos Electrónicos")
    valor_arte: Optional[float] = Field(default=None, description="Valor asegurado Obras de Arte")
    valor_dinero: Optional[float] = Field(default=None, description="Valor asegurado Dinero")
    asegurado_actualmente: str = Field(default="NO ESPECIFICA", description="¿Asegurado actualmente?")
    siniestros_previos: str = Field(default="NO ESPECIFICA", description="¿Siniestros previos?")

    # Copropiedades
    valor_cimentacion: Optional[Union[float, str]] = Field(
        default=None,
        description="Valor cimentación o '(incluida en ítem anterior)' si está incluida en Edificio",
    )
    valor_maquinaria: Optional[float] = Field(default=None, description="Valor maquinaria y equipo")
    valor_equipos_moviles: Optional[float] = Field(default=None, description="Valor equipos móviles y portátiles")
    valor_rce: Optional[float] = Field(default=None, description="Límite RCE")
    valor_dno: Optional[float] = Field(default=None, description="Límite D&O")
    valor_manejo: Optional[Union[float, str]] = Field(
        default=None,
        description="Límite Manejo numérico, o 'Según cada cotización' si varía entre aseguradoras",
    )
    numero_torres: str = Field(default="", description="Número de torres o bloques de la copropiedad (vacío si no especifica)")
    numero_apartamentos: str = Field(default="", description="Número de apartamentos o unidades privadas (vacío si no especifica)")
    numero_pisos: str = Field(default="", description="Número de pisos / sótanos (vacío si no especifica)")
    uso: str = Field(default="", description="Uso o giro de la copropiedad: Residencial, Comercial, Oficinas, Mixto (vacío si no especifica)")
    tipo_construccion: str = Field(default="", description="Tipo estructural de construcción (vacío si no especifica)")
    inf_copropiedad: str = Field(default="", description="Viñetas resumen con datos físicos verificados de la copropiedad")

    @field_validator(
        "tipo_cobertura",
        "tomador",
        "asegurado",
        "beneficiario",
        "ubicacion",
        "descripcion_proyecto",
        "identificacion",
        "marca",
        "placa",
        "linea",
        "modelo",
        "servicio",
        "zona_circulacion",
        "accesorios",
        "cedula",
        "direccion",
        "ciudad",
        "ano_construccion",
        "asegurado_actualmente",
        "siniestros_previos",
        "numero_torres",
        "numero_apartamentos",
        "numero_pisos",
        "uso",
        "tipo_construccion",
        "inf_copropiedad",
        mode="before",
    )
    @classmethod
    def _coerce_meta_strings(cls, v: Any) -> str:
        if v is None:
            return ""
        if isinstance(v, list):
            return "\n".join(str(x) for x in v if x is not None)
        return str(v).strip()

    @field_validator(
        "valor_asegurado",
        "valor_edificio",
        "valor_muebles",
        "valor_equipos",
        "valor_arte",
        "valor_dinero",
        "valor_maquinaria",
        "valor_equipos_moviles",
        "valor_rce",
        "valor_dno",
        mode="before",
    )
    @classmethod
    def _coerce_meta_floats(cls, v: Any) -> Optional[float]:
        return _parse_optional_float(v)

    @field_validator("valor_cimentacion", "valor_manejo", mode="before")
    @classmethod
    def _coerce_meta_float_or_str(cls, v: Any) -> Optional[Union[float, str]]:
        if v is None:
            return None
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return float(v)
        if isinstance(v, str):
            num = _parse_optional_float(v)
            if num is not None:
                return num
            s = v.strip()
            return s if s else None
        return None

    @field_validator("vigencia_construccion", "vigencia_mantenimiento", mode="before")
    @classmethod
    def _coerce_vigencia_dict(cls, v: Any) -> Optional[dict[str, str]]:
        if not isinstance(v, dict):
            return None
        return {str(k): ("" if val is None else str(val)) for k, val in v.items()}
