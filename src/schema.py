from __future__ import annotations

RAW_SHEET_NAME = "BASE BRUTA"

RAW_TO_CANONICAL_COLUMNS = {
    "Año": "year",
    "Mes": "month_name",
    "Dia de la semana": "weekday_name",
    "Fecha": "observed_at",
    "Tipo de medio": "media_type",
    "Categoria": "category",
    "Sector": "sector",
    "Sub-sector": "sub_sector",
    "Anunciante": "advertiser",
    "Marca": "brand_name",
    "Producto": "product_name",
    "Genero": "genre",
    "Medio": "outlet_name",
    "Tipo de aviso": "ad_type",
    "Programa": "program_name",
    "Hora": "observed_time",
    "Inversión": "gross_investment",
    "Inversión Neta": "net_investment",
    "Duracion": "duration_seconds",
    "Duracion TV": "tv_duration_seconds",
    "Version": "creative_version",
    "IND 18-69 Alto - Medio - Bajo": "target_index_18_69",
    "Multimedia": "evidence_url",
}

CANONICAL_FIELD_ORDER = (
    "year",
    "month_name",
    "month",
    "week_ending",
    "weekday_name",
    "observed_at",
    "media_type",
    "category",
    "sector",
    "sub_sector",
    "advertiser",
    "brand_name",
    "product_name",
    "genre",
    "outlet_name",
    "ad_type",
    "program_name",
    "observed_time",
    "gross_investment",
    "net_investment",
    "duration_seconds",
    "tv_duration_seconds",
    "creative_version",
    "target_index_18_69",
    "evidence_url",
)

SPANISH_MONTHS = {
    "ENERO": 1,
    "FEBRERO": 2,
    "MARZO": 3,
    "ABRIL": 4,
    "MAYO": 5,
    "JUNIO": 6,
    "JULIO": 7,
    "AGOSTO": 8,
    "SEPTIEMBRE": 9,
    "OCTUBRE": 10,
    "NOVIEMBRE": 11,
    "DICIEMBRE": 12,
}

MEDIA_TYPE_SLUGS = {
    "DIGITAL": "digital",
    "PRENSA": "prensa",
    "RADIO": "radio",
    "TV ABIERTA": "tv_abierta",
    "TV CABLE": "tv_cable",
    "VIA PUBLICA": "via_publica",
}

# Supuesto del primer data product: excluir marcas reguladas en Chile.
EXCLUDED_PRODUCT_BRANDS = {
    "MONTICELLO",
    "XPERTO",
}

BRAND_TO_QA_SHEET = {
    "APUESTAS ROYAL": "ROYAL",
}
