from __future__ import annotations

import argparse
import csv
import json
from collections import defaultdict
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from schema import (
    BRAND_TO_QA_SHEET,
    CANONICAL_FIELD_ORDER,
    EXCLUDED_PRODUCT_BRANDS,
    MEDIA_TYPE_SLUGS,
    RAW_SHEET_NAME,
    RAW_TO_CANONICAL_COLUMNS,
    SPANISH_MONTHS,
)


ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_INPUT = ROOT_DIR / "input" / "raw" / "BASE - COMPETENCIA - ENERO Y FEBRERO.xlsx"
PROCESSED_DETAIL_OUTPUT = ROOT_DIR / "input" / "processed" / "latest_base_bruta.csv"
MASTER_CSV_OUTPUT = ROOT_DIR / "output" / "master" / "master_investment_detail.csv"
MASTER_JSON_OUTPUT = ROOT_DIR / "output" / "master" / "master_investment_detail.json"
PRODUCT_OUTPUT_DIR = ROOT_DIR / "output" / "data_products" / "inversion_mensual_por_casino_ilegal"
VISUALIZATION_OUTPUT_DIR = ROOT_DIR / "output" / "visualizations"
SITE_OUTPUT_DIR = ROOT_DIR / "output" / "site"
VALIDATION_OUTPUT = ROOT_DIR / "output" / "master" / "validation_report.json"
QA_OUTPUT = ROOT_DIR / "output" / "master" / "qa_report.json"
VISUALIZATION_HTML_OUTPUT = VISUALIZATION_OUTPUT_DIR / "inversion_mensual_por_casino_ilegal.html"
VISUALIZATION_DATA_OUTPUT = VISUALIZATION_OUTPUT_DIR / "inversion_mensual_por_casino_ilegal_summary.json"
STACKED_SVG_OUTPUT = VISUALIZATION_OUTPUT_DIR / "inversion_por_marca_stackeada.svg"
LINES_SVG_OUTPUT = VISUALIZATION_OUTPUT_DIR / "inversion_por_mes_lineas.svg"
SITE_INDEX_OUTPUT = SITE_OUTPUT_DIR / "index.html"
SITE_SUMMARY_OUTPUT = SITE_OUTPUT_DIR / "data" / "inversion_mensual_por_casino_ilegal_summary.json"
SITE_MASTER_OUTPUT = SITE_OUTPUT_DIR / "data" / "master_investment_detail.json"

EXCEL_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
EXCEL_EPOCH = datetime(1899, 12, 30)
QA_TOLERANCE = 0.01


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build monthly illegal casino investment tables from the raw workbook."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=DEFAULT_INPUT,
        help="Path to the raw input workbook (.xlsx).",
    )
    return parser.parse_args()


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split())


def parse_number(value: str) -> str:
    raw = normalize_text(value)
    if not raw:
        return "0"
    number = float(raw)
    if number.is_integer():
        return str(int(number))
    return f"{number:.6f}".rstrip("0").rstrip(".")


def parse_optional_number(value: str) -> str:
    raw = normalize_text(value)
    if not raw:
        return ""
    try:
        return parse_number(raw)
    except ValueError:
        return raw


def excel_serial_to_date(value: str) -> str:
    raw = normalize_text(value)
    if not raw:
        return ""
    serial = float(raw)
    return (EXCEL_EPOCH + timedelta(days=serial)).date().isoformat()


def parse_shared_strings(workbook: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []

    root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
    strings: list[str] = []

    for item in root.findall("main:si", EXCEL_NS):
        texts = [node.text or "" for node in item.iterfind(".//main:t", EXCEL_NS)]
        strings.append("".join(texts))

    return strings


def resolve_sheet_target(workbook: ZipFile, sheet_name: str) -> str:
    workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
    rel_root = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rel_root}

    for sheet in workbook_root.find("main:sheets", EXCEL_NS):
        if sheet.attrib["name"] == sheet_name:
            relationship_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            return f"xl/{rel_map[relationship_id]}"

    raise ValueError(f"Worksheet not found: {sheet_name}")


def parse_worksheet_rows(workbook_path: Path, sheet_name: str) -> list[dict[str, str]]:
    with ZipFile(workbook_path) as workbook:
        shared_strings = parse_shared_strings(workbook)
        worksheet_path = resolve_sheet_target(workbook, sheet_name)
        worksheet_root = ET.fromstring(workbook.read(worksheet_path))

    rows: list[dict[str, str]] = []
    sheet_data = worksheet_root.find("main:sheetData", EXCEL_NS)
    if sheet_data is None:
        return rows

    for row in sheet_data:
        parsed_row: dict[str, str] = {}
        for cell in row.findall("main:c", EXCEL_NS):
            cell_ref = cell.attrib.get("r", "")
            column = "".join(character for character in cell_ref if character.isalpha())
            cell_type = cell.attrib.get("t")

            value = ""
            value_node = cell.find("main:v", EXCEL_NS)
            inline_node = cell.find("main:is", EXCEL_NS)

            if value_node is not None:
                raw_value = value_node.text or ""
                value = shared_strings[int(raw_value)] if cell_type == "s" else raw_value
            elif inline_node is not None:
                parts = [node.text or "" for node in inline_node.iterfind(".//main:t", EXCEL_NS)]
                value = "".join(parts)

            parsed_row[column] = value
        rows.append(parsed_row)

    return rows


def normalize_workbook_record(raw_record: dict[str, str], headers_by_column: dict[str, str]) -> dict[str, str]:
    normalized = {field: "" for field in CANONICAL_FIELD_ORDER}

    for column, raw_value in raw_record.items():
        header = headers_by_column.get(column, "")
        canonical_field = RAW_TO_CANONICAL_COLUMNS.get(header)
        if not canonical_field:
            continue
        normalized[canonical_field] = normalize_text(raw_value)

    normalized["brand_name"] = normalized["brand_name"].upper()
    normalized["media_type"] = normalized["media_type"].upper()
    normalized["month_name"] = normalized["month_name"].upper()
    normalized["observed_at"] = excel_serial_to_date(normalized["observed_at"])
    normalized["gross_investment"] = parse_number(normalized["gross_investment"])
    normalized["net_investment"] = parse_number(normalized["net_investment"])
    normalized["duration_seconds"] = parse_optional_number(normalized["duration_seconds"])
    normalized["tv_duration_seconds"] = parse_optional_number(normalized["tv_duration_seconds"])

    year = normalized["year"]
    month_number = SPANISH_MONTHS.get(normalized["month_name"])
    normalized["month"] = f"{year}-{month_number:02d}" if year and month_number else ""

    return normalized


def load_records(input_path: Path) -> list[dict[str, str]]:
    if not input_path.exists():
        raise FileNotFoundError(
            f"Input file not found: {input_path}. Put the source workbook under input/raw/."
        )

    if input_path.suffix.lower() != ".xlsx":
        raise ValueError(f"Unsupported input format: {input_path.suffix}. Expected .xlsx")

    rows = parse_worksheet_rows(input_path, RAW_SHEET_NAME)
    if not rows:
        raise ValueError(f"Worksheet {RAW_SHEET_NAME} is empty.")

    headers_by_column = rows[0]
    return [normalize_workbook_record(row, headers_by_column) for row in rows[1:]]


def validate_records(records: list[dict[str, str]]) -> list[str]:
    errors: list[str] = []
    required_fields = ("year", "month_name", "month", "observed_at", "media_type", "brand_name", "net_investment")

    for row_number, record in enumerate(records, start=2):
        for field in required_fields:
            if not record.get(field):
                errors.append(f"row {row_number}: missing required field '{field}'")

        if record.get("month_name") and record["month_name"] not in SPANISH_MONTHS:
            errors.append(f"row {row_number}: invalid month_name '{record['month_name']}'")

        if record.get("media_type") and record["media_type"] not in MEDIA_TYPE_SLUGS:
            errors.append(f"row {row_number}: unsupported media_type '{record['media_type']}'")

    return errors


def sort_months(months: set[str]) -> list[str]:
    return sorted(months, key=lambda value: datetime.strptime(value, "%Y-%m"))


def format_amount(value: float) -> str:
    return f"{value:.2f}"


def aggregate_monthly_tables(records: list[dict[str, str]]) -> tuple[list[str], list[str], dict[str, dict[str, dict[str, float]]]]:
    product_records = [
        record for record in records if record["brand_name"] and record["brand_name"] not in EXCLUDED_PRODUCT_BRANDS
    ]

    months = sort_months({record["month"] for record in product_records})
    brands = sorted({record["brand_name"] for record in product_records})

    aggregations: dict[str, dict[str, dict[str, float]]] = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

    for record in product_records:
        brand = record["brand_name"]
        month = record["month"]
        media_type = record["media_type"]
        net_investment = float(record["net_investment"])

        aggregations["total"][brand][month] += net_investment
        aggregations[MEDIA_TYPE_SLUGS[media_type]][brand][month] += net_investment

    return months, brands, aggregations


def build_summary_rows(months: list[str], brands: list[str], values_by_brand: dict[str, dict[str, float]]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []

    for brand in brands:
        month_values = values_by_brand.get(brand, {})
        total = 0.0
        row = {"brand_name": brand}
        for month in months:
            amount = month_values.get(month, 0.0)
            row[month] = format_amount(amount)
            total += amount
        row["total"] = format_amount(total)
        rows.append(row)

    return rows


def normalize_sheet_label(value: str) -> str:
    return normalize_text(value).upper()


def parse_sheet_float(value: str) -> float:
    raw = normalize_text(value)
    if not raw:
        return 0.0
    return float(raw)


def month_label_to_iso(year: str, month_label: str) -> str:
    month_number = SPANISH_MONTHS.get(normalize_sheet_label(month_label))
    if not month_number:
        raise ValueError(f"Unsupported month label in QA sheet: {month_label}")
    return f"{year}-{month_number:02d}"


def load_resumen_expectations(input_path: Path, months: list[str]) -> dict[str, dict[str, float]]:
    rows = parse_worksheet_rows(input_path, "RESUMEN")
    year = months[0][:4]
    months = {
        "C": month_label_to_iso(year, rows[1].get("C", "")),
        "D": month_label_to_iso(year, rows[1].get("D", "")),
    }
    expectations: dict[str, dict[str, float]] = {}

    for row in rows[2:]:
        brand = normalize_sheet_label(row.get("B", ""))
        if not brand:
            continue
        if brand == "TOTAL GENERAL":
            break
        if brand in EXCLUDED_PRODUCT_BRANDS:
            continue
        expectations[brand] = {
            months["C"]: parse_sheet_float(row.get("C", "")),
            months["D"]: parse_sheet_float(row.get("D", "")),
        }

    return expectations


def load_brand_media_expectations(input_path: Path, brand: str, months: list[str]) -> dict[str, dict[str, float]]:
    rows = parse_worksheet_rows(input_path, BRAND_TO_QA_SHEET.get(brand, brand))
    expectations: dict[str, dict[str, float]] = {}
    month_columns = {"C": months[0], "D": months[1]}

    for row in rows[5:10]:
        media_label = normalize_sheet_label(row.get("B", ""))
        if media_label == brand or media_label == "TOTAL GENERAL" or media_label not in MEDIA_TYPE_SLUGS:
            continue
        media_slug = MEDIA_TYPE_SLUGS[media_label]
        expectations[media_slug] = {
            month_columns["C"]: parse_sheet_float(row.get("C", "")),
            month_columns["D"]: parse_sheet_float(row.get("D", "")),
        }

    return expectations


def run_qa(
    input_path: Path,
    months: list[str],
    brands: list[str],
    aggregations: dict[str, dict[str, dict[str, float]]],
) -> dict[str, Any]:
    mismatches: list[dict[str, Any]] = []
    checks: list[dict[str, Any]] = []

    resumen_expectations = load_resumen_expectations(input_path, months)
    for brand in brands:
        for month in months:
            expected = resumen_expectations.get(brand, {}).get(month, 0.0)
            actual = aggregations["total"].get(brand, {}).get(month, 0.0)
            difference = round(actual - expected, 6)
            check = {
                "scope": "total",
                "brand_name": brand,
                "month": month,
                "expected": expected,
                "actual": actual,
                "difference": difference,
            }
            checks.append(check)
            if abs(difference) > QA_TOLERANCE:
                mismatches.append(check)

    for brand in brands:
        media_expectations = load_brand_media_expectations(input_path, brand, months)
        for media_slug, values in media_expectations.items():
            for month in months:
                expected = values.get(month, 0.0)
                actual = aggregations.get(media_slug, {}).get(brand, {}).get(month, 0.0)
                difference = round(actual - expected, 6)
                check = {
                    "scope": media_slug,
                    "brand_name": brand,
                    "month": month,
                    "expected": expected,
                    "actual": actual,
                    "difference": difference,
                }
                checks.append(check)
                if abs(difference) > QA_TOLERANCE:
                    mismatches.append(check)

    return {
        "passed": not mismatches,
        "tolerance": QA_TOLERANCE,
        "checks_run": len(checks),
        "mismatch_count": len(mismatches),
        "mismatches": mismatches,
    }


def build_visualization_payload(
    input_path: Path,
    records: list[dict[str, str]],
    months: list[str],
    brands: list[str],
    aggregations: dict[str, dict[str, dict[str, float]]],
    qa_report: dict[str, Any],
) -> dict[str, Any]:
    media_order = [slug for slug in ("tv_abierta", "tv_cable", "radio", "via_publica", "digital", "prensa") if slug in aggregations]
    brand_totals: list[dict[str, Any]] = []

    for brand in brands:
        media_breakdown = {}
        total = 0.0
        for media_slug in media_order:
            amount = sum(aggregations.get(media_slug, {}).get(brand, {}).get(month, 0.0) for month in months)
            media_breakdown[media_slug] = round(amount, 2)
            total += amount
        monthly_values = {month: round(aggregations["total"].get(brand, {}).get(month, 0.0), 2) for month in months}
        brand_totals.append(
            {
                "brand_name": brand,
                "total": round(total, 2),
                "monthly": monthly_values,
                "media_breakdown": media_breakdown,
            }
        )

    brand_totals.sort(key=lambda item: item["total"], reverse=True)

    sample_records = [
        {
            "brand_name": record["brand_name"],
            "observed_at": record["observed_at"],
            "media_type": record["media_type"],
            "outlet_name": record["outlet_name"],
            "program_name": record["program_name"],
            "ad_type": record["ad_type"],
            "creative_version": record["creative_version"],
            "evidence_url": record["evidence_url"],
            "net_investment": round(float(record["net_investment"]), 2),
        }
        for record in records
        if record["brand_name"] in brands and record["evidence_url"]
    ][:200]

    return {
        "title": "Inversion mensual por casino de apuesta ilegal",
        "currency": "CLP",
        "source_file": str(input_path.relative_to(ROOT_DIR)) if input_path.is_relative_to(ROOT_DIR) else str(input_path),
        "source_sheet": RAW_SHEET_NAME,
        "months": months,
        "brands": brands,
        "media_order": media_order,
        "brand_totals": brand_totals,
        "sample_records": sample_records,
        "qa_passed": qa_report["passed"],
        "qa_checks_run": qa_report["checks_run"],
        "generated_at": datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
    }


def svg_currency(value: float) -> str:
    return "$" + f"{round(value):,}".replace(",", ".")


def svg_compact(value: float) -> str:
    thresholds = (
        (1_000_000_000, "MM"),
        (1_000_000, "M"),
        (1_000, "mil"),
    )
    absolute = abs(value)
    for threshold, suffix in thresholds:
        if absolute >= threshold:
            scaled = value / threshold
            text = f"{scaled:.1f}".rstrip("0").rstrip(".")
            return f"${text} {suffix}"
    return svg_currency(value)


def svg_escape(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def build_stacked_bars_svg(payload: dict[str, Any]) -> str:
    media_colors = {
        "tv_abierta": "#b91c1c",
        "tv_cable": "#f97316",
        "radio": "#0f766e",
        "via_publica": "#7c3aed",
        "digital": "#2563eb",
        "prensa": "#475569",
    }
    labels = {
        "tv_abierta": "TV abierta",
        "tv_cable": "TV cable",
        "radio": "Radio",
        "via_publica": "Via publica",
        "digital": "Digital",
        "prensa": "Prensa",
    }
    width = 1280
    height = 820
    margin_left = 200
    margin_right = 170
    margin_top = 110
    margin_bottom = 80
    plot_width = width - margin_left - margin_right
    row_height = 48
    max_total = max((item["total"] for item in payload["brand_totals"]), default=1.0)
    ticks = 5
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}" role="img" aria-labelledby="title desc">',
        '<title id="title">Inversion por marca, barras stackeadas</title>',
        '<desc id="desc">Barras horizontales stackeadas con inversion total por marca y desglose por medio.</desc>',
        '<rect width="100%" height="100%" fill="#f6f2e9"/>',
        '<text x="48" y="54" font-family="Helvetica Neue, Arial, sans-serif" font-size="34" font-weight="700" fill="#1f2937">Inversion total por marca</text>',
        '<text x="48" y="84" font-family="Helvetica Neue, Arial, sans-serif" font-size="18" fill="#5f6b7a">Barras stackeadas por tipo de medio. Montos en CLP.</text>',
    ]

    legend_x = 48
    legend_y = 110
    for slug in payload["media_order"]:
        parts.append(f'<rect x="{legend_x}" y="{legend_y - 12}" width="14" height="14" rx="7" fill="{media_colors[slug]}"/>')
        parts.append(
            f'<text x="{legend_x + 24}" y="{legend_y}" font-family="Helvetica Neue, Arial, sans-serif" font-size="16" fill="#334155">{labels[slug]}</text>'
        )
        legend_x += 130

    for tick_index in range(ticks + 1):
        x = margin_left + plot_width * tick_index / ticks
        value = max_total * tick_index / ticks
        parts.append(f'<line x1="{x}" y1="{margin_top}" x2="{x}" y2="{height - margin_bottom}" stroke="#e8e3d8" stroke-width="1"/>')
        parts.append(
            f'<text x="{x}" y="{height - 34}" text-anchor="middle" font-family="Helvetica Neue, Arial, sans-serif" font-size="14" fill="#64748b">{svg_escape(svg_compact(value))}</text>'
        )

    for index, item in enumerate(payload["brand_totals"]):
        y = margin_top + index * row_height
        cursor = margin_left
        parts.append(
            f'<text x="{margin_left - 14}" y="{y + 22}" text-anchor="end" font-family="Helvetica Neue, Arial, sans-serif" font-size="16" fill="#1f2937">{svg_escape(item["brand_name"])}</text>'
        )
        for slug in payload["media_order"]:
            value = item["media_breakdown"].get(slug, 0.0)
            segment_width = 0 if max_total == 0 else plot_width * value / max_total
            if segment_width > 0:
                parts.append(
                    f'<rect x="{cursor}" y="{y + 6}" width="{segment_width}" height="24" rx="4" fill="{media_colors[slug]}"/>'
                )
                cursor += segment_width
        parts.append(
            f'<text x="{margin_left + plot_width + 14}" y="{y + 23}" font-family="Helvetica Neue, Arial, sans-serif" font-size="15" fill="#334155">{svg_escape(svg_currency(item["total"]))}</text>'
        )

    parts.append("</svg>")
    return "".join(parts)


def build_lines_svg(payload: dict[str, Any]) -> str:
    palette = ["#8b1e3f", "#0b6e4f", "#2563eb", "#f97316", "#7c3aed", "#b91c1c", "#0f766e", "#475569", "#d97706", "#4f46e5"]
    width = 1280
    height = 760
    margin_left = 82
    margin_right = 220
    margin_top = 80
    margin_bottom = 70
    plot_width = width - margin_left - margin_right
    plot_height = height - margin_top - margin_bottom
    max_value = max(
        (item["monthly"].get(month, 0.0) for item in payload["brand_totals"] for month in payload["months"]),
        default=1.0,
    )
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}" role="img" aria-labelledby="title desc">',
        '<title id="title">Inversion mensual por marca, lineas</title>',
        '<desc id="desc">Lineas con la evolucion mensual de la inversion total por marca.</desc>',
        '<rect width="100%" height="100%" fill="#f6f2e9"/>',
        '<text x="48" y="52" font-family="Helvetica Neue, Arial, sans-serif" font-size="34" font-weight="700" fill="#1f2937">Evolucion mensual por marca</text>',
        '<text x="48" y="82" font-family="Helvetica Neue, Arial, sans-serif" font-size="18" fill="#5f6b7a">Serie mensual de inversion neta total, en CLP.</text>',
    ]

    for tick_index in range(5):
        y = margin_top + plot_height - plot_height * tick_index / 4
        value = max_value * tick_index / 4
        parts.append(f'<line x1="{margin_left}" y1="{y}" x2="{width - margin_right}" y2="{y}" stroke="#e8e3d8" stroke-width="1"/>')
        parts.append(
            f'<text x="{margin_left - 12}" y="{y + 4}" text-anchor="end" font-family="Helvetica Neue, Arial, sans-serif" font-size="14" fill="#64748b">{svg_escape(svg_compact(value))}</text>'
        )

    for month_index, month in enumerate(payload["months"]):
        x = margin_left + (plot_width / 2 if len(payload["months"]) == 1 else plot_width * month_index / (len(payload["months"]) - 1))
        parts.append(f'<line x1="{x}" y1="{margin_top}" x2="{x}" y2="{height - margin_bottom}" stroke="#efeadd" stroke-width="1"/>')
        parts.append(
            f'<text x="{x}" y="{height - 26}" text-anchor="middle" font-family="Helvetica Neue, Arial, sans-serif" font-size="15" fill="#64748b">{svg_escape(month)}</text>'
        )

    legend_y = 116
    for index, item in enumerate(payload["brand_totals"]):
        color = palette[index % len(palette)]
        points = []
        for month_index, month in enumerate(payload["months"]):
            x = margin_left + (plot_width / 2 if len(payload["months"]) == 1 else plot_width * month_index / (len(payload["months"]) - 1))
            value = item["monthly"].get(month, 0.0)
            y = margin_top + plot_height - (0 if max_value == 0 else plot_height * value / max_value)
            points.append((x, y))
        point_string = " ".join(f"{x},{y}" for x, y in points)
        parts.append(f'<polyline fill="none" stroke="{color}" stroke-width="3" points="{point_string}"/>')
        for x, y in points:
            parts.append(f'<circle cx="{x}" cy="{y}" r="4.5" fill="{color}"/>')
        last_x, last_y = points[-1]
        parts.append(
            f'<text x="{last_x + 12}" y="{last_y + 5}" font-family="Helvetica Neue, Arial, sans-serif" font-size="14" font-weight="700" fill="{color}">{svg_escape(item["brand_name"])}</text>'
        )
        parts.append(f'<rect x="{width - margin_right + 20}" y="{legend_y - 12}" width="14" height="14" rx="7" fill="{color}"/>')
        parts.append(
            f'<text x="{width - margin_right + 42}" y="{legend_y}" font-family="Helvetica Neue, Arial, sans-serif" font-size="15" fill="#334155">{svg_escape(item["brand_name"])}</text>'
        )
        legend_y += 28

    parts.append("</svg>")
    return "".join(parts)


def build_visualization_html(payload: dict[str, Any]) -> str:
    payload_json = json.dumps(payload, ensure_ascii=True)
    return """<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Visualizacion: Inversion mensual por casino de apuesta ilegal</title>
  <style>
    :root {
      --bg: #f4f1e8;
      --panel: #fffdf8;
      --ink: #1f2937;
      --muted: #5f6b7a;
      --accent: #8b1e3f;
      --accent-2: #0b6e4f;
      --border: #d7d2c7;
      --grid: #e8e3d8;
      --shadow: 0 18px 40px rgba(31, 41, 55, 0.08);
      --tv_abierta: #b91c1c;
      --tv_cable: #f97316;
      --radio: #0f766e;
      --via_publica: #7c3aed;
      --digital: #2563eb;
      --prensa: #475569;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Helvetica Neue", Arial, sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at top left, rgba(139, 30, 63, 0.12), transparent 28%),
        radial-gradient(circle at top right, rgba(11, 110, 79, 0.10), transparent 24%),
        linear-gradient(180deg, #f6f2e9 0%, #f3efe5 100%);
    }
    a { color: var(--accent); }
    .page {
      max-width: 1280px;
      margin: 0 auto;
      padding: 32px 20px 64px;
    }
    .hero {
      display: grid;
      grid-template-columns: 2fr 1fr;
      gap: 20px;
      margin-bottom: 24px;
    }
    .panel {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 18px;
      box-shadow: var(--shadow);
      padding: 22px;
    }
    h1, h2, h3 { margin: 0 0 10px; }
    h1 { font-size: clamp(2rem, 4vw, 3.4rem); line-height: 0.95; }
    h2 { font-size: 1.35rem; }
    h3 { font-size: 1rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.08em; }
    p { margin: 0 0 10px; line-height: 1.5; }
    .lede { font-size: 1.08rem; max-width: 58ch; }
    .meta dt { font-size: 0.78rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.08em; }
    .meta dd { margin: 4px 0 14px; font-weight: bold; }
    .stats {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 14px;
      margin: 14px 0 26px;
    }
    .stat {
      background: rgba(246, 242, 233, 0.9);
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 14px;
    }
    .stat .label {
      display: block;
      font-size: 0.78rem;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.08em;
      margin-bottom: 6px;
    }
    .stat strong { font-size: 1.4rem; }
    .legend {
      display: flex;
      flex-wrap: wrap;
      gap: 10px 16px;
      margin-top: 12px;
    }
    .legend span {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      font-size: 0.92rem;
    }
    .legend i {
      width: 12px;
      height: 12px;
      border-radius: 999px;
      display: inline-block;
    }
    .charts {
      display: grid;
      grid-template-columns: 1.15fr 1fr;
      gap: 20px;
      margin-bottom: 24px;
    }
    .chart-wrap {
      overflow-x: auto;
      padding-bottom: 8px;
    }
    svg {
      width: 100%;
      min-width: 480px;
      height: auto;
      display: block;
    }
    .axis-label { fill: var(--muted); font-size: 12px; }
    .axis-line, .grid-line { stroke: var(--grid); stroke-width: 1; }
    .line-label { font-size: 11px; font-weight: bold; }
    .table-wrap {
      overflow: auto;
      border: 1px solid var(--border);
      border-radius: 14px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.95rem;
      background: white;
    }
    th, td {
      padding: 10px 12px;
      border-bottom: 1px solid var(--grid);
      text-align: left;
      vertical-align: top;
    }
    th {
      font-size: 0.78rem;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.06em;
      position: sticky;
      top: 0;
      background: #faf7f0;
    }
    .viewer {
      display: grid;
      grid-template-columns: 300px 1fr;
      gap: 20px;
      margin-top: 24px;
    }
    .controls {
      display: grid;
      gap: 12px;
      align-content: start;
    }
    .control {
      display: grid;
      gap: 6px;
    }
    .control label {
      font-size: 0.8rem;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.08em;
    }
    .control select, .control input {
      width: 100%;
      padding: 10px 12px;
      border: 1px solid var(--border);
      border-radius: 10px;
      background: white;
      font: inherit;
    }
    .note {
      color: var(--muted);
      font-size: 0.92rem;
    }
    .pill {
      display: inline-block;
      padding: 4px 10px;
      border-radius: 999px;
      background: rgba(11, 110, 79, 0.1);
      color: var(--accent-2);
      font-size: 0.85rem;
      font-weight: bold;
    }
    @media (max-width: 960px) {
      .hero, .charts, .viewer, .stats { grid-template-columns: 1fr; }
      .page { padding: 18px 14px 40px; }
    }
  </style>
</head>
<body>
  <div class="page">
    <section class="hero">
      <div class="panel">
        <h3>Visualizacion simple</h3>
        <h1>Inversion mensual por marca en una sola pagina</h1>
        <p class="lede">Esta pagina esta pensada para abrirse directamente en un navegador comun. Incluye un grafico de barras stackeadas por marca, un grafico de lineas por mes y un explorador opcional de piezas si se carga el JSON maestro.</p>
        <div class="stats" id="stats"></div>
        <div class="legend" id="legend"></div>
      </div>
      <aside class="panel">
        <h3>Contexto</h3>
        <dl class="meta">
          <dt>Moneda</dt>
          <dd id="metaCurrency"></dd>
          <dt>Fuente</dt>
          <dd id="metaSource"></dd>
          <dt>Hoja</dt>
          <dd id="metaSheet"></dd>
          <dt>QA</dt>
          <dd><span class="pill" id="metaQa"></span></dd>
        </dl>
      </aside>
    </section>

    <section class="charts">
      <article class="panel">
        <h2>Barras stackeadas por marca</h2>
        <p class="note">Cada barra suma la inversion total por marca y la divide por tipo de medio.</p>
        <div class="chart-wrap"><svg id="stackedBars" viewBox="0 0 960 560" aria-label="Grafico de barras stackeadas"></svg></div>
      </article>
      <article class="panel">
        <h2>Lineas por mes</h2>
        <p class="note">Evolucion mensual de la inversion total por marca segun los meses disponibles en el workbook.</p>
        <div class="chart-wrap"><svg id="lineChart" viewBox="0 0 760 560" aria-label="Grafico de lineas"></svg></div>
      </article>
    </section>

    <section class="panel">
      <h2>Tabla resumen</h2>
      <p class="note">Totales acumulados por marca en CLP.</p>
      <div class="table-wrap">
        <table id="summaryTable"></table>
      </div>
    </section>

    <section class="viewer">
      <aside class="panel controls">
        <div>
          <h2>Explorador de piezas</h2>
          <p class="note">La pagina intentara cargar automaticamente el JSON maestro si esta publicada como sitio. Si la abres localmente, tambien puedes cargar el archivo manualmente.</p>
        </div>
        <div class="control">
          <label for="jsonLoader">Cargar JSON maestro</label>
          <input id="jsonLoader" type="file" accept=".json,application/json">
        </div>
        <div class="control">
          <label for="brandFilter">Marca</label>
          <select id="brandFilter"><option value="">Todas</option></select>
        </div>
        <div class="control">
          <label for="mediaFilter">Tipo de medio</label>
          <select id="mediaFilter"><option value="">Todos</option></select>
        </div>
        <div class="control">
          <label for="searchFilter">Buscar texto</label>
          <input id="searchFilter" type="search" placeholder="medio, programa, version">
        </div>
        <p class="note" id="piecesStatus">Intentando cargar el JSON maestro. Si no esta disponible, se mostrara una muestra embebida.</p>
      </aside>
      <section class="panel">
        <div class="table-wrap">
          <table id="piecesTable"></table>
        </div>
      </section>
    </section>
  </div>

  <script id="payload" type="application/json">__PAYLOAD__</script>
  <script>
    const payload = JSON.parse(document.getElementById("payload").textContent);
    const numberFormatter = new Intl.NumberFormat("es-CL", { maximumFractionDigits: 0 });
    const compactFormatter = new Intl.NumberFormat("es-CL", { notation: "compact", maximumFractionDigits: 1 });
    const monthFormatter = new Intl.DateTimeFormat("es-CL", { month: "short", year: "numeric" });
    let pieceRecords = payload.sample_records.slice();
    let usingEmbeddedSamples = true;

    function formatMoney(value) {
      return "$" + numberFormatter.format(Math.round(value));
    }

    function formatCompact(value) {
      return "$" + compactFormatter.format(value);
    }

    function prettyMonth(value) {
      const [year, month] = value.split("-").map(Number);
      return monthFormatter.format(new Date(year, month - 1, 1));
    }

    function mediaLabel(slug) {
      const labels = {
        tv_abierta: "TV abierta",
        tv_cable: "TV cable",
        radio: "Radio",
        via_publica: "Via publica",
        digital: "Digital",
        prensa: "Prensa"
      };
      return labels[slug] || slug;
    }

    function colorFor(slug) {
      return getComputedStyle(document.documentElement).getPropertyValue("--" + slug).trim() || "#64748b";
    }

    function setMeta() {
      document.getElementById("metaCurrency").textContent = payload.currency;
      document.getElementById("metaSource").textContent = payload.source_file;
      document.getElementById("metaSheet").textContent = payload.source_sheet;
      document.getElementById("metaQa").textContent = payload.qa_passed ? "QA OK (" + payload.qa_checks_run + " chequeos)" : "QA con observaciones";

      const totalInvestment = payload.brand_totals.reduce((sum, item) => sum + item.total, 0);
      const topBrand = payload.brand_totals[0];
      const stats = [
        { label: "Marcas", value: payload.brands.length },
        { label: "Meses", value: payload.months.length },
        { label: "Inversion total", value: formatCompact(totalInvestment) },
        { label: "Marca lider", value: topBrand.brand_name + " · " + formatCompact(topBrand.total) }
      ];
      document.getElementById("stats").innerHTML = stats.map((item) =>
        '<div class="stat"><span class="label">' + item.label + '</span><strong>' + item.value + '</strong></div>'
      ).join("");

      document.getElementById("legend").innerHTML = payload.media_order.map((slug) =>
        '<span><i style="background:' + colorFor(slug) + '"></i>' + mediaLabel(slug) + '</span>'
      ).join("");
    }

    function renderStackedBars() {
      const svg = document.getElementById("stackedBars");
      const width = 960;
      const height = Math.max(520, 90 + payload.brand_totals.length * 42);
      svg.setAttribute("viewBox", "0 0 " + width + " " + height);
      const margin = { top: 24, right: 28, bottom: 36, left: 170 };
      const plotWidth = width - margin.left - margin.right;
      const rowHeight = 34;
      const maxValue = Math.max(...payload.brand_totals.map((item) => item.total), 1);
      const ticks = 5;
      let content = "";

      for (let i = 0; i <= ticks; i += 1) {
        const value = maxValue * i / ticks;
        const x = margin.left + (plotWidth * i / ticks);
        content += '<line class="grid-line" x1="' + x + '" y1="' + margin.top + '" x2="' + x + '" y2="' + (height - margin.bottom) + '"></line>';
        content += '<text class="axis-label" x="' + x + '" y="' + (height - 12) + '" text-anchor="middle">' + formatCompact(value) + '</text>';
      }

      payload.brand_totals.forEach((item, index) => {
        const y = margin.top + index * rowHeight + 4;
        let cursor = margin.left;
        content += '<text class="axis-label" x="' + (margin.left - 12) + '" y="' + (y + 16) + '" text-anchor="end">' + item.brand_name + '</text>';
        payload.media_order.forEach((slug) => {
          const value = item.media_breakdown[slug] || 0;
          const segmentWidth = plotWidth * (value / maxValue);
          if (segmentWidth > 0) {
            content += '<rect x="' + cursor + '" y="' + y + '" width="' + segmentWidth + '" height="22" rx="4" fill="' + colorFor(slug) + '"></rect>';
            cursor += segmentWidth;
          }
        });
        content += '<text class="axis-label" x="' + (margin.left + plotWidth + 8) + '" y="' + (y + 16) + '">' + formatMoney(item.total) + '</text>';
      });

      svg.innerHTML = content;
    }

    function renderLineChart() {
      const svg = document.getElementById("lineChart");
      const width = 760;
      const height = 560;
      svg.setAttribute("viewBox", "0 0 " + width + " " + height);
      const margin = { top: 30, right: 120, bottom: 48, left: 64 };
      const plotWidth = width - margin.left - margin.right;
      const plotHeight = height - margin.top - margin.bottom;
      const maxValue = Math.max(...payload.brand_totals.flatMap((item) => payload.months.map((month) => item.monthly[month] || 0)), 1);
      let content = "";

      for (let i = 0; i <= 4; i += 1) {
        const y = margin.top + plotHeight - (plotHeight * i / 4);
        const value = maxValue * i / 4;
        content += '<line class="grid-line" x1="' + margin.left + '" y1="' + y + '" x2="' + (width - margin.right) + '" y2="' + y + '"></line>';
        content += '<text class="axis-label" x="' + (margin.left - 10) + '" y="' + (y + 4) + '" text-anchor="end">' + formatCompact(value) + '</text>';
      }

      payload.months.forEach((month, index) => {
        const x = margin.left + (payload.months.length === 1 ? plotWidth / 2 : plotWidth * index / (payload.months.length - 1));
        content += '<line class="axis-line" x1="' + x + '" y1="' + margin.top + '" x2="' + x + '" y2="' + (height - margin.bottom) + '"></line>';
        content += '<text class="axis-label" x="' + x + '" y="' + (height - 16) + '" text-anchor="middle">' + prettyMonth(month) + '</text>';
      });

      payload.brand_totals.forEach((item) => {
        const color = colorFor(payload.media_order[payload.brand_totals.indexOf(item) % payload.media_order.length] || "digital");
        const points = payload.months.map((month, index) => {
          const x = margin.left + (payload.months.length === 1 ? plotWidth / 2 : plotWidth * index / (payload.months.length - 1));
          const value = item.monthly[month] || 0;
          const y = margin.top + plotHeight - (plotHeight * value / maxValue);
          return { x, y, value };
        });
        content += '<polyline fill="none" stroke="' + color + '" stroke-width="2.5" points="' + points.map((p) => p.x + ',' + p.y).join(' ') + '"></polyline>';
        points.forEach((point) => {
          content += '<circle cx="' + point.x + '" cy="' + point.y + '" r="4" fill="' + color + '"></circle>';
        });
        const lastPoint = points[points.length - 1];
        content += '<text class="line-label" x="' + (lastPoint.x + 8) + '" y="' + (lastPoint.y + 4) + '" fill="' + color + '">' + item.brand_name + '</text>';
      });

      svg.innerHTML = content;
    }

    function renderSummaryTable() {
      const table = document.getElementById("summaryTable");
      const header = ['<thead><tr><th>Marca</th>', ...payload.months.map((month) => '<th>' + prettyMonth(month) + '</th>'), '<th>Total</th></tr></thead>'].join("");
      const rows = payload.brand_totals.map((item) => {
        return '<tr><td><strong>' + item.brand_name + '</strong></td>' +
          payload.months.map((month) => '<td>' + formatMoney(item.monthly[month] || 0) + '</td>').join('') +
          '<td>' + formatMoney(item.total) + '</td></tr>';
      }).join("");
      table.innerHTML = header + '<tbody>' + rows + '</tbody>';
    }

    function buildPiecesFilters(records) {
      const brandFilter = document.getElementById("brandFilter");
      const mediaFilter = document.getElementById("mediaFilter");
      const brands = Array.from(new Set(records.map((item) => item.brand_name).filter(Boolean))).sort();
      const media = Array.from(new Set(records.map((item) => item.media_type).filter(Boolean))).sort();
      brandFilter.innerHTML = '<option value="">Todas</option>' + brands.map((item) => '<option value="' + item + '">' + item + '</option>').join('');
      mediaFilter.innerHTML = '<option value="">Todos</option>' + media.map((item) => '<option value="' + item + '">' + item + '</option>').join('');
    }

    function renderPiecesTable() {
      const brandValue = document.getElementById("brandFilter").value;
      const mediaValue = document.getElementById("mediaFilter").value;
      const searchValue = document.getElementById("searchFilter").value.trim().toLowerCase();
      const rows = pieceRecords
        .filter((item) => !brandValue || item.brand_name === brandValue)
        .filter((item) => !mediaValue || item.media_type === mediaValue)
        .filter((item) => {
          if (!searchValue) return true;
          return [item.outlet_name, item.program_name, item.creative_version, item.ad_type].join(' ').toLowerCase().includes(searchValue);
        })
        .sort((left, right) => (right.net_investment || 0) - (left.net_investment || 0))
        .slice(0, 80);

      const table = document.getElementById("piecesTable");
      table.innerHTML =
        '<thead><tr><th>Fecha</th><th>Marca</th><th>Medio</th><th>Programa</th><th>Pieza</th><th>Inversion neta</th><th>Evidencia</th></tr></thead>' +
        '<tbody>' +
        rows.map((item) => '<tr>' +
          '<td>' + (item.observed_at || '') + '</td>' +
          '<td><strong>' + (item.brand_name || '') + '</strong></td>' +
          '<td>' + [item.media_type, item.outlet_name].filter(Boolean).join(' · ') + '</td>' +
          '<td>' + (item.program_name || '') + '</td>' +
          '<td>' + [item.ad_type, item.creative_version].filter(Boolean).join(' / ') + '</td>' +
          '<td>' + formatMoney(item.net_investment || 0) + '</td>' +
          '<td>' + (item.evidence_url ? '<a href="' + item.evidence_url + '" target="_blank" rel="noreferrer">abrir</a>' : '') + '</td>' +
          '</tr>').join('') +
        '</tbody>';

      document.getElementById("piecesStatus").textContent = rows.length + " registros visibles" +
        (usingEmbeddedSamples ? " (muestra embebida)" : " (desde JSON cargado)");
    }

    function bindJsonLoader() {
      const loader = document.getElementById("jsonLoader");
      loader.addEventListener("change", async (event) => {
        const file = event.target.files && event.target.files[0];
        if (!file) return;
        const text = await file.text();
        const parsed = JSON.parse(text);
        pieceRecords = parsed.map((item) => ({
          brand_name: item.brand_name,
          observed_at: item.observed_at,
          media_type: item.media_type,
          outlet_name: item.outlet_name,
          program_name: item.program_name,
          ad_type: item.ad_type,
          creative_version: item.creative_version,
          evidence_url: item.evidence_url,
          net_investment: Number(item.net_investment || 0)
        }));
        usingEmbeddedSamples = false;
        buildPiecesFilters(pieceRecords);
        renderPiecesTable();
      });

      ["brandFilter", "mediaFilter", "searchFilter"].forEach((id) => {
        document.getElementById(id).addEventListener("input", renderPiecesTable);
      });
    }

    async function tryAutoLoadMasterJson() {
      try {
        const response = await fetch("./data/master_investment_detail.json");
        if (!response.ok) throw new Error("master json unavailable");
        const parsed = await response.json();
        pieceRecords = parsed.map((item) => ({
          brand_name: item.brand_name,
          observed_at: item.observed_at,
          media_type: item.media_type,
          outlet_name: item.outlet_name,
          program_name: item.program_name,
          ad_type: item.ad_type,
          creative_version: item.creative_version,
          evidence_url: item.evidence_url,
          net_investment: Number(item.net_investment || 0)
        }));
        usingEmbeddedSamples = false;
        buildPiecesFilters(pieceRecords);
        renderPiecesTable();
      } catch (error) {
        renderPiecesTable();
      }
    }

    setMeta();
    renderStackedBars();
    renderLineChart();
    renderSummaryTable();
    buildPiecesFilters(pieceRecords);
    bindJsonLoader();
    tryAutoLoadMasterJson();
  </script>
</body>
</html>
""".replace("__PAYLOAD__", payload_json)


def write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=True, indent=2)


def build_validation_report(
    input_path: Path,
    records: list[dict[str, str]],
    product_brands: list[str],
    product_months: list[str],
    table_names: list[str],
    qa_passed: bool,
    errors: list[str],
) -> dict[str, Any]:
    return {
        "input_file": str(input_path.relative_to(ROOT_DIR)) if input_path.is_relative_to(ROOT_DIR) else str(input_path),
        "worksheet_name": RAW_SHEET_NAME,
        "raw_record_count": len(records),
        "product_record_count": sum(1 for record in records if record["brand_name"] not in EXCLUDED_PRODUCT_BRANDS),
        "product_brands": product_brands,
        "excluded_product_brands": sorted(EXCLUDED_PRODUCT_BRANDS),
        "months": product_months,
        "tables_generated": table_names,
        "qa_passed": qa_passed,
        "error_count": len(errors),
        "errors": errors,
        "generated_at": datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
    }


def main() -> int:
    args = parse_args()
    records = load_records(args.input)
    errors = validate_records(records)

    if errors:
        write_json(
            VALIDATION_OUTPUT,
            build_validation_report(
                input_path=args.input,
                records=records,
                product_brands=[],
                product_months=[],
                table_names=[],
                qa_passed=False,
                errors=errors,
            ),
        )
        raise SystemExit("Validation failed. See output/master/validation_report.json for details.")

    write_csv(PROCESSED_DETAIL_OUTPUT, list(CANONICAL_FIELD_ORDER), records)
    write_csv(MASTER_CSV_OUTPUT, list(CANONICAL_FIELD_ORDER), records)
    write_json(MASTER_JSON_OUTPUT, records)

    months, brands, aggregations = aggregate_monthly_tables(records)
    summary_fieldnames = ["brand_name", *months, "total"]
    table_names: list[str] = []

    for table_name, values_by_brand in sorted(aggregations.items()):
        output_path = PRODUCT_OUTPUT_DIR / f"{table_name}.csv"
        summary_rows = build_summary_rows(months, brands, values_by_brand)
        write_csv(output_path, summary_fieldnames, summary_rows)
        table_names.append(f"{table_name}.csv")

    qa_report = run_qa(args.input, months, brands, aggregations)
    write_json(QA_OUTPUT, qa_report)
    if not qa_report["passed"]:
        raise SystemExit("QA failed. See output/master/qa_report.json for details.")

    visualization_payload = build_visualization_payload(args.input, records, months, brands, aggregations, qa_report)
    write_json(VISUALIZATION_DATA_OUTPUT, visualization_payload)
    visualization_html = build_visualization_html(visualization_payload)
    write_text(VISUALIZATION_HTML_OUTPUT, visualization_html)
    write_text(STACKED_SVG_OUTPUT, build_stacked_bars_svg(visualization_payload))
    write_text(LINES_SVG_OUTPUT, build_lines_svg(visualization_payload))
    write_text(SITE_INDEX_OUTPUT, visualization_html)
    write_json(SITE_SUMMARY_OUTPUT, visualization_payload)
    write_json(SITE_MASTER_OUTPUT, records)

    write_json(
        VALIDATION_OUTPUT,
        build_validation_report(
            input_path=args.input,
            records=records,
            product_brands=brands,
            product_months=months,
            table_names=table_names,
            qa_passed=qa_report["passed"],
            errors=[],
        ),
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
