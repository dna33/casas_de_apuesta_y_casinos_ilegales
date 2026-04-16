from __future__ import annotations

import argparse
import csv
import html
import json
import re
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
PROCESSED_DETAIL_OUTPUT = ROOT_DIR / "input" / "processed" / "latest_base_bruta.csv"
MASTER_CSV_OUTPUT = ROOT_DIR / "output" / "master" / "master_investment_detail.csv"
MASTER_JSON_OUTPUT = ROOT_DIR / "output" / "master" / "master_investment_detail.json"
PRODUCT_OUTPUT_DIR = ROOT_DIR / "output" / "data_products" / "inversion_semanal_por_casino_ilegal"
CHANGES_OUTPUT_DIR = ROOT_DIR / "output" / "data_products" / "cambios_vs_corte_anterior_semanal"
VISUALIZATION_OUTPUT_DIR = ROOT_DIR / "output" / "visualizations"
SITE_OUTPUT_DIR = ROOT_DIR / "output" / "site"
VALIDATION_OUTPUT = ROOT_DIR / "output" / "master" / "validation_report.json"
QA_OUTPUT = ROOT_DIR / "output" / "master" / "qa_report.json"
VISUALIZATION_HTML_OUTPUT = VISUALIZATION_OUTPUT_DIR / "inversion_semanal_por_casino_ilegal.html"
VISUALIZATION_DATA_OUTPUT = VISUALIZATION_OUTPUT_DIR / "inversion_semanal_por_casino_ilegal_summary.json"
STACKED_SVG_OUTPUT = VISUALIZATION_OUTPUT_DIR / "inversion_por_marca_stackeada.svg"
HEATMAP_SVG_OUTPUT = VISUALIZATION_OUTPUT_DIR / "inversion_por_semana_heatmap.svg"
SITE_INDEX_OUTPUT = SITE_OUTPUT_DIR / "index.html"
SITE_SUMMARY_OUTPUT = SITE_OUTPUT_DIR / "data" / "inversion_semanal_por_casino_ilegal_summary.json"
SITE_MASTER_OUTPUT = SITE_OUTPUT_DIR / "data" / "master_investment_detail.json"
REPO_URL = "https://github.com/dna33/casas_de_apuesta_y_casinos_ilegales"
RAW_SHEET_CANDIDATES = (RAW_SHEET_NAME, "DATOS")
SUMMARY_SHEET_CANDIDATES = ("RESUMEN", "CRUCES")

EXCEL_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
EXCEL_EPOCH = datetime(1899, 12, 30)
QA_TOLERANCE = 0.01


def find_available_workbooks() -> list[Path]:
    return sorted((ROOT_DIR / "input" / "raw").glob("*.xlsx"))


def workbook_sheet_names(workbook_path: Path) -> list[str]:
    with ZipFile(workbook_path) as workbook:
        workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
    return [sheet.attrib["name"] for sheet in workbook_root.find("main:sheets", EXCEL_NS)]


def resolve_available_sheet_name(workbook_path: Path, candidates: tuple[str, ...]) -> str | None:
    available_sheets = set(workbook_sheet_names(workbook_path))
    return next((candidate for candidate in candidates if candidate in available_sheets), None)


def workbook_coverage_end(workbook_path: Path) -> str:
    raw_sheet_name = resolve_available_sheet_name(workbook_path, RAW_SHEET_CANDIDATES)
    if not raw_sheet_name:
        return ""
    rows = parse_worksheet_rows(workbook_path, raw_sheet_name)
    if len(rows) <= 1:
        return ""
    header_row = rows[0]
    date_column = next((column for column, header in header_row.items() if normalize_text(header) == "Fecha"), None)
    if not date_column:
        return ""
    max_date = ""
    for row in rows[1:]:
        raw_date = row.get(date_column, "")
        if not raw_date:
            continue
        iso_date = excel_serial_to_date(raw_date)
        if iso_date > max_date:
            max_date = iso_date
    return max_date


def default_input_workbook() -> Path:
    workbooks = find_available_workbooks()
    if not workbooks:
        return ROOT_DIR / "input" / "raw" / "latest.xlsx"
    return max(workbooks, key=lambda path: (workbook_coverage_end(path), path.name))


def default_previous_workbook(current_input: Path) -> Path | None:
    workbooks = [path for path in find_available_workbooks() if path.resolve() != current_input.resolve()]
    if not workbooks:
        return None
    return max(workbooks, key=lambda path: (workbook_coverage_end(path), path.name))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build weekly illegal casino investment tables from the raw workbook."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=default_input_workbook(),
        help="Path to the current raw input workbook (.xlsx). Defaults to the newest workbook under input/raw/.",
    )
    parser.add_argument(
        "--previous-input",
        type=Path,
        default=None,
        help="Optional path to the previous workbook used to compute changes between cuts.",
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
    if normalized["observed_at"]:
        observed_date = datetime.fromisoformat(normalized["observed_at"]).date()
        week_ending = observed_date + timedelta(days=(6 - observed_date.weekday()))
        normalized["week_ending"] = week_ending.isoformat()

    return normalized


def load_records(input_path: Path) -> list[dict[str, str]]:
    if not input_path.exists():
        raise FileNotFoundError(
            f"Input file not found: {input_path}. Put the source workbook under input/raw/."
        )

    if input_path.suffix.lower() != ".xlsx":
        raise ValueError(f"Unsupported input format: {input_path.suffix}. Expected .xlsx")

    raw_sheet_name = resolve_available_sheet_name(input_path, RAW_SHEET_CANDIDATES)
    if not raw_sheet_name:
        raise ValueError(
            f"Could not find a supported raw sheet in {input_path.name}. Expected one of: {', '.join(RAW_SHEET_CANDIDATES)}"
        )

    rows = parse_worksheet_rows(input_path, raw_sheet_name)
    if not rows:
        raise ValueError(f"Worksheet {raw_sheet_name} is empty.")

    headers_by_column = rows[0]
    return [normalize_workbook_record(row, headers_by_column) for row in rows[1:]]


def validate_records(records: list[dict[str, str]]) -> list[str]:
    errors: list[str] = []
    required_fields = ("year", "month_name", "month", "week_ending", "observed_at", "media_type", "brand_name", "net_investment")

    for row_number, record in enumerate(records, start=2):
        for field in required_fields:
            if not record.get(field):
                errors.append(f"row {row_number}: missing required field '{field}'")

        if record.get("month_name") and record["month_name"] not in SPANISH_MONTHS:
            errors.append(f"row {row_number}: invalid month_name '{record['month_name']}'")

        if record.get("media_type") and record["media_type"] not in MEDIA_TYPE_SLUGS:
            errors.append(f"row {row_number}: unsupported media_type '{record['media_type']}'")

    return errors


def sort_periods(periods: set[str]) -> list[str]:
    return sorted(periods, key=lambda value: datetime.strptime(value, "%Y-%m-%d" if len(value) == 10 else "%Y-%m"))


def published_records(records: list[dict[str, str]]) -> list[dict[str, str]]:
    return [record for record in records if record["brand_name"] not in EXCLUDED_PRODUCT_BRANDS]


def format_amount(value: float) -> str:
    return f"{value:.2f}"


def format_cut_label(input_path: Path) -> str:
    coverage_end = workbook_coverage_end(input_path)
    if coverage_end:
        observed_date = datetime.fromisoformat(coverage_end).date()
        month_name = {
            1: "enero",
            2: "febrero",
            3: "marzo",
            4: "abril",
            5: "mayo",
            6: "junio",
            7: "julio",
            8: "agosto",
            9: "septiembre",
            10: "octubre",
            11: "noviembre",
            12: "diciembre",
        }[observed_date.month]
        return f"Corte al {observed_date.day:02d} de {month_name} de {observed_date.year}"
    return "Corte disponible"


def aggregate_period_tables(
    records: list[dict[str, str]],
    period_field: str,
) -> tuple[list[str], list[str], dict[str, dict[str, dict[str, float]]]]:
    product_records = [
        record for record in records if record["brand_name"] and record["brand_name"] not in EXCLUDED_PRODUCT_BRANDS
    ]

    periods = sort_periods({record[period_field] for record in product_records})
    brands = sorted({record["brand_name"] for record in product_records})

    aggregations: dict[str, dict[str, dict[str, float]]] = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

    for record in product_records:
        brand = record["brand_name"]
        period = record[period_field]
        media_type = record["media_type"]
        net_investment = float(record["net_investment"])

        aggregations["total"][brand][period] += net_investment
        aggregations[MEDIA_TYPE_SLUGS[media_type]][brand][period] += net_investment

    return periods, brands, aggregations


def build_summary_rows(periods: list[str], brands: list[str], values_by_brand: dict[str, dict[str, float]]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []

    for brand in brands:
        period_values = values_by_brand.get(brand, {})
        total = 0.0
        row = {"brand_name": brand}
        for period in periods:
            amount = period_values.get(period, 0.0)
            row[period] = format_amount(amount)
            total += amount
        row["total"] = format_amount(total)
        rows.append(row)

    return rows


def normalize_sheet_label(value: str) -> str:
    return normalize_text(value).upper()


def excel_column_number(column_letters: str) -> int:
    value = 0
    for character in column_letters:
        value = value * 26 + (ord(character.upper()) - 64)
    return value


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


def load_resumen_expectations(input_path: Path, monthly_periods: list[str]) -> dict[str, dict[str, float]]:
    summary_sheet_name = resolve_available_sheet_name(input_path, SUMMARY_SHEET_CANDIDATES)
    if summary_sheet_name == "CRUCES":
        return load_cruces_expectations(input_path, monthly_periods)
    if not summary_sheet_name:
        return {}

    rows = parse_worksheet_rows(input_path, summary_sheet_name)
    year = monthly_periods[0][:4]
    month_columns: dict[str, str] = {}
    for column_letter, label in rows[1].items():
        normalized_label = normalize_sheet_label(label)
        if normalized_label in SPANISH_MONTHS and excel_column_number(column_letter) < excel_column_number("G"):
            month_columns[column_letter] = month_label_to_iso(year, label)
    expectations: dict[str, dict[str, float]] = {}

    for row in rows[2:]:
        brand = normalize_sheet_label(row.get("B", ""))
        if not brand:
            continue
        if brand == "TOTAL GENERAL":
            break
        if brand in EXCLUDED_PRODUCT_BRANDS:
            continue
        expectations[brand] = {period: parse_sheet_float(row.get(column_letter, "")) for column_letter, period in month_columns.items()}

    return expectations


def load_cruces_expectations(input_path: Path, monthly_periods: list[str]) -> dict[str, dict[str, float]]:
    rows = parse_worksheet_rows(input_path, "CRUCES")
    if len(rows) < 3:
        return {}

    year = monthly_periods[0][:4]
    month_columns: dict[str, str] = {}
    for column_letter, label in rows[1].items():
        normalized_label = normalize_sheet_label(label)
        if normalized_label in SPANISH_MONTHS and excel_column_number(column_letter) < excel_column_number("G"):
            month_columns[column_letter] = month_label_to_iso(year, label)

    expectations: dict[str, dict[str, float]] = {}
    for row in rows[2:]:
        brand = normalize_sheet_label(row.get("A", ""))
        if not brand:
            continue
        if brand == "TOTAL GENERAL":
            break
        if brand in EXCLUDED_PRODUCT_BRANDS:
            continue
        expectations[brand] = {
            period: parse_sheet_float(row.get(column_letter, ""))
            for column_letter, period in month_columns.items()
        }

    return expectations


def load_brand_media_expectations(input_path: Path, brand: str, monthly_periods: list[str]) -> dict[str, dict[str, float]]:
    sheet_name = BRAND_TO_QA_SHEET.get(brand, brand)
    if sheet_name not in workbook_sheet_names(input_path):
        return {}

    rows = parse_worksheet_rows(input_path, sheet_name)
    expectations: dict[str, dict[str, float]] = {}
    year = monthly_periods[0][:4]
    month_columns: dict[str, str] = {}
    for column_letter, label in rows[1].items():
        normalized_label = normalize_sheet_label(label)
        if normalized_label in SPANISH_MONTHS and excel_column_number(column_letter) < excel_column_number("F"):
            month_columns[column_letter] = month_label_to_iso(year, label)

    for row in rows[5:10]:
        media_label = normalize_sheet_label(row.get("B", ""))
        if media_label == brand or media_label == "TOTAL GENERAL" or media_label not in MEDIA_TYPE_SLUGS:
            continue
        media_slug = MEDIA_TYPE_SLUGS[media_label]
        expectations[media_slug] = {period: parse_sheet_float(row.get(column_letter, "")) for column_letter, period in month_columns.items()}

    return expectations


def run_qa(
    input_path: Path,
    monthly_periods: list[str],
    brands: list[str],
    monthly_aggregations: dict[str, dict[str, dict[str, float]]],
) -> dict[str, Any]:
    summary_sheet_name = resolve_available_sheet_name(input_path, SUMMARY_SHEET_CANDIDATES)
    mismatches: list[dict[str, Any]] = []
    checks: list[dict[str, Any]] = []

    resumen_expectations = load_resumen_expectations(input_path, monthly_periods)
    available_expectation_months = sorted(
        {
            month
            for values in resumen_expectations.values()
            for month, expected in values.items()
            if expected or month in monthly_periods
        }
    )
    if summary_sheet_name == "CRUCES" and not set(monthly_periods).issubset(set(available_expectation_months)):
        return {
            "passed": True,
            "skipped": True,
            "skip_reason": "CRUCES does not cover all months present in DATOS; skipping blocking QA for this cut.",
            "tolerance": QA_TOLERANCE,
            "checks_run": 0,
            "mismatch_count": 0,
            "mismatches": [],
        }

    for brand in brands:
        for month in monthly_periods:
            expected = resumen_expectations.get(brand, {}).get(month, 0.0)
            actual = monthly_aggregations["total"].get(brand, {}).get(month, 0.0)
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
        media_expectations = load_brand_media_expectations(input_path, brand, monthly_periods)
        for media_slug, values in media_expectations.items():
            for month in monthly_periods:
                expected = values.get(month, 0.0)
                actual = monthly_aggregations.get(media_slug, {}).get(brand, {}).get(month, 0.0)
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
    source_sheet_name: str | None,
    records: list[dict[str, str]],
    periods: list[str],
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
            amount = sum(aggregations.get(media_slug, {}).get(brand, {}).get(period, 0.0) for period in periods)
            media_breakdown[media_slug] = round(amount, 2)
            total += amount
        period_values = {period: round(aggregations["total"].get(brand, {}).get(period, 0.0), 2) for period in periods}
        brand_totals.append(
            {
                "brand_name": brand,
                "total": round(total, 2),
                "series": period_values,
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

    readme_text = (ROOT_DIR / "README.md").read_text(encoding="utf-8")

    return {
        "title": "Inversion semanal por casino de apuesta ilegal",
        "currency": "CLP",
        "repo_url": REPO_URL,
        "source_file": format_cut_label(input_path),
        "source_sheet": source_sheet_name,
        "period_granularity": "week",
        "periods": periods,
        "brands": brands,
        "media_order": media_order,
        "brand_totals": brand_totals,
        "sample_records": sample_records,
        "readme_html": markdown_to_html(readme_text),
        "qa_passed": qa_report["passed"],
        "qa_checks_run": qa_report["checks_run"],
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


def render_inline_markdown(text: str) -> str:
    escaped = html.escape(text)
    escaped = re.sub(r"`([^`]+)`", r"<code>\1</code>", escaped)
    escaped = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", escaped)
    escaped = re.sub(r"\*([^*]+)\*", r"<em>\1</em>", escaped)
    escaped = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r'<a href="\2" target="_blank" rel="noreferrer">\1</a>', escaped)
    return escaped


def markdown_to_html(markdown_text: str) -> str:
    lines = markdown_text.splitlines()
    parts: list[str] = []
    paragraph: list[str] = []
    list_items: list[str] = []
    in_code = False
    code_lines: list[str] = []

    def flush_paragraph() -> None:
        nonlocal paragraph
        if paragraph:
            parts.append(f"<p>{render_inline_markdown(' '.join(paragraph).strip())}</p>")
            paragraph = []

    def flush_list() -> None:
        nonlocal list_items
        if list_items:
            items = "".join(f"<li>{render_inline_markdown(item)}</li>" for item in list_items)
            parts.append(f"<ul>{items}</ul>")
            list_items = []

    def flush_code() -> None:
        nonlocal code_lines
        if code_lines:
            parts.append(f"<pre><code>{html.escape(chr(10).join(code_lines))}</code></pre>")
            code_lines = []

    for raw_line in lines:
        line = raw_line.rstrip()
        stripped = line.strip()

        if stripped.startswith("```"):
            flush_paragraph()
            flush_list()
            if in_code:
                flush_code()
                in_code = False
            else:
                in_code = True
            continue

        if in_code:
            code_lines.append(line)
            continue

        if not stripped:
            flush_paragraph()
            flush_list()
            continue

        if stripped == "---":
            flush_paragraph()
            flush_list()
            parts.append("<hr>")
            continue

        if stripped.startswith("#"):
            flush_paragraph()
            flush_list()
            level = min(len(stripped) - len(stripped.lstrip("#")), 6)
            content = stripped[level:].strip()
            parts.append(f"<h{level + 1}>{render_inline_markdown(content)}</h{level + 1}>")
            continue

        if re.match(r"^\d+\.\s+", stripped):
            flush_paragraph()
            flush_list()
            parts.append(f"<p>{render_inline_markdown(stripped)}</p>")
            continue

        if stripped.startswith("- "):
            flush_paragraph()
            list_items.append(stripped[2:].strip())
            continue

        if stripped.startswith("![") and "](" in stripped:
            flush_paragraph()
            flush_list()
            match = re.match(r"!\[([^\]]*)\]\(([^)]+)\)", stripped)
            if match:
                alt_text, src = match.groups()
                parts.append(
                    f'<figure><img src="{html.escape(src)}" alt="{html.escape(alt_text)}"><figcaption>{html.escape(alt_text)}</figcaption></figure>'
                )
                continue

        paragraph.append(stripped)

    flush_paragraph()
    flush_list()
    if in_code:
        flush_code()

    return "".join(parts)


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
        '<title id="title">Distribucion estimada de inversion por marca y medio</title>',
        '<desc id="desc">Barras horizontales stackeadas con la distribucion estimada de inversion por marca y desglose por medio.</desc>',
        '<rect width="100%" height="100%" fill="#f6f2e9"/>',
        '<text x="48" y="54" font-family="Helvetica Neue, Arial, sans-serif" font-size="34" font-weight="700" fill="#1f2937">Distribucion estimada de inversion por marca y medio</text>',
        '<text x="48" y="84" font-family="Helvetica Neue, Arial, sans-serif" font-size="18" fill="#5f6b7a">Composicion por tipo de medio. Montos estimados en CLP segun observacion y tarifas estandar.</text>',
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
    width = 1280
    height = 860
    margin_left = 210
    margin_right = 70
    margin_top = 110
    margin_bottom = 100
    plot_width = width - margin_left - margin_right
    plot_height = height - margin_top - margin_bottom
    max_value = max(
        (item["series"].get(period, 0.0) for item in payload["brand_totals"] for period in payload["periods"]),
        default=1.0,
    )
    rows = max(len(payload["brand_totals"]), 1)
    cols = max(len(payload["periods"]), 1)
    cell_width = plot_width / cols
    cell_height = plot_height / rows

    def heat_color(value: float) -> str:
        ratio = 0.0 if max_value == 0 else min(max(value / max_value, 0.0), 1.0) ** 0.55
        start = (244, 241, 232)
        end = (139, 30, 63)
        channels = [round(start[i] + (end[i] - start[i]) * ratio) for i in range(3)]
        return "#" + "".join(f"{channel:02x}" for channel in channels)

    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}" role="img" aria-labelledby="title desc">',
        '<title id="title">Mapa de calor semanal de la inversion estimada por marca</title>',
        '<desc id="desc">Mapa de calor con la evolucion semanal estimada de la inversion por marca.</desc>',
        '<rect width="100%" height="100%" fill="#f6f2e9"/>',
        '<text x="48" y="52" font-family="Helvetica Neue, Arial, sans-serif" font-size="34" font-weight="700" fill="#1f2937">Mapa de calor semanal de la inversion estimada por marca</text>',
        '<text x="48" y="82" font-family="Helvetica Neue, Arial, sans-serif" font-size="18" fill="#5f6b7a">Cada celda representa un corte semanal de 2026. Cuanto mas intenso el color, mayor la inversion estimada observada.</text>',
    ]

    for period_index, period in enumerate(payload["periods"]):
        x = margin_left + cell_width * period_index + cell_width / 2
        parts.append(
            f'<text x="{x}" y="{margin_top - 18}" text-anchor="middle" font-family="Helvetica Neue, Arial, sans-serif" font-size="14" fill="#64748b">{svg_escape(period)}</text>'
        )
        parts.append(f'<line x1="{margin_left + cell_width * period_index}" y1="{margin_top}" x2="{margin_left + cell_width * period_index}" y2="{margin_top + plot_height}" stroke="#efeadd" stroke-width="1"/>')

    for brand_index, item in enumerate(payload["brand_totals"]):
        y = margin_top + cell_height * brand_index
        parts.append(
            f'<text x="{margin_left - 14}" y="{y + cell_height / 2 + 5}" text-anchor="end" font-family="Helvetica Neue, Arial, sans-serif" font-size="16" fill="#1f2937">{svg_escape(item["brand_name"])}</text>'
        )
        for period_index, period in enumerate(payload["periods"]):
            value = item["series"].get(period, 0.0)
            x = margin_left + cell_width * period_index
            parts.append(
                f'<rect x="{x + 2}" y="{y + 2}" width="{cell_width - 4}" height="{cell_height - 4}" rx="6" fill="{heat_color(value)}"/>'
            )
            if value > 0:
                text_color = "#ffffff" if max_value and value / max_value > 0.45 else "#1f2937"
                parts.append(
                    f'<text x="{x + cell_width / 2}" y="{y + cell_height / 2 + 4}" text-anchor="middle" font-family="Helvetica Neue, Arial, sans-serif" font-size="12" fill="{text_color}">{svg_escape(svg_compact(value))}</text>'
                )

    legend_x = 48
    legend_y = height - 42
    for step in range(6):
        value = max_value * step / 5
        x = legend_x + step * 90
        parts.append(f'<rect x="{x}" y="{legend_y - 16}" width="44" height="16" rx="6" fill="{heat_color(value)}"/>')
        parts.append(
            f'<text x="{x + 22}" y="{legend_y + 18}" text-anchor="middle" font-family="Helvetica Neue, Arial, sans-serif" font-size="13" fill="#64748b">{svg_escape(svg_compact(value))}</text>'
        )

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
      --bg: #080b10;
      --panel: rgba(18, 24, 34, 0.86);
      --panel-strong: #111827;
      --ink: #eef2f7;
      --muted: #94a3b8;
      --soft: #cbd5e1;
      --accent: #d7b56d;
      --accent-2: #72d6c9;
      --campaign: #ff784f;
      --campaign-2: #56c7ff;
      --danger: #e06a6a;
      --border: rgba(148, 163, 184, 0.22);
      --grid: rgba(148, 163, 184, 0.14);
      --shadow: 0 28px 90px rgba(0, 0, 0, 0.34);
      --tv_abierta: #d36a5f;
      --tv_cable: #d99f62;
      --radio: #6ac4b0;
      --via_publica: #a889dd;
      --digital: #6aa5e8;
      --prensa: #9aa7b7;
    }
    * { box-sizing: border-box; }
    html { scroll-behavior: smooth; }
    body {
      margin: 0;
      font-family: "Avenir Next", "Helvetica Neue", Arial, sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at 18% 0%, rgba(255, 120, 79, 0.18), transparent 28%),
        radial-gradient(circle at 82% 10%, rgba(86, 199, 255, 0.12), transparent 26%),
        linear-gradient(180deg, #07090d 0%, #0d121b 48%, #07090d 100%);
      min-height: 100vh;
    }
    body::before {
      content: "";
      position: fixed;
      inset: 0;
      pointer-events: none;
      background-image: linear-gradient(rgba(255,255,255,0.025) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.02) 1px, transparent 1px);
      background-size: 44px 44px;
      mask-image: linear-gradient(180deg, rgba(0,0,0,0.65), transparent 78%);
    }
    a { color: var(--accent); }
    .page {
      max-width: 1320px;
      margin: 0 auto;
      padding: 28px 20px 56px;
      position: relative;
    }
    .topbar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      margin-bottom: 42px;
      color: var(--muted);
      font-size: 0.88rem;
      letter-spacing: 0.04em;
      text-transform: uppercase;
    }
    .brand-mark {
      display: inline-flex;
      align-items: center;
      gap: 10px;
      color: var(--ink);
      font-weight: 700;
    }
    .brand-mark::before {
      content: "";
      width: 9px;
      height: 9px;
      border-radius: 999px;
      background: var(--accent);
      box-shadow: 0 0 28px rgba(255, 120, 79, 0.58);
    }
    .nav-links { display: flex; gap: 18px; flex-wrap: wrap; }
    .nav-links a { color: var(--muted); text-decoration: none; }
    .nav-links a:hover { color: var(--ink); }
    .kicker {
      color: var(--campaign);
      font-size: 0.78rem;
      letter-spacing: 0.2em;
      text-transform: uppercase;
      margin-bottom: 18px;
      font-weight: 800;
    }
    .hero {
      display: grid;
      grid-template-columns: minmax(0, 1.7fr) minmax(280px, 0.8fr);
      gap: 22px;
      margin-bottom: 24px;
      align-items: stretch;
    }
    .panel {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 26px;
      box-shadow: var(--shadow);
      padding: 26px;
      backdrop-filter: blur(18px);
    }
    .hero-main {
      min-height: 470px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      background:
        linear-gradient(135deg, rgba(255, 120, 79, 0.16), transparent 36%),
        radial-gradient(circle at 90% 12%, rgba(86, 199, 255, 0.16), transparent 24%),
        linear-gradient(160deg, rgba(255,255,255,0.05), transparent 42%),
        var(--panel);
      position: relative;
      overflow: hidden;
    }
    .hero-main::after {
      content: "NO";
      position: absolute;
      right: clamp(18px, 4vw, 58px);
      top: clamp(18px, 5vw, 70px);
      font-family: Georgia, "Times New Roman", serif;
      font-size: clamp(5rem, 16vw, 15rem);
      line-height: 0.78;
      color: rgba(255, 120, 79, 0.08);
      letter-spacing: -0.12em;
      pointer-events: none;
    }
    .context-panel {
      background:
        linear-gradient(180deg, rgba(255,255,255,0.055), rgba(255,255,255,0.018)),
        rgba(12, 17, 25, 0.9);
    }
    h1, h2, h3 { margin: 0 0 10px; }
    h1 {
      font-family: Georgia, "Times New Roman", serif;
      font-size: clamp(2.7rem, 7.6vw, 7.6rem);
      line-height: 0.84;
      letter-spacing: -0.075em;
      max-width: 980px;
    }
    .title-small {
      display: block;
      font-family: "Avenir Next", "Helvetica Neue", Arial, sans-serif;
      font-size: clamp(1rem, 1.8vw, 1.45rem);
      line-height: 1;
      letter-spacing: 0.18em;
      text-transform: uppercase;
      color: var(--campaign-2);
      margin-bottom: 12px;
    }
    .title-punch {
      color: var(--campaign);
      text-shadow: 0 0 42px rgba(255, 120, 79, 0.22);
    }
    .title-rest { display: block; }
    h2 { font-size: clamp(1.35rem, 2.2vw, 2.05rem); letter-spacing: -0.035em; }
    h3 { font-size: 0.8rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.16em; }
    p { margin: 0 0 10px; line-height: 1.65; }
    .lede { font-size: clamp(1.05rem, 1.7vw, 1.28rem); max-width: 68ch; color: var(--soft); }
    .note { color: var(--muted); font-size: 0.94rem; }
    .meta { margin: 0; }
    .meta dt { font-size: 0.72rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.14em; }
    .meta dd { margin: 5px 0 18px; font-weight: 700; color: var(--ink); overflow-wrap: anywhere; }
    .stats {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 14px;
      margin: 30px 0 0;
    }
    .stat {
      background: rgba(255, 255, 255, 0.045);
      border: 1px solid var(--border);
      border-radius: 20px;
      padding: 18px;
    }
    .stat .label {
      display: block;
      font-size: 0.78rem;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.14em;
      margin-bottom: 6px;
    }
    .stat strong { font-size: clamp(1.22rem, 2vw, 1.85rem); letter-spacing: -0.04em; }
    .stat small {
      display: block;
      color: var(--muted);
      line-height: 1.45;
      margin-top: 7px;
      font-size: 0.82rem;
    }
    .briefing {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      margin-top: 22px;
    }
    .briefing-card {
      border: 1px solid var(--border);
      background:
        linear-gradient(150deg, rgba(255, 120, 79, 0.10), rgba(86, 199, 255, 0.06)),
        rgba(0, 0, 0, 0.18);
      border-radius: 24px;
      padding: 18px;
    }
    .briefing-card span {
      display: block;
      color: var(--campaign-2);
      font-size: 0.72rem;
      text-transform: uppercase;
      letter-spacing: 0.14em;
      font-weight: 800;
      margin-bottom: 7px;
    }
    .briefing-card strong { display: block; color: var(--ink); line-height: 1.35; }
    .legend {
      display: flex;
      flex-wrap: wrap;
      gap: 10px 16px;
      margin-top: 20px;
    }
    .legend span {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      font-size: 0.92rem;
      color: var(--soft);
    }
    .legend i {
      width: 12px;
      height: 12px;
      border-radius: 999px;
      display: inline-block;
    }
    .line-legend {
      display: flex;
      flex-wrap: wrap;
      gap: 10px 16px;
      margin-top: 18px;
    }
    .line-legend span {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      font-size: 0.9rem;
      color: var(--soft);
    }
    .line-legend i {
      width: 12px;
      height: 12px;
      border-radius: 999px;
      display: inline-block;
    }
    .charts {
      display: grid;
      grid-template-columns: 1fr;
      gap: 22px;
      margin-bottom: 24px;
    }
    .section-label {
      color: var(--campaign);
      letter-spacing: 0.16em;
      text-transform: uppercase;
      font-size: 0.76rem;
      font-weight: 800;
      margin-bottom: 10px;
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
      overflow: visible;
    }
    .axis-label { fill: var(--muted); font-size: 12px; }
    .axis-line, .grid-line { stroke: var(--grid); stroke-width: 1; }
    .line-label { font-size: 11px; font-weight: bold; }
    .table-wrap {
      overflow: auto;
      border: 1px solid var(--border);
      border-radius: 20px;
      max-height: 680px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.95rem;
      background: rgba(8, 11, 16, 0.42);
    }
    th, td {
      padding: 13px 14px;
      border-bottom: 1px solid var(--grid);
      text-align: left;
      vertical-align: top;
      color: var(--soft);
    }
    th {
      font-size: 0.78rem;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.1em;
      position: sticky;
      top: 0;
      background: #101722;
      z-index: 2;
      cursor: default;
    }
    tbody tr:hover { background: rgba(255,255,255,0.045); }
    td strong { color: var(--ink); }
    .media-badge {
      display: inline-flex;
      align-items: center;
      border: 1px solid var(--border);
      border-radius: 999px;
      padding: 3px 9px;
      font-size: 0.78rem;
      color: var(--ink);
      background: rgba(255,255,255,0.055);
      margin-bottom: 6px;
    }
    #piecesTable th:first-child,
    #piecesTable td:first-child {
      min-width: 170px;
      white-space: nowrap;
    }
    .viewer {
      display: grid;
      grid-template-columns: 320px 1fr;
      gap: 22px;
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
      border-radius: 14px;
      background: rgba(255, 255, 255, 0.055);
      color: var(--ink);
      font: inherit;
    }
    .control select option { background: #101722; color: var(--ink); }
    .methodology {
      display: grid;
      grid-template-columns: minmax(0, 1fr) minmax(260px, 0.7fr);
      gap: 22px;
      margin-top: 24px;
    }
    .method-list {
      display: grid;
      gap: 14px;
      margin: 18px 0 0;
      padding: 0;
      list-style: none;
    }
    .method-list li {
      border-top: 1px solid var(--grid);
      padding-top: 14px;
      color: var(--soft);
    }
    .pill {
      display: inline-block;
      padding: 4px 10px;
      border-radius: 999px;
      background: rgba(114, 214, 201, 0.12);
      color: var(--accent-2);
      font-size: 0.85rem;
      font-weight: bold;
    }
    .footer {
      margin-top: 32px;
      color: var(--muted);
      display: flex;
      justify-content: space-between;
      gap: 16px;
      flex-wrap: wrap;
      border-top: 1px solid var(--grid);
      padding-top: 22px;
      font-size: 0.92rem;
    }
    @media (max-width: 960px) {
      .hero, .charts, .viewer, .stats, .methodology { grid-template-columns: 1fr; }
      .briefing { grid-template-columns: 1fr; }
      .page { padding: 18px 14px 40px; }
      .topbar { align-items: flex-start; flex-direction: column; margin-bottom: 28px; }
      .hero-main { min-height: auto; }
      h1 { font-size: clamp(2.7rem, 16vw, 4.8rem); }
      .panel { padding: 20px; border-radius: 22px; }
      svg { min-width: 760px; }
      .chart-wrap { margin: 0 -4px; }
    }
  </style>
</head>
<body>
  <div class="page">
    <header class="topbar">
      <div class="brand-mark">Observatorio publicitario</div>
      <nav class="nav-links" aria-label="Navegacion principal">
        <a href="#serie">Serie semanal</a>
        <a href="#tabla">Tabla</a>
        <a href="#piezas">Piezas</a>
        <a href="#metodologia">Metodo</a>
      </nav>
    </header>
    <section class="hero">
      <div class="panel hero-main">
        <div>
          <div class="kicker">Datos abiertos · Chile · Publicidad observada</div>
          <h1><span class="title-small">Juego ilegal</span><span class="title-punch">Bombardeo</span><span class="title-rest">publicitario a los jovenes</span></h1>
        </div>
        <p class="lede">Esta pagina muestra una estimacion de la inversion publicitaria observada en el dominio publico. Registra apariciones en television, radio, internet, via publica y otros soportes, y las valoriza con tarifas estandar para aproximar con buena precision cuanto estan invirtiendo las marcas observadas.</p>
        <p class="note">Lo que ves aqui no es una factura ni una declaracion corporativa directa, sino una medicion de publicidad visible en el dominio publico multiplicada por una tarifa estandar. Por esa metodologia, los montos pueden presentar diferencias menores respecto de los valores efectivamente transados o facturados.</p>
        <div class="briefing" id="briefing"></div>
        <div class="stats" id="stats"></div>
        <div class="legend" id="legend"></div>
      </div>
      <aside class="panel context-panel">
        <h3>Contexto</h3>
        <dl class="meta">
          <dt>Moneda</dt>
          <dd id="metaCurrency"></dd>
          <dt>Fuente</dt>
          <dd id="metaSource"></dd>
          <dt>Hoja</dt>
          <dd id="metaSheet"></dd>
          <dt>Actualizado</dt>
          <dd id="metaUpdatedAt"></dd>
          <dt>QA</dt>
          <dd><span class="pill" id="metaQa"></span></dd>
        </dl>
      </aside>
    </section>

    <section class="charts" id="serie">
      <article class="panel">
        <div class="section-label">Chart principal</div>
        <h2>Mapa de calor semanal por marca</h2>
        <p class="note">Cada celda representa un corte semanal de 2026 y evita la superposicion de valores: cuanto mas intenso el color, mayor la inversion estimada de esa marca en ese corte.</p>
        <div class="chart-wrap"><svg id="lineChart" viewBox="0 0 760 560" aria-label="Mapa de calor semanal"></svg></div>
        <div class="line-legend" id="lineLegend"></div>
      </article>
      <article class="panel">
        <div class="section-label">Composicion</div>
        <h2>Distribucion estimada por marca y medio</h2>
        <p class="note">Cada barra resume la estimacion total por marca y la descompone por tipo de medio segun la publicidad observada y valorizada con tarifa estandar.</p>
        <div class="chart-wrap"><svg id="stackedBars" viewBox="0 0 960 560" aria-label="Grafico de barras stackeadas"></svg></div>
      </article>
    </section>

    <section class="panel" id="tabla">
      <div class="section-label">Tabla de lectura rapida</div>
      <h2>Tabla resumen</h2>
      <p class="note">Totales acumulados estimados por marca en CLP a partir de publicidad observada y valorizada con tarifa estandar.</p>
      <div class="table-wrap">
        <table id="summaryTable"></table>
      </div>
    </section>

    <section class="viewer" id="piezas">
      <aside class="panel controls">
        <div>
          <div class="section-label">Exploracion</div>
          <h2>Explorador de piezas</h2>
          <p class="note">Filtra por marca, medio y texto para revisar piezas, programas, evidencia y magnitud estimada. En el sitio publicado se carga automaticamente el JSON maestro.</p>
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
        <div class="control">
          <label for="sortFilter">Ordenar por</label>
          <select id="sortFilter">
            <option value="investment">Mayor inversion estimada</option>
            <option value="observations">Mas apariciones</option>
            <option value="recent">Mas reciente</option>
            <option value="brand">Marca</option>
          </select>
        </div>
        <p class="note" id="piecesStatus">Intentando cargar el JSON maestro. Si no esta disponible, se mostrara una muestra embebida.</p>
      </aside>
      <section class="panel">
        <div class="table-wrap">
          <table id="piecesTable"></table>
        </div>
      </section>
    </section>

    <section class="methodology" id="metodologia">
      <article class="panel">
        <div class="section-label">Metodo</div>
        <h2>Como leer estos datos</h2>
        <p class="lede">La medicion cruza apariciones publicitarias observadas en espacios publicos con tarifas estandar de mercado. El resultado es una estimacion comparable de inversion, no una declaracion tributaria ni un registro contable de las empresas.</p>
        <ul class="method-list">
          <li><strong>Observacion:</strong> cada fila registra fecha, marca, medio, programa, tipo de aviso y evidencia disponible.</li>
          <li><strong>Valorizacion:</strong> la aparicion se multiplica por una tarifa estandar para estimar inversion neta.</li>
          <li><strong>Publicacion:</strong> se excluyen marcas reguladas en Chile y se publican agregados semanales junto con una muestra explorable de piezas.</li>
        </ul>
      </article>
      <aside class="panel">
        <h3>Notas editoriales</h3>
        <p class="note">Los montos estan expresados en pesos chilenos. Las semanas cierran en domingo. La lectura correcta es comparativa: magnitud, tendencia y composicion por medio.</p>
        <p class="note">Este observatorio documenta publicidad observable; no reemplaza asesorias legales, regulatorias ni financieras.</p>
        <p class="note">Los datos se verifican con informacion de <a href="https://www.megatime.cl/" target="_blank" rel="noreferrer">megatime.cl</a>, empresa chilena especializada en monitoreo, verificacion y valorizacion de publicidad en medios, usada para observar apariciones publicitarias y estimar inversion con tarifas estandar.</p>
        <a id="repoLink" target="_blank" rel="noreferrer">Repositorio y datos abiertos</a>
      </aside>
    </section>

    <footer class="footer">
      <span>Observatorio abierto de publicidad de apuestas online en Chile</span>
      <span>Actualizado segun la version publicada del sitio</span>
    </footer>
  </div>

  <script id="payload" type="application/json">__PAYLOAD__</script>
  <script>
    const payload = JSON.parse(document.getElementById("payload").textContent);
    const numberFormatter = new Intl.NumberFormat("es-CL", { maximumFractionDigits: 0 });
    const compactFormatter = new Intl.NumberFormat("es-CL", { notation: "compact", maximumFractionDigits: 1 });
    const periodFormatter = new Intl.DateTimeFormat("es-CL", { day: "2-digit", month: "short", year: "numeric" });
    let pieceRecords = payload.sample_records.slice();
    let usingEmbeddedSamples = true;

    function formatMoney(value) {
      return "$" + numberFormatter.format(Math.round(value));
    }

    function formatCompact(value) {
      return "$" + compactFormatter.format(value);
    }

    function prettyPeriod(value) {
      const [year, month, day] = value.split("-").map(Number);
      return periodFormatter.format(new Date(year, month - 1, day));
    }

    function compactPeriod(value) {
      const [year, month, day] = value.split("-").map(Number);
      return new Intl.DateTimeFormat("es-CL", { day: "2-digit", month: "short" })
        .format(new Date(year, month - 1, day))
        .replace(".", "");
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

    function brandColor(index) {
      const palette = ["#8b1e3f", "#0b6e4f", "#2563eb", "#f97316", "#7c3aed", "#b91c1c", "#0f766e", "#475569", "#d97706", "#4f46e5"];
      return palette[index % palette.length];
    }

    function setMeta() {
      document.getElementById("metaCurrency").textContent = payload.currency;
      document.getElementById("metaSource").textContent = payload.source_file;
      document.getElementById("metaSheet").textContent = payload.source_sheet;
      const pageUpdatedAt = document.lastModified
        ? new Intl.DateTimeFormat("es-CL", { dateStyle: "short", timeStyle: "short" }).format(new Date(document.lastModified))
        : "No disponible";
      document.getElementById("metaUpdatedAt").textContent = pageUpdatedAt;
      document.getElementById("metaQa").textContent = payload.qa_passed ? "QA OK (" + payload.qa_checks_run + " chequeos)" : "QA con observaciones";
      document.getElementById("repoLink").href = payload.repo_url;
      document.getElementById("repoLink").textContent = payload.repo_url;

      const totalInvestment = payload.brand_totals.reduce((sum, item) => sum + item.total, 0);
      const topBrand = payload.brand_totals[0];
      const latestPeriod = payload.periods[payload.periods.length - 1];
      const latestLeader = payload.brand_totals
        .map((item) => ({ brand_name: item.brand_name, value: item.series[latestPeriod] || 0 }))
        .sort((left, right) => right.value - left.value)[0];
      const leadingMedia = payload.media_order
        .map((slug) => ({
          slug,
          value: payload.brand_totals.reduce((sum, item) => sum + (item.media_breakdown[slug] || 0), 0)
        }))
        .sort((left, right) => right.value - left.value)[0];
      document.getElementById("briefing").innerHTML = [
        { label: "Ultimo corte", value: latestLeader.brand_name + " lidera la semana con " + formatCompact(latestLeader.value) },
        { label: "Mayor acumulado", value: topBrand.brand_name + " concentra " + formatCompact(topBrand.total) + " del periodo" },
        { label: "Medio dominante", value: mediaLabel(leadingMedia.slug) + " suma " + formatCompact(leadingMedia.value) }
      ].map((item) => '<div class="briefing-card"><span>' + item.label + '</span><strong>' + item.value + '</strong></div>').join("");
      const stats = [
        { label: "Marcas", value: payload.brands.length, note: "incluidas en el producto publico" },
        { label: "Cortes", value: payload.periods.length, note: "semanas con cierre dominical" },
        { label: "Inversion total", value: formatCompact(totalInvestment), note: "estimacion acumulada en CLP" },
        { label: "Marca lider", value: topBrand.brand_name, note: formatCompact(topBrand.total) + " acumulados" }
      ];
      document.getElementById("stats").innerHTML = stats.map((item) =>
        '<div class="stat"><span class="label">' + item.label + '</span><strong>' + item.value + '</strong><small>' + item.note + '</small></div>'
      ).join("");

      document.getElementById("legend").innerHTML = payload.media_order.map((slug) =>
        '<span><i style="background:' + colorFor(slug) + '"></i>' + mediaLabel(slug) + '</span>'
      ).join("");
    }

    function renderStackedBars() {
      const svg = document.getElementById("stackedBars");
      const width = 980;
      const height = Math.max(520, 90 + payload.brand_totals.length * 42);
      svg.setAttribute("viewBox", "0 0 " + width + " " + height);
      const margin = { top: 24, right: 118, bottom: 36, left: 170 };
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
        content += '<text class="axis-label" x="' + (width - 12) + '" y="' + (y + 16) + '" text-anchor="end">' + formatMoney(item.total) + '</text>';
      });

      svg.innerHTML = content;
    }

    function renderLineChart() {
      const svg = document.getElementById("lineChart");
      const width = 980;
      const height = Math.max(540, 120 + payload.brand_totals.length * 42);
      svg.setAttribute("viewBox", "0 0 " + width + " " + height);
      const margin = { top: 70, right: 24, bottom: 70, left: 180 };
      const plotWidth = width - margin.left - margin.right;
      const plotHeight = height - margin.top - margin.bottom;
      const maxValue = Math.max(...payload.brand_totals.flatMap((item) => payload.periods.map((period) => item.series[period] || 0)), 1);
      const cellWidth = plotWidth / Math.max(payload.periods.length, 1);
      const cellHeight = plotHeight / Math.max(payload.brand_totals.length, 1);
      function heatColor(value) {
        const ratio = Math.pow(Math.min(Math.max(value / maxValue, 0), 1), 0.55);
        const start = [244, 241, 232];
        const end = [139, 30, 63];
        const channels = start.map((channel, index) => Math.round(channel + (end[index] - channel) * ratio));
        return "rgb(" + channels.join(",") + ")";
      }
      let content = "";

      payload.periods.forEach((period, index) => {
        const x = margin.left + cellWidth * index;
        content += '<line class="axis-line" x1="' + x + '" y1="' + margin.top + '" x2="' + x + '" y2="' + (height - margin.bottom) + '"></line>';
        content += '<text class="axis-label" x="' + (x + cellWidth / 2) + '" y="' + (margin.top - 12) + '" text-anchor="middle">' + compactPeriod(period) + '</text>';
      });

      payload.brand_totals.forEach((item, rowIndex) => {
        const y = margin.top + cellHeight * rowIndex;
        content += '<text class="axis-label" x="' + (margin.left - 12) + '" y="' + (y + cellHeight / 2 + 4) + '" text-anchor="end">' + item.brand_name + '</text>';
        payload.periods.forEach((period, columnIndex) => {
          const value = item.series[period] || 0;
          const x = margin.left + cellWidth * columnIndex;
          const textColor = value / maxValue > 0.45 ? '#ffffff' : '#1f2937';
          content += '<rect x="' + (x + 2) + '" y="' + (y + 2) + '" width="' + (cellWidth - 4) + '" height="' + (cellHeight - 4) + '" rx="6" fill="' + heatColor(value) + '"></rect>';
          if (value > 0) {
            content += '<text x="' + (x + cellWidth / 2) + '" y="' + (y + cellHeight / 2 + 4) + '" text-anchor="middle" font-size="11" fill="' + textColor + '">' + formatCompact(value) + '</text>';
          }
        });
      });

      svg.innerHTML = content;
      document.getElementById("lineLegend").innerHTML =
        '<span><i style="background:rgb(244,241,232)"></i>Menor inversion semanal</span>' +
        '<span><i style="background:rgb(139,30,63)"></i>Mayor inversion semanal</span>';
    }

    function renderSummaryTable() {
      const table = document.getElementById("summaryTable");
      const header = ['<thead><tr><th>Marca</th>', ...payload.periods.map((period) => '<th>' + prettyPeriod(period) + '</th>'), '<th>Total</th></tr></thead>'].join("");
      const rows = payload.brand_totals.map((item) => {
        return '<tr><td><strong>' + item.brand_name + '</strong></td>' +
          payload.periods.map((period) => '<td>' + formatMoney(item.series[period] || 0) + '</td>').join('') +
          '<td>' + formatMoney(item.total) + '</td></tr>';
      }).join("");
      const totalsByPeriod = payload.periods.map((period) =>
        payload.brand_totals.reduce((sum, item) => sum + (item.series[period] || 0), 0)
      );
      const grandTotal = payload.brand_totals.reduce((sum, item) => sum + item.total, 0);
      const totalRow = '<tr><td><strong>Total</strong></td>' +
        totalsByPeriod.map((value) => '<td><strong>' + formatMoney(value) + '</strong></td>').join('') +
        '<td><strong>' + formatMoney(grandTotal) + '</strong></td></tr>';
      table.innerHTML = header + '<tbody>' + rows + totalRow + '</tbody>';
    }

    function buildPiecesFilters(records) {
      const brandFilter = document.getElementById("brandFilter");
      const mediaFilter = document.getElementById("mediaFilter");
      const brands = Array.from(new Set(records.map((item) => item.brand_name).filter(Boolean))).sort();
      const media = Array.from(new Set(records.map((item) => item.media_type).filter(Boolean))).sort();
      brandFilter.innerHTML = '<option value="">Todas</option>' + brands.map((item) => '<option value="' + item + '">' + item + '</option>').join('');
      mediaFilter.innerHTML = '<option value="">Todos</option>' + media.map((item) => '<option value="' + item + '">' + item + '</option>').join('');
    }

    function aggregatePieceRecords(records) {
      const groups = new Map();
      records.forEach((item) => {
        const key = [
          item.brand_name || '',
          item.media_type || '',
          item.outlet_name || '',
          item.program_name || '',
          item.ad_type || '',
          item.creative_version || '',
          item.evidence_url || ''
        ].join('||');
        if (!groups.has(key)) {
          groups.set(key, {
            brand_name: item.brand_name || '',
            media_type: item.media_type || '',
            outlet_name: item.outlet_name || '',
            program_name: item.program_name || '',
            ad_type: item.ad_type || '',
            creative_version: item.creative_version || '',
            evidence_url: item.evidence_url || '',
            net_investment: 0,
            observations: 0,
            first_seen_at: item.observed_at || '',
            last_seen_at: item.observed_at || ''
          });
        }
        const group = groups.get(key);
        group.net_investment += Number(item.net_investment || 0);
        group.observations += 1;
        if (item.observed_at && (!group.first_seen_at || item.observed_at < group.first_seen_at)) group.first_seen_at = item.observed_at;
        if (item.observed_at && (!group.last_seen_at || item.observed_at > group.last_seen_at)) group.last_seen_at = item.observed_at;
      });
      return Array.from(groups.values());
    }

    function renderPiecesTable() {
      const brandValue = document.getElementById("brandFilter").value;
      const mediaValue = document.getElementById("mediaFilter").value;
      const searchValue = document.getElementById("searchFilter").value.trim().toLowerCase();
      const sortValue = document.getElementById("sortFilter").value;
      const rows = aggregatePieceRecords(pieceRecords)
        .filter((item) => !brandValue || item.brand_name === brandValue)
        .filter((item) => !mediaValue || item.media_type === mediaValue)
        .filter((item) => {
          if (!searchValue) return true;
          return [item.outlet_name, item.program_name, item.creative_version, item.ad_type].join(' ').toLowerCase().includes(searchValue);
        })
        .sort((left, right) => {
          if (sortValue === "observations") return (right.observations || 0) - (left.observations || 0);
          if (sortValue === "recent") return (right.last_seen_at || "").localeCompare(left.last_seen_at || "");
          if (sortValue === "brand") return (left.brand_name || "").localeCompare(right.brand_name || "");
          return (right.net_investment || 0) - (left.net_investment || 0);
        })
        .slice(0, 80);

      const table = document.getElementById("piecesTable");
      table.innerHTML =
        '<thead><tr><th>Periodo</th><th>Marca</th><th>Medio</th><th>Programa</th><th>Pieza</th><th>Apariciones</th><th>Inversion neta</th><th>Evidencia</th></tr></thead>' +
        '<tbody>' +
        rows.map((item) => '<tr>' +
          '<td>' + (item.first_seen_at === item.last_seen_at ? (item.first_seen_at || '') : [item.first_seen_at || '', item.last_seen_at || ''].filter(Boolean).join(' a ')) + '</td>' +
          '<td><strong>' + (item.brand_name || '') + '</strong></td>' +
          '<td>' + (item.media_type ? '<span class="media-badge">' + item.media_type + '</span><br>' : '') + (item.outlet_name || '') + '</td>' +
          '<td>' + (item.program_name || '') + '</td>' +
          '<td>' + [item.ad_type, item.creative_version].filter(Boolean).join(' / ') + '</td>' +
          '<td>' + (item.observations || 0) + '</td>' +
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

      ["brandFilter", "mediaFilter", "searchFilter", "sortFilter"].forEach((id) => {
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


def compare_aggregations(
    current_aggregations: dict[str, dict[str, dict[str, float]]],
    previous_aggregations: dict[str, dict[str, dict[str, float]]],
    brands: list[str],
    periods: list[str],
) -> dict[str, dict[str, dict[str, float]]]:
    scopes = sorted(set(current_aggregations) | set(previous_aggregations))
    deltas: dict[str, dict[str, dict[str, float]]] = defaultdict(lambda: defaultdict(dict))

    for scope in scopes:
        for brand in brands:
            for period in periods:
                current_value = current_aggregations.get(scope, {}).get(brand, {}).get(period, 0.0)
                previous_value = previous_aggregations.get(scope, {}).get(brand, {}).get(period, 0.0)
                deltas[scope][brand][period] = round(current_value - previous_value, 2)

    return deltas


def build_changes_report(
    current_input: Path,
    previous_input: Path | None,
    delta_aggregations: dict[str, dict[str, dict[str, float]]],
) -> dict[str, Any]:
    changed_brands: list[dict[str, Any]] = []
    total_scope = delta_aggregations.get("total", {})

    for brand in sorted(total_scope):
        total_change = round(sum(total_scope[brand].values()), 2)
        if total_change != 0:
            changed_brands.append({"brand_name": brand, "total_change": total_change})

    changed_brands.sort(key=lambda item: abs(item["total_change"]), reverse=True)
    return {
        "current_input": format_cut_label(current_input),
        "previous_input": format_cut_label(previous_input) if previous_input else None,
        "changed_brand_count": len(changed_brands),
        "changed_brands": changed_brands,
    }


def build_validation_report(
    input_path: Path,
    worksheet_name: str | None,
    records: list[dict[str, str]],
    product_brands: list[str],
    product_periods: list[str],
    table_names: list[str],
    previous_input_path: Path | None,
    qa_passed: bool,
    errors: list[str],
) -> dict[str, Any]:
    return {
        "input_file": format_cut_label(input_path),
        "previous_input_file": format_cut_label(previous_input_path) if previous_input_path else None,
        "worksheet_name": worksheet_name,
        "raw_record_count": len(records),
        "product_record_count": sum(1 for record in records if record["brand_name"] not in EXCLUDED_PRODUCT_BRANDS),
        "product_brands": product_brands,
        "excluded_product_scope": "marcas reguladas en Chile",
        "excluded_product_brand_count": len(EXCLUDED_PRODUCT_BRANDS),
        "period_granularity": "week",
        "periods": product_periods,
        "tables_generated": table_names,
        "qa_passed": qa_passed,
        "error_count": len(errors),
        "errors": errors,
    }


def main() -> int:
    args = parse_args()
    previous_input = args.previous_input or default_previous_workbook(args.input)
    raw_sheet_name = resolve_available_sheet_name(args.input, RAW_SHEET_CANDIDATES)
    records = load_records(args.input)
    public_records = published_records(records)
    errors = validate_records(records)

    if errors:
        write_json(
            VALIDATION_OUTPUT,
            build_validation_report(
                input_path=args.input,
                worksheet_name=raw_sheet_name,
                records=records,
                product_brands=[],
                product_periods=[],
                table_names=[],
                previous_input_path=previous_input,
                qa_passed=False,
                errors=errors,
            ),
        )
        raise SystemExit("Validation failed. See output/master/validation_report.json for details.")

    write_csv(PROCESSED_DETAIL_OUTPUT, list(CANONICAL_FIELD_ORDER), public_records)
    write_csv(MASTER_CSV_OUTPUT, list(CANONICAL_FIELD_ORDER), public_records)
    write_json(MASTER_JSON_OUTPUT, public_records)

    periods, brands, aggregations = aggregate_period_tables(records, "week_ending")
    monthly_periods, _, monthly_aggregations = aggregate_period_tables(records, "month")
    summary_fieldnames = ["brand_name", *periods, "total"]
    table_names: list[str] = []

    for table_name, values_by_brand in sorted(aggregations.items()):
        output_path = PRODUCT_OUTPUT_DIR / f"{table_name}.csv"
        summary_rows = build_summary_rows(periods, brands, values_by_brand)
        write_csv(output_path, summary_fieldnames, summary_rows)
        table_names.append(f"{table_name}.csv")

    previous_records = load_records(previous_input) if previous_input else []
    previous_periods, previous_brands, previous_aggregations = aggregate_period_tables(previous_records, "week_ending") if previous_records else ([], [], {})
    delta_periods = sort_periods(set(periods) | set(previous_periods))
    delta_brands = sorted(set(brands) | set(previous_brands))
    delta_aggregations = compare_aggregations(aggregations, previous_aggregations, delta_brands, delta_periods)

    for table_name, values_by_brand in sorted(delta_aggregations.items()):
        output_path = CHANGES_OUTPUT_DIR / f"{table_name}.csv"
        summary_rows = build_summary_rows(delta_periods, delta_brands, values_by_brand)
        write_csv(output_path, ["brand_name", *delta_periods, "total"], summary_rows)

    write_json(CHANGES_OUTPUT_DIR / "changes_report.json", build_changes_report(args.input, previous_input, delta_aggregations))

    qa_report = run_qa(args.input, monthly_periods, brands, monthly_aggregations)
    write_json(QA_OUTPUT, qa_report)
    if not qa_report["passed"]:
        raise SystemExit("QA failed. See output/master/qa_report.json for details.")

    visualization_payload = build_visualization_payload(args.input, raw_sheet_name, public_records, periods, brands, aggregations, qa_report)
    write_json(VISUALIZATION_DATA_OUTPUT, visualization_payload)
    visualization_html = build_visualization_html(visualization_payload)
    write_text(VISUALIZATION_HTML_OUTPUT, visualization_html)
    write_text(STACKED_SVG_OUTPUT, build_stacked_bars_svg(visualization_payload))
    write_text(HEATMAP_SVG_OUTPUT, build_lines_svg(visualization_payload))
    write_text(SITE_INDEX_OUTPUT, visualization_html)
    write_json(SITE_SUMMARY_OUTPUT, visualization_payload)
    write_json(SITE_MASTER_OUTPUT, public_records)

    write_json(
        VALIDATION_OUTPUT,
        build_validation_report(
            input_path=args.input,
            worksheet_name=raw_sheet_name,
            records=records,
            product_brands=brands,
            product_periods=periods,
            table_names=table_names + ["changes_report.json"],
            previous_input_path=previous_input,
            qa_passed=qa_report["passed"],
            errors=[],
        ),
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
