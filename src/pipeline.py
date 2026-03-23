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
VALIDATION_OUTPUT = ROOT_DIR / "output" / "master" / "validation_report.json"
QA_OUTPUT = ROOT_DIR / "output" / "master" / "qa_report.json"

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
