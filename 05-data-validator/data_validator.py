"""
Data Validator - Validasi data Excel berdasarkan rules di config YAML
Mendukung konfigurasi via YAML dan CLI arguments (hybrid)
"""

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from common.config import load_config, resolve_config_path, resolve_path  # noqa: E402
from common.excel import read_excel  # noqa: E402
from common.logging_setup import get_logger, setup_logging  # noqa: E402
from common.phone import phone_digits  # noqa: E402

log = get_logger("data_validator")


FILL_ERROR = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
FILL_HEADER = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")


class Validator:
    """Class untuk validasi data."""

    def __init__(self, df: pd.DataFrame):
        self.df = df
        self.errors = []  # List of {row, column, value, rule, message}

    def validate_required(self, column: str, rule: dict) -> None:
        """Validasi field tidak boleh kosong."""
        for idx, value in self.df[column].items():
            if pd.isna(value) or str(value).strip() == "":
                self.errors.append(
                    {
                        "row": idx + 2,  # +2 karena header dan 0-index
                        "column": column,
                        "value": value,
                        "rule": "required",
                        "message": f"{column} tidak boleh kosong",
                    }
                )

    def validate_email(self, column: str, rule: dict) -> None:
        """Validasi format email."""
        email_pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        for idx, value in self.df[column].items():
            if pd.notna(value) and str(value).strip() != "":
                if not re.match(email_pattern, str(value)):
                    self.errors.append(
                        {
                            "row": idx + 2,
                            "column": column,
                            "value": value,
                            "rule": "email",
                            "message": f"Format email tidak valid: {value}",
                        }
                    )

    def validate_phone(self, column: str, rule: dict) -> None:
        """Validasi nomor telepon: minimal dan maksimal digit."""
        min_digits = rule.get("min_digits", 10)
        max_digits = rule.get("max_digits", 15)

        for idx, value in self.df[column].items():
            if pd.notna(value) and str(value).strip() != "":
                digits = phone_digits(value) or ""
                if len(digits) < min_digits:
                    message = f"No HP kurang dari {min_digits} digit: {value}"
                elif len(digits) > max_digits:
                    message = f"No HP lebih dari {max_digits} digit: {value}"
                else:
                    continue

                self.errors.append(
                    {
                        "row": idx + 2,
                        "column": column,
                        "value": value,
                        "rule": "phone",
                        "message": message,
                    }
                )

    def validate_date_range(self, column: str, rule: dict) -> None:
        """Validasi tanggal dalam range."""
        min_date = rule.get("min")
        max_date = rule.get("max")

        if min_date:
            min_date = datetime.strptime(min_date, "%Y-%m-%d")
        if max_date:
            max_date = datetime.strptime(max_date, "%Y-%m-%d")

        for idx, value in self.df[column].items():
            if pd.notna(value) and str(value).strip() != "":
                try:
                    if isinstance(value, datetime):
                        date_val = value
                    else:
                        date_val = pd.to_datetime(value)

                    if min_date and date_val < min_date:
                        self.errors.append(
                            {
                                "row": idx + 2,
                                "column": column,
                                "value": value,
                                "rule": "date_range",
                                "message": (
                                    f'Tanggal sebelum {min_date.strftime("%Y-%m-%d")}: {value}'
                                ),
                            }
                        )
                    elif max_date and date_val > max_date:
                        self.errors.append(
                            {
                                "row": idx + 2,
                                "column": column,
                                "value": value,
                                "rule": "date_range",
                                "message": (
                                    f'Tanggal setelah {max_date.strftime("%Y-%m-%d")}: {value}'
                                ),
                            }
                        )
                except (ValueError, TypeError, OverflowError):
                    self.errors.append(
                        {
                            "row": idx + 2,
                            "column": column,
                            "value": value,
                            "rule": "date_range",
                            "message": f"Format tanggal tidak valid: {value}",
                        }
                    )

    def validate_number_range(self, column: str, rule: dict) -> None:
        """Validasi angka dalam range."""
        min_val = rule.get("min")
        max_val = rule.get("max")

        for idx, value in self.df[column].items():
            if pd.notna(value) and str(value).strip() != "":
                try:
                    num_val = float(value)
                    if min_val is not None and num_val < min_val:
                        self.errors.append(
                            {
                                "row": idx + 2,
                                "column": column,
                                "value": value,
                                "rule": "number_range",
                                "message": f"Nilai kurang dari {min_val}: {value}",
                            }
                        )
                    elif max_val is not None and num_val > max_val:
                        self.errors.append(
                            {
                                "row": idx + 2,
                                "column": column,
                                "value": value,
                                "rule": "number_range",
                                "message": f"Nilai lebih dari {max_val}: {value}",
                            }
                        )
                except (ValueError, TypeError):
                    self.errors.append(
                        {
                            "row": idx + 2,
                            "column": column,
                            "value": value,
                            "rule": "number_range",
                            "message": f"Bukan angka valid: {value}",
                        }
                    )

    def validate_regex(self, column: str, rule: dict) -> None:
        """Validasi dengan regex pattern."""
        pattern = rule.get("pattern")
        message = rule.get("message", f"Tidak sesuai format: {pattern}")

        if not pattern:
            return

        for idx, value in self.df[column].items():
            if pd.notna(value) and str(value).strip() != "":
                if not re.match(pattern, str(value)):
                    self.errors.append(
                        {
                            "row": idx + 2,
                            "column": column,
                            "value": value,
                            "rule": "regex",
                            "message": message,
                        }
                    )

    def validate_in_list(self, column: str, rule: dict) -> None:
        """Validasi nilai harus dalam list."""
        valid_values = rule.get("values", [])

        for idx, value in self.df[column].items():
            if pd.notna(value) and str(value).strip() != "":
                if str(value) not in [str(v) for v in valid_values]:
                    self.errors.append(
                        {
                            "row": idx + 2,
                            "column": column,
                            "value": value,
                            "rule": "in_list",
                            "message": f'Nilai "{value}" tidak valid. '
                            f"Harus salah satu dari: {valid_values}",
                        }
                    )

    def validate_unique(self, column: str, rule: dict) -> None:
        """Validasi nilai harus unik (tidak duplikat)."""
        seen = {}
        for idx, value in self.df[column].items():
            if pd.notna(value) and str(value).strip() != "":
                str_val = str(value)
                if str_val in seen:
                    self.errors.append(
                        {
                            "row": idx + 2,
                            "column": column,
                            "value": value,
                            "rule": "unique",
                            "message": f"Duplikat dengan baris {seen[str_val]}: {value}",
                        }
                    )
                else:
                    seen[str_val] = idx + 2

    def validate(self, rules: dict) -> list:
        """Jalankan semua validasi berdasarkan rules."""
        for column, column_rules in rules.items():
            if column not in self.df.columns:
                log.warning("  Peringatan: Kolom '%s' tidak ditemukan, skip validasi", column)
                continue

            for rule in column_rules:
                rule_type = rule.get("type")

                if rule_type == "required":
                    self.validate_required(column, rule)
                elif rule_type == "email":
                    self.validate_email(column, rule)
                elif rule_type == "phone":
                    self.validate_phone(column, rule)
                elif rule_type == "date_range":
                    self.validate_date_range(column, rule)
                elif rule_type == "number_range":
                    self.validate_number_range(column, rule)
                elif rule_type == "regex":
                    self.validate_regex(column, rule)
                elif rule_type == "in_list":
                    self.validate_in_list(column, rule)
                elif rule_type == "unique":
                    self.validate_unique(column, rule)

        return self.errors


def print_report(total_rows: int, errors: list) -> None:
    """Tampilkan laporan validasi ke console."""
    error_rows = set(e["row"] for e in errors)
    valid_count = total_rows - len(error_rows)

    log.info("")
    log.info("=" * 50)
    log.info("LAPORAN VALIDASI")
    log.info("=" * 50)

    log.info("")
    log.info("Ringkasan:")
    log.info("  Total baris : %s", total_rows)
    log.info("  Valid       : %s", valid_count)
    log.info("  Error       : %s baris (%s masalah)", len(error_rows), len(errors))

    if not errors:
        log.info("")
        log.info("Semua data VALID!")
        return

    log.info("")
    log.info("--- DETAIL ERROR (%s) ---", len(errors))
    for error in errors[:30]:
        log.info("  Baris %s: %s", error["row"], error["message"])

    if len(errors) > 30:
        log.info("  ... dan %s error lainnya", len(errors) - 30)


def export_to_excel(
    df: pd.DataFrame, errors: list, output_path: str, text_columns: tuple[str, ...] = ()
) -> None:
    """Export hasil validasi ke Excel (semua nilai ditulis sebagai teks)."""
    workbook = Workbook()

    ws_summary = workbook.active
    ws_summary.title = "Summary"

    error_rows = set(e["row"] for e in errors)
    valid_count = len(df) - len(error_rows)

    ws_summary["A1"] = "Laporan Validasi Data"
    ws_summary["A1"].font = Font(bold=True, size=14)
    ws_summary["A3"] = "Total Baris:"
    ws_summary["B3"] = len(df)
    ws_summary["A4"] = "Valid:"
    ws_summary["B4"] = valid_count
    ws_summary["A5"] = "Error:"
    ws_summary["B5"] = len(error_rows)
    ws_summary["A6"] = "Total Masalah:"
    ws_summary["B6"] = len(errors)

    if errors:
        ws_errors = workbook.create_sheet("Detail Error")
        ws_errors.append(["Baris", "Kolom", "Nilai", "Rule", "Pesan"])

        for cell in ws_errors[1]:
            cell.fill = FILL_HEADER
            cell.font = Font(bold=True, color="FFFFFF")

        for error in errors:
            ws_errors.append(
                [
                    error["row"],
                    error["column"],
                    str(error["value"]) if error["value"] is not None else "",
                    error["rule"],
                    error["message"],
                ]
            )

    ws_data = workbook.create_sheet("Data")
    ws_data.append(list(df.columns))
    for cell in ws_data[1]:
        cell.fill = FILL_HEADER
        cell.font = Font(bold=True, color="FFFFFF")

    error_cells = {(error["row"], error["column"]) for error in errors}
    text_indexes = [
        index for index, col in enumerate(df.columns, 1) if col in text_columns
    ]

    for idx, row in df.iterrows():
        excel_row = idx + 2
        row_data = [str(row[col]) if pd.notna(row[col]) else "" for col in df.columns]
        ws_data.append(row_data)

        for col_idx in text_indexes:
            ws_data.cell(row=excel_row, column=col_idx).number_format = "@"

        for col_idx, col in enumerate(df.columns, 1):
            if (excel_row, col) in error_cells:
                ws_data.cell(row=excel_row, column=col_idx).fill = FILL_ERROR
                ws_data.cell(row=excel_row, column=col_idx).number_format = "@"

    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_file)
    log.info("")
    log.info("Laporan Excel disimpan ke: %s", output_path)


def phone_columns(rules: dict) -> tuple[str, ...]:
    """Ambil kolom yang punya rule bertipe ``phone``."""
    columns = []
    for column, column_rules in rules.items():
        for rule in column_rules:
            if rule.get("type") == "phone":
                columns.append(column)
                break
    return tuple(columns)


def validate_data(input_file: str, rules: dict, output_file: str = None) -> list:
    """Main function untuk validasi data."""
    if not Path(input_file).exists():
        raise FileNotFoundError(f"File '{input_file}' tidak ditemukan")

    log.info("Membaca file: %s", input_file)
    text_columns = phone_columns(rules)
    df = read_excel(input_file, text_columns=text_columns)
    log.info("Total baris: %s", len(df))
    log.info("Kolom: %s", ", ".join(map(str, df.columns)))

    log.info("")
    log.info("Memvalidasi data...")
    validator = Validator(df)
    errors = validator.validate(rules)

    print_report(len(df), errors)

    if output_file:
        export_to_excel(df, errors, output_file, text_columns)

    return errors


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Validasi data Excel berdasarkan rules di config.")
    parser.add_argument("input", nargs="?", help="File Excel input")
    parser.add_argument("--output", "-o", help="Export hasil ke file Excel")
    parser.add_argument(
        "--config",
        "-c",
        default="config.yaml",
        help="File konfigurasi YAML (default: config.yaml)",
    )
    return parser


def main(argv: list[str] | None = None) -> None:
    setup_logging()
    parser = build_parser()
    args = parser.parse_args(argv)

    script_dir = Path(__file__).parent
    config_path = resolve_config_path(args.config, script_dir)
    config = load_config(config_path)

    # CLI selalu menang; path dari CLI relatif terhadap cwd.
    if args.input:
        input_file = Path(args.input)
    else:
        input_file = resolve_path(script_dir, config["input"]) if config.get("input") else None

    if args.output:
        output_file = Path(args.output)
    elif config.get("output"):
        output_file = resolve_path(script_dir, config["output"])
    else:
        output_file = None

    rules = config.get("rules", {})

    if not input_file:
        log.error("Error: File input harus diisi (via CLI atau config.yaml)")
        parser.print_help()
        sys.exit(1)

    if not rules:
        log.error("Error: Rules validasi harus didefinisikan di config.yaml")
        sys.exit(1)

    try:
        validate_data(
            str(input_file),
            rules,
            str(output_file) if output_file else None,
        )
    except (FileNotFoundError, ValueError) as exc:
        log.error("Error: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
