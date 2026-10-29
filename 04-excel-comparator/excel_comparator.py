"""
Excel Comparator - Membandingkan 2 file Excel dan menampilkan perbedaannya
Mendukung konfigurasi via YAML dan CLI arguments (hybrid)
"""

import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from common.config import load_config, resolve_config_path, resolve_path  # noqa: E402
from common.logging_setup import get_logger, setup_logging  # noqa: E402

log = get_logger("excel_comparator")


# Warna untuk highlight
FILL_ADDED = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Hijau
FILL_DELETED = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")  # Merah
FILL_CHANGED = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # Kuning


def sort_key(value: str):
    """Urutkan key numerik secara numerik, sisanya secara alfabetis."""
    text = str(value)
    try:
        return (0, int(text), "")
    except ValueError:
        return (1, 0, text)


def warn_duplicate_keys(df: pd.DataFrame, key_column: str, label: str) -> None:
    """Catat peringatan jika kolom kunci punya nilai duplikat."""
    keys = df[key_column].astype(str)
    duplicates = keys[keys.duplicated()].unique().tolist()
    if not duplicates:
        return

    preview = ", ".join(duplicates[:10])
    remaining = f" (+{len(duplicates) - 10} lainnya)" if len(duplicates) > 10 else ""
    log.warning(
        "Peringatan: %s kolom kunci '%s' duplikat: %s%s; memakai baris pertama",
        label,
        key_column,
        preview,
        remaining,
    )


def compare_dataframes(
    df_old: pd.DataFrame, df_new: pd.DataFrame, key_column: str = None
) -> dict:
    """
    Bandingkan dua DataFrame dan return perbedaannya.

    Jika ada key duplikat, baris pertama yang dipakai untuk matching.
    """
    result = {"added": [], "deleted": [], "changed": [], "summary": {}}

    if key_column and key_column in df_old.columns and key_column in df_new.columns:
        warn_duplicate_keys(df_old, key_column, "file lama")
        warn_duplicate_keys(df_new, key_column, "file baru")

        old_keys = set(df_old[key_column].astype(str))
        new_keys = set(df_new[key_column].astype(str))

        # Baris baru (ada di new, tidak di old)
        for key in sorted(new_keys - old_keys, key=sort_key):
            row = df_new[df_new[key_column].astype(str) == key].iloc[0].to_dict()
            result["added"].append({"key": key, "data": row})

        # Baris dihapus (ada di old, tidak di new)
        for key in sorted(old_keys - new_keys, key=sort_key):
            row = df_old[df_old[key_column].astype(str) == key].iloc[0].to_dict()
            result["deleted"].append({"key": key, "data": row})

        # Baris yang sama, cek perubahan nilai
        for key in sorted(old_keys & new_keys, key=sort_key):
            old_row = df_old[df_old[key_column].astype(str) == key].iloc[0]
            new_row = df_new[df_new[key_column].astype(str) == key].iloc[0]

            for col in df_old.columns:
                if col not in df_new.columns:
                    continue
                old_val = old_row[col]
                new_val = new_row[col]

                if pd.isna(old_val) and pd.isna(new_val):
                    continue
                if str(old_val) != str(new_val):
                    result["changed"].append(
                        {
                            "key": key,
                            "column": col,
                            "old_value": old_val,
                            "new_value": new_val,
                        }
                    )
    else:
        max_rows = max(len(df_old), len(df_new))

        for i in range(max_rows):
            if i >= len(df_old):
                result["added"].append({"row": i + 1, "data": df_new.iloc[i].to_dict()})
            elif i >= len(df_new):
                result["deleted"].append({"row": i + 1, "data": df_old.iloc[i].to_dict()})
            else:
                for col in df_old.columns:
                    if col not in df_new.columns:
                        continue
                    old_val = df_old.iloc[i][col]
                    new_val = df_new.iloc[i][col]

                    if pd.isna(old_val) and pd.isna(new_val):
                        continue
                    if str(old_val) != str(new_val):
                        result["changed"].append(
                            {
                                "row": i + 1,
                                "column": col,
                                "old_value": old_val,
                                "new_value": new_val,
                            }
                        )

    result["summary"] = {
        "added_count": len(result["added"]),
        "deleted_count": len(result["deleted"]),
        "changed_count": len(result["changed"]),
    }

    return result


def print_report(result: dict, key_column: str = None) -> None:
    """Tampilkan laporan perbandingan ke console."""
    summary = result["summary"]

    log.info("")
    log.info("=" * 50)
    log.info("LAPORAN PERBANDINGAN")
    log.info("=" * 50)

    log.info("")
    log.info("Ringkasan:")
    log.info("  - Baris baru     : %s", summary["added_count"])
    log.info("  - Baris dihapus  : %s", summary["deleted_count"])
    log.info("  - Cell berubah   : %s", summary["changed_count"])

    if not (summary["added_count"] or summary["deleted_count"] or summary["changed_count"]):
        log.info("")
        log.info("Kedua file IDENTIK, tidak ada perbedaan.")
        return

    if result["changed"]:
        log.info("")
        log.info("--- PERUBAHAN (%s) ---", len(result["changed"]))
        for item in result["changed"][:20]:
            if key_column:
                log.info(
                    "  [%s=%s] %s: %s -> %s",
                    key_column,
                    item["key"],
                    item["column"],
                    item["old_value"],
                    item["new_value"],
                )
            else:
                log.info(
                    "  Baris %s, %s: %s -> %s",
                    item["row"],
                    item["column"],
                    item["old_value"],
                    item["new_value"],
                )
        if len(result["changed"]) > 20:
            log.info("  ... dan %s perubahan lainnya", len(result["changed"]) - 20)

    if result["added"]:
        log.info("")
        log.info("--- BARIS BARU (%s) ---", len(result["added"]))
        for item in result["added"][:10]:
            if key_column:
                log.info("  [%s=%s]", key_column, item["key"])
            else:
                log.info("  Baris %s", item["row"])
        if len(result["added"]) > 10:
            log.info("  ... dan %s baris lainnya", len(result["added"]) - 10)

    if result["deleted"]:
        log.info("")
        log.info("--- BARIS DIHAPUS (%s) ---", len(result["deleted"]))
        for item in result["deleted"][:10]:
            if key_column:
                log.info("  [%s=%s]", key_column, item["key"])
            else:
                log.info("  Baris %s", item["row"])
        if len(result["deleted"]) > 10:
            log.info("  ... dan %s baris lainnya", len(result["deleted"]) - 10)


def export_to_excel(result: dict, output_path: str, key_column: str = None) -> None:
    """Export hasil perbandingan ke Excel dengan highlight."""
    workbook = Workbook()

    ws_summary = workbook.active
    ws_summary.title = "Summary"
    ws_summary["A1"] = "Laporan Perbandingan Excel"
    ws_summary["A1"].font = Font(bold=True, size=14)
    ws_summary["A3"] = "Baris Baru:"
    ws_summary["B3"] = result["summary"]["added_count"]
    ws_summary["A4"] = "Baris Dihapus:"
    ws_summary["B4"] = result["summary"]["deleted_count"]
    ws_summary["A5"] = "Cell Berubah:"
    ws_summary["B5"] = result["summary"]["changed_count"]

    if result["changed"]:
        ws_changed = workbook.create_sheet("Perubahan")
        if key_column:
            ws_changed.append([key_column, "Kolom", "Nilai Lama", "Nilai Baru"])
        else:
            ws_changed.append(["Baris", "Kolom", "Nilai Lama", "Nilai Baru"])

        for item in result["changed"]:
            if key_column:
                ws_changed.append(
                    [item["key"], item["column"], str(item["old_value"]), str(item["new_value"])]
                )
            else:
                ws_changed.append(
                    [item["row"], item["column"], str(item["old_value"]), str(item["new_value"])]
                )

        for cell in ws_changed[1]:
            cell.fill = FILL_CHANGED
            cell.font = Font(bold=True)

    if result["added"]:
        ws_added = workbook.create_sheet("Baris Baru")
        headers = list(result["added"][0]["data"].keys())
        ws_added.append(headers)
        for cell in ws_added[1]:
            cell.fill = FILL_ADDED
            cell.font = Font(bold=True)

        for item in result["added"]:
            ws_added.append(list(item["data"].values()))

    if result["deleted"]:
        ws_deleted = workbook.create_sheet("Baris Dihapus")
        headers = list(result["deleted"][0]["data"].keys())
        ws_deleted.append(headers)
        for cell in ws_deleted[1]:
            cell.fill = FILL_DELETED
            cell.font = Font(bold=True)

        for item in result["deleted"]:
            ws_deleted.append(list(item["data"].values()))

    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_file)
    log.info("")
    log.info("Laporan Excel disimpan ke: %s", output_path)


def compare_excel(
    file_old: str,
    file_new: str,
    key_column: str = None,
    output_file: str = None,
) -> dict:
    """Main function untuk membandingkan 2 file Excel."""
    if not Path(file_old).exists():
        raise FileNotFoundError(f"File '{file_old}' tidak ditemukan")
    if not Path(file_new).exists():
        raise FileNotFoundError(f"File '{file_new}' tidak ditemukan")

    log.info("File lama : %s", file_old)
    log.info("File baru : %s", file_new)
    if key_column:
        log.info("Key column: %s", key_column)

    df_old = pd.read_excel(file_old)
    df_new = pd.read_excel(file_new)

    log.info("")
    log.info("File lama: %s baris, %s kolom", len(df_old), len(df_old.columns))
    log.info("File baru: %s baris, %s kolom", len(df_new), len(df_new.columns))

    if key_column:
        if key_column not in df_old.columns:
            raise ValueError(
                f"Kolom '{key_column}' tidak ditemukan di file lama. "
                f"Kolom tersedia: {', '.join(map(str, df_old.columns))}"
            )
        if key_column not in df_new.columns:
            raise ValueError(f"Kolom '{key_column}' tidak ditemukan di file baru")

    result = compare_dataframes(df_old, df_new, key_column)
    print_report(result, key_column)

    if output_file:
        export_to_excel(result, output_file, key_column)

    return result


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Membandingkan 2 file Excel dan menampilkan perbedaannya."
    )
    parser.add_argument("file_old", nargs="?", help="File Excel lama")
    parser.add_argument("file_new", nargs="?", help="File Excel baru")
    parser.add_argument("--key", "-k", help="Kolom kunci untuk matching baris")
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
    if args.file_old:
        file_old = Path(args.file_old)
    else:
        file_old = resolve_path(script_dir, config["file_old"]) if config.get("file_old") else None

    if args.file_new:
        file_new = Path(args.file_new)
    else:
        file_new = resolve_path(script_dir, config["file_new"]) if config.get("file_new") else None

    key_column = args.key or config.get("key_column")

    if args.output:
        output_file = Path(args.output)
    elif config.get("output"):
        output_file = resolve_path(script_dir, config["output"])
    else:
        output_file = None

    if not file_old or not file_new:
        log.error("Error: Kedua file harus diisi (via CLI atau config.yaml)")
        parser.print_help()
        sys.exit(1)

    try:
        compare_excel(
            str(file_old),
            str(file_new),
            key_column,
            str(output_file) if output_file else None,
        )
    except (FileNotFoundError, ValueError) as exc:
        log.error("Error: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
