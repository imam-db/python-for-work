"""Argparse entry point that dispatches to the utility scripts."""

from __future__ import annotations

import argparse
import sys

from pfw.loader import load_module

UTILITIES = {
    "merge": ("01-merge-excel", "merge_excel.py", "Merge workbooks from a folder"),
    "clean": ("02-data-cleaner", "data_cleaner.py", "Clean names, dates, and phones"),
    "split": ("03-excel-splitter", "excel_splitter.py", "Split a workbook by column"),
    "compare": ("04-excel-comparator", "excel_comparator.py", "Diff two workbook versions"),
    "validate": ("05-data-validator", "data_validator.py", "Validate rows against YAML rules"),
}

SAMPLES = {
    "merge": ("01-merge-excel", "generate_sample.py", "generate_sample"),
    "clean": ("02-data-cleaner", "generate_sample.py", "generate_dirty_data"),
    "split": ("03-excel-splitter", "generate_sample.py", "generate_sample_data"),
    "compare": ("04-excel-comparator", "generate_sample.py", "generate_sample_data"),
    "validate": ("05-data-validator", "generate_sample.py", "generate_sample_data"),
}


def build_parser() -> argparse.ArgumentParser:
    commands = "\n".join(
        f"  {name:<10} {summary}" for name, (_, _, summary) in UTILITIES.items()
    )
    return argparse.ArgumentParser(
        prog="pfw",
        description="Run the Python for Work Excel utilities from one command.",
        epilog=f"commands:\n{commands}\n  {'sample':<10} Generate sample data for a utility",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )


def build_sample_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pfw sample",
        description="Generate sample data for a utility.",
    )
    parser.add_argument("name", choices=list(SAMPLES), help="Utility to generate data for")
    parser.add_argument(
        "-o",
        "--output",
        help="Output folder or file (default: the utility's sample/ folder)",
    )
    return parser


def run_utility(command: str, argv: list[str]) -> int:
    folder, filename, _ = UTILITIES[command]
    module = load_module(folder, filename)

    try:
        module.main(argv)
    except SystemExit as exc:
        code = exc.code
        if code is None:
            return 0
        return code if isinstance(code, int) else 1
    return 0


def run_sample(name: str, output: str | None = None) -> None:
    folder, filename, function = SAMPLES[name]
    generator = getattr(load_module(folder, filename), function)

    if output:
        generator(output)
    else:
        generator()


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)

    if not argv or argv[0] in ("-h", "--help"):
        build_parser().print_help()
        return 0

    command, rest = argv[0], argv[1:]

    if command in UTILITIES:
        return run_utility(command, rest)

    if command == "sample":
        try:
            args = build_sample_parser().parse_args(rest)
        except SystemExit as exc:
            return int(exc.code or 0)
        run_sample(args.name, args.output)
        return 0

    build_parser().error(f"unknown command: {command}")
    return 2


if __name__ == "__main__":
    sys.exit(main())
