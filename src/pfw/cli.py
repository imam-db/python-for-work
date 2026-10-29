"""Argparse entry point that dispatches to the pfw modules."""

from __future__ import annotations

import argparse
import sys

from pfw import __version__, clean, diff, load, merge, sample, split, validate

UTILITIES = [
    (merge, "merge", "Merge tabular files from a folder"),
    (clean, "clean", "Clean names, dates, and phones"),
    (split, "split", "Split a tabular file by column"),
    (diff, "diff", "Diff two file versions"),
    (validate, "validate", "Validate rows against YAML rules"),
    (load, "load", "Load a file into DuckDB or Postgres"),
]


def build_parser() -> argparse.ArgumentParser:
    commands = "\n".join(f"  {name:<10} {summary}" for _, name, summary in UTILITIES)
    parser = argparse.ArgumentParser(
        prog="pfw",
        description="Run the Python for Work utilities from one command.",
        epilog=(
            f"commands:\n{commands}\n"
            f"  {'sample':<10} Generate sample data for a utility"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--version", action="version", version=f"pfw {__version__}")

    subparsers = parser.add_subparsers(dest="command", metavar="command")
    for module, name, summary in UTILITIES:
        subparser = subparsers.add_parser(name, help=summary, description=summary)
        module.add_arguments(subparser)
        subparser.set_defaults(handler=module.run)

    sample_parser = subparsers.add_parser(
        "sample",
        help="Generate sample data for a utility",
        description="Generate sample data for a utility.",
    )
    sample.add_arguments(sample_parser)
    sample_parser.set_defaults(handler=sample.run)

    return parser


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    parser = build_parser()

    try:
        args = parser.parse_args(argv)
    except SystemExit as exc:
        return int(exc.code or 0)

    if not getattr(args, "handler", None):
        parser.print_help()
        return 0

    return int(args.handler(args))


if __name__ == "__main__":
    sys.exit(main())
