"""Konfigurasi pytest: buat utility script bisa di-import dari folder bertanda hubung."""

import importlib.util
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]

UTILITY_DIRS = [
    "01-merge-excel",
    "02-data-cleaner",
    "03-excel-splitter",
    "04-excel-comparator",
    "05-data-validator",
]

for relative in [".", *UTILITY_DIRS]:
    directory = str(ROOT / relative)
    if directory not in sys.path:
        sys.path.insert(0, directory)


@pytest.fixture(scope="session")
def load_utility():
    """Load satu file utility dengan nama modul unik (untuk generate_sample)."""
    cache = {}

    def _load(utility: str, filename: str):
        key = (utility, filename)
        if key not in cache:
            path = ROOT / utility / filename
            alias = f"pfw_{utility.replace('-', '_')}_{path.stem}"
            spec = importlib.util.spec_from_file_location(alias, path)
            module = importlib.util.module_from_spec(spec)
            sys.modules[alias] = module
            spec.loader.exec_module(module)
            cache[key] = module
        return cache[key]

    return _load
