"""Pembacaan konfigurasi YAML dan resolusi path."""

from pathlib import Path

import yaml


def load_config(config_path: str | Path) -> dict:
    """Load konfigurasi dari file YAML.

    Mengembalikan dict kosong jika file tidak ada (utility yang punya default
    di kode) atau isinya kosong.
    """
    config_file = Path(config_path)
    if not config_file.exists():
        return {}
    with config_file.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def resolve_path(base_dir: str | Path, value: str | Path) -> Path:
    """Resolve path relatif terhadap ``base_dir``; path absolut dibiarkan."""
    path = Path(value).expanduser()
    if path.is_absolute():
        return path
    return Path(base_dir) / path


def resolve_config_path(value: str | Path, base_dir: str | Path) -> Path:
    """Cari file config: absolut, lalu cwd, lalu folder utility."""
    path = Path(value).expanduser()
    if path.is_absolute():
        return path
    if path.exists():
        return path
    return Path(base_dir) / path
