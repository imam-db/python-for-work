"""YAML config loading and path resolution."""

from pathlib import Path

import yaml


def load_config(config_path: str | Path) -> dict:
    """Load configuration from a YAML file.

    Returns an empty dict when the file is missing (utilities may ship code
    defaults) or when the file is empty.
    """
    config_file = Path(config_path)
    if not config_file.exists():
        return {}
    with config_file.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def resolve_path(base_dir: str | Path, value: str | Path) -> Path:
    """Resolve ``value`` relative to ``base_dir``; leave absolute paths as-is."""
    path = Path(value).expanduser()
    if path.is_absolute():
        return path
    return Path(base_dir) / path


def resolve_config_path(value: str | Path, base_dir: str | Path) -> Path:
    """Find a config file: absolute path, then cwd, then under ``base_dir``."""
    path = Path(value).expanduser()
    if path.is_absolute():
        return path
    if path.exists():
        return path
    return Path(base_dir) / path
