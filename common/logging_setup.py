"""Konfigurasi logging konsisten untuk semua utility CLI."""

import logging
import sys


def setup_logging(level: int = logging.INFO) -> logging.Logger:
    """Siapkan logger ``pfw`` dengan output ke stdout.

    Format sengaja tanpa prefix agar console UX tetap sama seperti versi print.
    """
    logger = logging.getLogger("pfw")
    logger.setLevel(level)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stdout)
        handler.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(handler)
    logger.propagate = False
    return logger


def get_logger(name: str | None = None) -> logging.Logger:
    """Ambil logger anak dari ``pfw``."""
    if name:
        return logging.getLogger(f"pfw.{name}")
    return logging.getLogger("pfw")
