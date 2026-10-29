"""Shared logging setup for every pfw CLI utility."""

import logging
import sys


def setup_logging(level: int = logging.INFO) -> logging.Logger:
    """Configure the ``pfw`` logger to write plain messages to stdout.

    The formatter intentionally omits prefixes so console UX matches the older
    print-based tools.
    """
    logger = logging.getLogger("pfw")
    logger.setLevel(level)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stdout)
        handler.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(handler)
    return logger


def get_logger(name: str | None = None) -> logging.Logger:
    """Return a child logger under ``pfw``."""
    if name:
        return logging.getLogger(f"pfw.{name}")
    return logging.getLogger("pfw")
