"""Normalize and format Indonesian phone numbers without dropping digits."""

import re

DEFAULT_MIN_DIGITS = 10
DEFAULT_MAX_DIGITS = 15


def phone_digits(value: object) -> str | None:
    """Return phone digits with a leading ``0``, without length checks.

    A ``62`` country prefix becomes ``0``; a bare ``8...`` number gets a leading
    ``0``. Returns ``None`` when there are no digits at all.
    """
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    digits = re.sub(r"\D", "", text)
    if not digits:
        return None

    if digits.startswith("62"):
        digits = "0" + digits[2:]
    elif digits.startswith("8"):
        digits = "0" + digits

    return digits


def normalize_phone(
    value: object,
    min_digits: int = DEFAULT_MIN_DIGITS,
    max_digits: int = DEFAULT_MAX_DIGITS,
) -> str | None:
    """Return valid phone digits (10-15 by default) or ``None``."""
    digits = phone_digits(value)
    if digits is None:
        return None
    if not min_digits <= len(digits) <= max_digits:
        return None
    return digits


def format_phone(digits: str) -> str:
    """Insert hyphens without truncating any digit.

    - 12 digits: ``0xxx-xxxx-xxxx``
    - 11 digits: ``0xxx-xxx-xxxx``
    - otherwise: groups of 4 from the left, remainder kept
    """
    if len(digits) == 12:
        return f"{digits[:4]}-{digits[4:8]}-{digits[8:]}"
    if len(digits) == 11:
        return f"{digits[:4]}-{digits[4:7]}-{digits[7:]}"
    return "-".join(digits[i : i + 4] for i in range(0, len(digits), 4))
