"""Normalisasi dan format nomor telepon Indonesia tanpa kehilangan digit."""

import re

DEFAULT_MIN_DIGITS = 10
DEFAULT_MAX_DIGITS = 15


def phone_digits(value: object) -> str | None:
    """Kembalikan digit telepon dengan awalan ``0`` tanpa cek panjang.

    Awalan ``62`` diganti ``0`` dan awalan ``8`` (tanpa 0/62) ditambah ``0``.
    Mengembalikan ``None`` jika tidak ada digit sama sekali.
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
    """Kembalikan digit telepon yang valid (10-15 digit) atau ``None``."""
    digits = phone_digits(value)
    if digits is None:
        return None
    if not min_digits <= len(digits) <= max_digits:
        return None
    return digits


def format_phone(digits: str) -> str:
    """Beri tanda hubung tanpa memotong digit apa pun.

    - 12 digit: ``0xxx-xxxx-xxxx``
    - 11 digit: ``0xxx-xxx-xxxx``
    - lainnya: grup 4 digit dari kiri, sisa digit tetap disertakan
    """
    if len(digits) == 12:
        return f"{digits[:4]}-{digits[4:8]}-{digits[8:]}"
    if len(digits) == 11:
        return f"{digits[:4]}-{digits[4:7]}-{digits[7:]}"
    return "-".join(digits[i : i + 4] for i in range(0, len(digits), 4))
