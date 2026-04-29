from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP


def format_number_es(value: float | int | Decimal, decimals: int = 2, use_thousands: bool = True) -> str:
    if value is None:
        return ""
    try:
        dec_value = Decimal(str(value))
    except Exception:
        return str(value).strip()
    quant = Decimal(f"1.{'0' * decimals}") if decimals > 0 else Decimal("1")
    quantized = dec_value.quantize(quant, rounding=ROUND_HALF_UP)
    sign = "-" if quantized < 0 else ""
    normalized = format(abs(quantized), f".{decimals}f")
    integer_part, decimal_part = normalized.split(".") if "." in normalized else (normalized, "")
    if use_thousands:
        groups = []
        for i in range(len(integer_part), 0, -3):
            groups.append(integer_part[max(0, i - 3):i])
        integer_part = " ".join(reversed(groups))
    return f"{sign}{integer_part},{decimal_part}" if decimals > 0 else f"{sign}{integer_part}"
