from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime, timedelta
from decimal import Decimal
from typing import Any

PLACEHOLDER_PATTERN = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}")


def normalize_column_name(name: str) -> str:
    text = str(name or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^a-z0-9_]", "", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def format_cell_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, float):
        txt = f"{value:.10f}".rstrip("0").rstrip(".")
        return txt
    if isinstance(value, Decimal):
        return format(value, "f")
    return str(value).strip()


def parse_numeric(value: Any) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    text = format_cell_value(value)
    if not text:
        raise ValueError("Valor vacío")

    cleaned = text.replace(" ", "")
    cleaned = cleaned.replace("S/", "").replace("$", "")
    cleaned = re.sub(r"[^0-9,.-]", "", cleaned)

    if cleaned.count(",") > 0 and cleaned.count(".") > 0:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        cleaned = cleaned.replace(",", ".")

    return float(cleaned)


def extract_placeholders(text: str) -> set[str]:
    return {match.group(1).strip() for match in PLACEHOLDER_PATTERN.finditer(text or "")}


def parse_period_sort_value(value: Any) -> datetime | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, (int, float)):
        # Excel serial date
        try:
            base = datetime(1899, 12, 30)
            return base + timedelta(days=float(value))
        except Exception:
            return None

    text = str(value).strip()
    if not text:
        return None
    text_low = text.lower()

    month_map = {
        "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
        "jul": 7, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12,
    }

    m = re.fullmatch(r"([a-záéíóú]{3})[-/](\d{2,4})", text_low)
    if m:
        mon = month_map.get(m.group(1)[:3])
        year = int(m.group(2))
        year = 2000 + year if year < 100 else year
        if mon:
            return datetime(year, mon, 1)

    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            dt = datetime.strptime(text, fmt)
            if fmt == "%m/%Y":
                dt = datetime(dt.year, dt.month, 1)
            return dt
        except ValueError:
            continue
    return None
