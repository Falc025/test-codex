from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime
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
