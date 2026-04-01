from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime


@dataclass(slots=True)
class DocumentData:
    """Representa un registro del Excel listo para poblar una plantilla."""

    fields: dict[str, str] = field(default_factory=dict)
    row_number: int = 0

    @classmethod
    def from_raw(cls, raw_fields: dict[str, object], row_number: int) -> "DocumentData":
        normalized: dict[str, str] = {}
        for key, value in raw_fields.items():
            clean_key = str(key).strip().lower()
            normalized[clean_key] = cls._format_value(value)
        return cls(fields=normalized, row_number=row_number)

    @staticmethod
    def _format_value(value: object) -> str:
        if value is None:
            return ""
        if isinstance(value, (datetime, date)):
            return value.strftime("%d/%m/%Y")
        if hasattr(value, "strftime"):
            try:
                return value.strftime("%d/%m/%Y")
            except Exception:
                pass
        return str(value).strip()

    def value(self, key: str, default: str = "") -> str:
        return self.fields.get(key.strip().lower(), default)

    def numeric_value(self, key: str, default: float = 0.0) -> float:
        raw = self.value(key)
        if not raw:
            return default
        normalized = raw.replace("S/", "").replace(" ", "")
        # Soporta formato 1.234,56 o 1234.56
        if "," in normalized and "." in normalized:
            normalized = normalized.replace(".", "").replace(",", ".")
        elif "," in normalized:
            normalized = normalized.replace(",", ".")
        try:
            return float(normalized)
        except ValueError:
            return default

    def preview_values(self) -> dict[str, str]:
        return {
            "expediente": self.value("n° siged") or self.value("n_siged") or self.value("osinumero"),
            "fecha": self.value("fecha_not_req") or self.value("sofecha"),
            "administrado": self.value("razon_social") or self.value("razon") or "-",
        }

    def total_amount(self) -> float:
        return self.numeric_value("total")

    def to_placeholders(self) -> dict[str, str]:
        placeholders: dict[str, str] = {}
        for key, value in self.fields.items():
            placeholders[f"{{{{{key}}}}}"] = value
            placeholders[f"{{{key}}}"] = value
        return placeholders
