from __future__ import annotations

from dataclasses import dataclass
from datetime import date


@dataclass(slots=True)
class DocumentData:
    expediente: str
    fecha: str
    administrado: str

    @classmethod
    def from_raw(cls, expediente: str, fecha_value: object, administrado: str) -> "DocumentData":
        fecha = cls._format_fecha(fecha_value)
        return cls(
            expediente=str(expediente).strip(),
            fecha=fecha,
            administrado=str(administrado).strip(),
        )

    @staticmethod
    def _format_fecha(value: object) -> str:
        if value is None:
            return ""
        if hasattr(value, "strftime"):
            try:
                return value.strftime("%d/%m/%Y")
            except Exception:
                pass
        if isinstance(value, date):
            return value.strftime("%d/%m/%Y")
        return str(value).strip()

    def to_placeholders(self) -> dict[str, str]:
        return {
            "{{expediente}}": self.expediente,
            "{{fecha}}": self.fecha,
            "{{administrado}}": self.administrado,
        }
