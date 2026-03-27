from __future__ import annotations

from pathlib import Path
from typing import Iterable


class ValidationError(Exception):
    """Error de validación de datos de entrada."""


class InputValidator:
    REQUIRED_SHEET = "Datos"
    REQUIRED_FIELDS = ("expediente", "fecha", "administrado")

    @staticmethod
    def validate_excel_path(path: str | Path) -> Path:
        excel_path = Path(path)
        if not excel_path.exists() or not excel_path.is_file():
            raise ValidationError("El archivo Excel no existe o no es accesible.")
        if excel_path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValidationError("El archivo Excel debe ser .xlsx o .xlsm.")
        return excel_path

    @staticmethod
    def validate_template_path(path: str | Path) -> Path:
        template_path = Path(path)
        if not template_path.exists() or not template_path.is_file():
            raise ValidationError("La plantilla Word no existe o no es accesible.")
        if template_path.suffix.lower() != ".docx":
            raise ValidationError("La plantilla debe ser un archivo .docx.")
        return template_path

    @staticmethod
    def validate_output_dir(path: str | Path) -> Path:
        output_dir = Path(path)
        if not output_dir.exists() or not output_dir.is_dir():
            raise ValidationError("La carpeta de salida no existe.")
        return output_dir

    @classmethod
    def validate_sheet_exists(cls, sheet_names: Iterable[str]) -> None:
        if cls.REQUIRED_SHEET not in set(sheet_names):
            raise ValidationError(f"No se encontró la hoja requerida '{cls.REQUIRED_SHEET}'.")

    @classmethod
    def validate_required_fields(cls, data: dict[str, object]) -> None:
        missing = [field for field in cls.REQUIRED_FIELDS if not str(data.get(field, "")).strip()]
        if missing:
            raise ValidationError(
                "Faltan campos obligatorios en el Excel: " + ", ".join(missing)
            )
