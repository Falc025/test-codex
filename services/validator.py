from __future__ import annotations

from pathlib import Path
from typing import Iterable


class ValidationError(Exception):
    """Error de validación de datos de entrada."""


class InputValidator:
    REQUIRED_SHEET = "Datos"
    REQUIRED_COLUMNS = (
        "ruc",
        "razon_social",
        "domicilio",
        "periodo",
        "total",
    )

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
            raise ValidationError(f"La plantilla no existe: {template_path}")
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
    def validate_header_columns(cls, header_columns: list[str]) -> None:
        header_set = {col.strip().lower() for col in header_columns if col}
        missing = [col for col in cls.REQUIRED_COLUMNS if col not in header_set]
        if missing:
            raise ValidationError(
                "Faltan columnas obligatorias en la hoja 'Datos': " + ", ".join(missing)
            )

    @classmethod
    def validate_row_data(cls, data: dict[str, object], row_number: int) -> None:
        missing = [col for col in cls.REQUIRED_COLUMNS if not str(data.get(col, "")).strip()]
        if missing:
            raise ValidationError(
                f"Fila {row_number}: faltan campos obligatorios ({', '.join(missing)})."
            )

    @classmethod
    def validate_templates(cls, zero_path: str | Path, positive_path: str | Path, negative_path: str | Path) -> tuple[Path, Path, Path]:
        return (
            cls.validate_template_path(zero_path),
            cls.validate_template_path(positive_path),
            cls.validate_template_path(negative_path),
        )
