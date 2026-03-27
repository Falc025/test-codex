from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from models.document_data import DocumentData
from services.validator import InputValidator, ValidationError


class ExcelReader:
    """Lee Excel estructurado y transforma información a modelo de dominio."""

    def __init__(self, validator: InputValidator | None = None) -> None:
        self.validator = validator or InputValidator()

    def read_document_data(self, excel_path: str | Path) -> DocumentData:
        validated_path = self.validator.validate_excel_path(excel_path)

        try:
            workbook = load_workbook(validated_path, data_only=True)
        except Exception as exc:
            raise ValidationError(f"No se pudo abrir el Excel: {exc}") from exc

        self.validator.validate_sheet_exists(workbook.sheetnames)
        sheet = workbook[InputValidator.REQUIRED_SHEET]

        # Formato esperado del MVP:
        # Hoja 'Datos' con pares clave/valor en columnas A:B.
        # A2='expediente', B2='EXP-001' ...
        raw_data: dict[str, object] = {}
        for row in sheet.iter_rows(min_row=1, max_col=2, values_only=True):
            key, value = row
            if not key:
                continue
            normalized_key = str(key).strip().lower()
            raw_data[normalized_key] = value

        self.validator.validate_required_fields(raw_data)

        return DocumentData.from_raw(
            expediente=raw_data["expediente"],
            fecha_value=raw_data["fecha"],
            administrado=raw_data["administrado"],
        )
