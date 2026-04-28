from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from app.core.document_registry import MODULES
from app.utils.file_utils import ensure_directory


@dataclass
class ValidationResult:
    ok: bool
    messages: list[str]


class Validator:
    def validate_master_paths(self, excel_path: str, template_paths: dict[str, str], output_dir: str) -> ValidationResult:
        errors: list[str] = []
        if not Path(excel_path).is_file():
            errors.append("Excel maestro no existe.")

        for key, path in template_paths.items():
            p = Path(path)
            if not p.is_file() or p.suffix.lower() != ".docx":
                errors.append(f"Plantilla inválida o inexistente: {key}")

        try:
            ensure_directory(output_dir)
            for module_key in MODULES:
                ensure_directory(Path(output_dir) / module_key)
        except Exception as exc:
            errors.append(f"No se pudo crear carpeta de salida: {exc}")

        return ValidationResult(ok=not errors, messages=errors)

    def validate_sheets(self, available_sheets: list[str]) -> ValidationResult:
        missing = [item.sheet_name for item in MODULES.values() if item.sheet_name not in available_sheets]
        if missing:
            return ValidationResult(False, [f"Hojas faltantes en Excel maestro: {', '.join(missing)}"])
        return ValidationResult(True, [])

    def validate_placeholders_vs_columns(self, placeholders: set[str], columns: list[str], template_key: str) -> list[str]:
        cols = set(columns)
        absent = sorted([p for p in placeholders if p not in cols])
        if not absent:
            return []
        return [f"{template_key}: placeholders sin columna -> {', '.join(absent)}"]
