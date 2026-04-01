from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from app.utils.file_utils import ensure_directory


REQUIRED_COLUMNS = {
    "razon_social",
    "ruc",
    "domicilio",
    "sector",
    "periodo",
    "resul_req",
    "fecha_not_req",
    "fecha_venci",
    "bi_det",
    "alicuota",
    "apr_det",
    "apr_decla",
    "apr_pagado",
    "apr_omitido",
    "fecha_decla",
    "int_moratorio",
    "total",
    "fecha_calculo",
    "base_legal",
    "correo_especialista",
}


@dataclass
class ValidationResult:
    ok: bool
    messages: list[str]


class Validator:
    def validate_paths(
        self,
        excel_path: str,
        tpl_zero: str,
        tpl_neg: str,
        tpl_pos: str,
        output_dir: str,
    ) -> ValidationResult:
        errors: list[str] = []
        if not Path(excel_path).is_file():
            errors.append("Excel no existe.")
        for p, name in [(tpl_zero, "cero"), (tpl_neg, "negativo"), (tpl_pos, "positivo")]:
            path = Path(p)
            if not path.is_file() or path.suffix.lower() != ".docx":
                errors.append(f"Plantilla {name} inválida.")
        try:
            ensure_directory(output_dir)
        except Exception as exc:
            errors.append(f"No se pudo crear carpeta de salida: {exc}")
        return ValidationResult(ok=not errors, messages=errors)

    def validate_columns(self, columns: list[str]) -> tuple[list[str], list[str]]:
        available = set(columns)
        missing = sorted(list(REQUIRED_COLUMNS - available))
        return sorted(list(available)), missing
