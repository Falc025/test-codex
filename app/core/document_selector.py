from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from app.utils.text_utils import parse_numeric


@dataclass
class SelectionResult:
    template_key: str
    template_path: Path


class DocumentSelector:
    def select(self, module_key: str, raw: dict[str, Any], display: dict[str, str], template_paths: dict[str, str]) -> SelectionResult:
        if module_key == "RD_APR":
            value = self._numeric_from_row(raw, display, ["apr_omitido"])
            key = "rd_apr_cero" if value == 0 else ("rd_apr_negativo" if value < 0 else "rd_apr_positivo")
            return SelectionResult(key, Path(template_paths[key]))
        if module_key == "RD_DGE":
            value = self._numeric_from_row(raw, display, ["dge_omitido", "total"])
            key = "rd_dge_cero" if value == 0 else ("rd_dge_negativo" if value < 0 else "rd_dge_positivo")
            return SelectionResult(key, Path(template_paths[key]))
        if module_key == "RM_APR":
            key = self._rm_key(raw, display, "apr")
            return SelectionResult(key, Path(template_paths[key]))
        if module_key == "RM_DGE":
            key = self._rm_key(raw, display, "dge")
            return SelectionResult(key, Path(template_paths[key]))
        raise ValueError(f"Módulo no soportado: {module_key}")

    def _numeric_from_row(self, raw: dict[str, Any], display: dict[str, str], fields: list[str]) -> float:
        for field in fields:
            source = raw.get(field)
            if source is None:
                source = display.get(field, "")
            if source in (None, ""):
                continue
            try:
                return parse_numeric(source)
            except Exception:
                continue
        raise ValueError(f"No se pudo resolver valor numérico para: {', '.join(fields)}")

    def _rm_key(self, raw: dict[str, Any], display: dict[str, str], concept: str) -> str:
        token = ""
        for c in ["tipo_plantilla", "tipo_multa", "infraccion"]:
            val = raw.get(c)
            if val in (None, ""):
                val = display.get(c, "")
            if val not in (None, ""):
                token = str(val).lower().strip()
                break
        normalized = token.replace(" ", "").replace("_", "").replace(".", "")
        if "176" in normalized and "1" in normalized:
            return f"rm_{concept}_176_1"
        if "178" in normalized and "1" in normalized:
            return f"rm_{concept}_178_1"
        raise ValueError("No se pudo determinar tipo de multa (176-1 o 178-1)")
