from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from app.core.document_registry import MODULES, TEMPLATE_LABELS
from app.core.document_selector import DocumentSelector
from app.core.excel_reader import ExcelRow
from app.core.template_engine import TemplateEngine
from app.services.logging_service import LoggingService
from app.utils.file_utils import sanitize_filename, unique_path


@dataclass
class GenerationSummary:
    total_filas: int
    ok: int
    error: int
    advertencias: int


class DocumentGenerator:
    def __init__(self, template_engine: TemplateEngine, selector: DocumentSelector) -> None:
        self.template_engine = template_engine
        self.selector = selector

    def generate_module(
        self,
        module_key: str,
        rows: list[ExcelRow],
        template_paths: dict[str, str],
        output_dir: str,
        placeholders_by_template: dict[str, set[str]],
        logger: LoggingService,
        progress_cb,
        log_cb,
    ) -> tuple[GenerationSummary, list[dict[str, str]]]:
        if module_key not in MODULES:
            raise ValueError(f"Módulo no soportado: {module_key}")

        module = MODULES[module_key]
        out = Path(output_dir) / module_key
        out.mkdir(parents=True, exist_ok=True)

        ok = 0
        err = 0
        warns = 0
        report_rows: list[dict[str, str]] = []

        for idx, row_obj in enumerate(rows, start=1):
            raw = row_obj.raw
            display = row_obj.display
            ruc = display.get("ruc", "")
            razon = display.get("razon_social", "")
            periodo = display.get("periodo", "")
            detail = ""
            generated = ""

            try:
                selection = self.selector.select(module_key, raw, display, template_paths)
                template_placeholders = placeholders_by_template.get(selection.template_key, set())
                data_display = {ph: display.get(ph, "") for ph in template_placeholders}

                for ph in template_placeholders:
                    if ph not in display:
                        warns += 1
                        msg = f"{module_key} fila {row_obj.source_row}: placeholder '{ph}' sin columna, se usa vacío"
                        logger.warning(msg)
                        log_cb(msg)

                filename = (
                    f"{module_key}_{selection.template_key}_"
                    f"{sanitize_filename(ruc, 'sinruc')}_"
                    f"{sanitize_filename(razon, 'sinrazon')}_"
                    f"{sanitize_filename(periodo, 'sinperiodo')}.docx"
                )
                output_path = unique_path(out / filename)
                self.template_engine.render(selection.template_path, data_display, output_path)

                ok += 1
                status = "OK"
                generated = str(output_path)
                logger.info(f"{module_key} fila Excel {row_obj.source_row} generada: {generated}")
                log_cb(f"OK {module_key} fila {row_obj.source_row}: {generated}")
                template_used = TEMPLATE_LABELS.get(selection.template_key, selection.template_key)
            except Exception as exc:
                err += 1
                status = "ERROR"
                detail = str(exc)
                template_used = ""
                logger.error(f"{module_key} fila Excel {row_obj.source_row} con error: {exc}")
                log_cb(f"ERROR {module_key} fila {row_obj.source_row}: {exc}")

            report_rows.append(
                {
                    "tipo_documental": module.document_type,
                    "concepto": module.concept,
                    "hoja_excel": module.sheet_name,
                    "fila": str(row_obj.source_row),
                    "ruc": ruc,
                    "razon_social": razon,
                    "periodo": periodo,
                    "plantilla_usada": template_used,
                    "archivo_generado": generated,
                    "estado": status,
                    "detalle_error": detail,
                }
            )

            progress_cb(idx, len(rows))

        return (
            GenerationSummary(total_filas=len(rows), ok=ok, error=err, advertencias=warns),
            report_rows,
        )
