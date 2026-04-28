from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from app.core.document_registry import MODULES, TEMPLATE_LABELS
from app.core.document_selector import DocumentSelector
from app.core.excel_reader import ExcelRow
from app.core.template_engine import TemplateEngine
from app.services.logging_service import LoggingService
from app.utils.file_utils import sanitize_filename, unique_path
from app.utils.text_utils import parse_period_sort_value


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

    def sort_records_by_period(self, rows: list[ExcelRow], logger: LoggingService, module_key: str) -> list[ExcelRow]:
        sortable: list[tuple[ExcelRow, object]] = []
        for row in rows:
            period_raw = row.raw.get("periodo")
            sort_dt = parse_period_sort_value(period_raw)
            if sort_dt is None:
                logger.warning(
                    f"{module_key} fila {row.source_row}: periodo no convertible para orden descendente, se enviará al final"
                )
                sort_key = datetime.min
            else:
                sort_key = sort_dt
            sortable.append((row, sort_key))
        sortable.sort(key=lambda item: item[1], reverse=True)
        return [item[0] for item in sortable]

    @staticmethod
    def build_output_filename(record_display: dict[str, str], sheet_name: str) -> str:
        razon = sanitize_filename(record_display.get("razon_social", "") or "SIN_RAZON")
        periodo = sanitize_filename(record_display.get("periodo", "") or "SIN_PERIODO")
        sector = sanitize_filename(record_display.get("sector", "") or "SIN_SECTOR")
        return f"{sheet_name}_{razon}_{periodo}_{sector}.docx"

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
        ordered_rows = self.sort_records_by_period(rows, logger, module_key)

        for idx, row_obj in enumerate(ordered_rows, start=1):
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
                data_display = dict(display)

                for ph in template_placeholders:
                    if ph not in display:
                        warns += 1
                        msg = f"{module_key} fila {row_obj.source_row}: placeholder '{ph}' sin columna, se usa vacío"
                        logger.warning(msg)
                        log_cb(msg)

                if not display.get("razon_social"):
                    warns += 1
                    logger.warning(f"{module_key} fila {row_obj.source_row}: razon_social vacío, se usará SIN_RAZON")
                if not display.get("periodo"):
                    warns += 1
                    logger.warning(f"{module_key} fila {row_obj.source_row}: periodo vacío, se usará SIN_PERIODO")
                filename = self.build_output_filename(display, module.sheet_name)
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

            progress_cb(idx, len(ordered_rows))

        return (
            GenerationSummary(total_filas=len(ordered_rows), ok=ok, error=err, advertencias=warns),
            report_rows,
        )
