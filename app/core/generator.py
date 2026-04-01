from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from app.core.template_engine import TemplateEngine
from app.services.logging_service import LoggingService
from app.services.report_service import ReportService
from app.utils.file_utils import sanitize_filename, unique_path
from app.utils.text_utils import parse_numeric


@dataclass
class GenerationSummary:
    total_filas: int
    ok: int
    error: int
    advertencias: int
    reporte_path: Path
    log_path: Path


class DocumentGenerator:
    def __init__(self, template_engine: TemplateEngine) -> None:
        self.template_engine = template_engine

    def _select_template(self, apr_omitido: float, tpl_zero: Path, tpl_neg: Path, tpl_pos: Path) -> tuple[str, Path]:
        if apr_omitido == 0:
            return "cero", tpl_zero
        if apr_omitido < 0:
            return "negativo", tpl_neg
        return "positivo", tpl_pos

    def generate(
        self,
        rows: list[dict[str, str]],
        tpl_zero: str,
        tpl_neg: str,
        tpl_pos: str,
        output_dir: str,
        placeholders_by_tpl: dict[str, set[str]],
        logger: LoggingService,
        progress_cb,
        log_cb,
    ) -> GenerationSummary:
        out = Path(output_dir)
        report = ReportService(out)
        results: list[dict[str, str]] = []
        ok = 0
        err = 0
        warns = 0

        for idx, row in enumerate(rows, start=1):
            ruc = row.get("ruc", "")
            razon = row.get("razon_social", "")
            periodo = row.get("periodo", "")
            try:
                apr_value = parse_numeric(row.get("apr_omitido", ""))
                tpl_key, tpl_path = self._select_template(apr_value, Path(tpl_zero), Path(tpl_neg), Path(tpl_pos))

                template_placeholders = placeholders_by_tpl.get(tpl_key, set())
                data = {ph: row.get(ph, "") for ph in template_placeholders}
                for ph in template_placeholders:
                    if ph not in row:
                        warns += 1
                        msg = f"Fila {idx}: columna faltante para placeholder '{ph}', se reemplaza con vacío"
                        logger.warning(msg)
                        log_cb(msg)

                filename = f"{tpl_key}_{sanitize_filename(ruc, 'sinruc')}_{sanitize_filename(razon, 'sinrazon')}_{sanitize_filename(periodo, 'sinperiodo')}.docx"
                output_path = unique_path(out / filename)
                self.template_engine.render(tpl_path, data, output_path)
                ok += 1

                status = "OK"
                detail = ""
                generated = str(output_path)
                logger.info(f"Fila {idx} generada correctamente: {generated}")
                log_cb(f"OK fila {idx}: {generated}")
            except Exception as exc:
                err += 1
                status = "ERROR"
                detail = str(exc)
                generated = ""
                tpl_key = ""
                logger.error(f"Fila {idx} con error: {exc}")
                log_cb(f"ERROR fila {idx}: {exc}")

            results.append(
                {
                    "fila": str(idx),
                    "ruc": ruc,
                    "razon_social": razon,
                    "periodo": periodo,
                    "plantilla_usada": tpl_key,
                    "archivo_generado": generated,
                    "estado": status,
                    "detalle_error": detail,
                }
            )
            progress_cb(idx, len(rows))

        report_path = report.save(results)
        return GenerationSummary(
            total_filas=len(rows),
            ok=ok,
            error=err,
            advertencias=warns,
            reporte_path=report_path,
            log_path=logger.log_file,
        )
