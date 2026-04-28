from __future__ import annotations

from pathlib import Path

import pandas as pd


class ReportService:
    COLUMNS = [
        "tipo_documental",
        "concepto",
        "hoja_excel",
        "fila",
        "ruc",
        "razon_social",
        "periodo",
        "plantilla_usada",
        "archivo_generado",
        "estado",
        "detalle_error",
    ]

    def __init__(self, output_dir: Path) -> None:
        self.output_dir = output_dir
        self.report_path = output_dir / "reporte_generacion.xlsx"

    def save(self, rows: list[dict[str, str]]) -> Path:
        normalized = [{col: row.get(col, "") for col in self.COLUMNS} for row in rows]
        df = pd.DataFrame(normalized, columns=self.COLUMNS)
        df.to_excel(self.report_path, index=False)
        return self.report_path
