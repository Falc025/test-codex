from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from app.utils.text_utils import format_cell_value, normalize_column_name


@dataclass
class ExcelData:
    rows: list[dict[str, str]]
    columns: list[str]
    sheet_names: list[str]
    selected_sheet: str


class ExcelReader:
    def get_sheet_names(self, excel_path: str | Path) -> list[str]:
        xls = pd.ExcelFile(excel_path)
        return xls.sheet_names

    def read(self, excel_path: str | Path, sheet_name: str | None = None) -> ExcelData:
        xls = pd.ExcelFile(excel_path)
        selected = sheet_name or xls.sheet_names[0]
        df = pd.read_excel(excel_path, sheet_name=selected, engine="openpyxl")
        df = df.dropna(how="all")

        normalized_cols: list[str] = []
        for col in df.columns:
            normalized = normalize_column_name(col)
            normalized_cols.append(normalized)
        df.columns = normalized_cols

        rows: list[dict[str, str]] = []
        for _, rec in df.iterrows():
            row: dict[str, str] = {}
            for col in df.columns:
                val = rec.get(col)
                if pd.isna(val):
                    val = ""
                row[col] = format_cell_value(val)
            rows.append(row)

        return ExcelData(rows=rows, columns=list(df.columns), sheet_names=xls.sheet_names, selected_sheet=selected)
