from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import re
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format
from openpyxl.worksheet.worksheet import Worksheet

from app.utils.text_utils import normalize_column_name


@dataclass
class ExcelRow:
    raw: dict[str, Any]
    display: dict[str, str]
    source_row: int


@dataclass
class ExcelData:
    rows: list[ExcelRow]
    columns: list[str]
    sheet_names: list[str]
    selected_sheet: str


class ExcelReader:
    MONTH_ABBR_ES = {
        1: "Ene",
        2: "Feb",
        3: "Mar",
        4: "Abr",
        5: "May",
        6: "Jun",
        7: "Jul",
        8: "Ago",
        9: "Sep",
        10: "Oct",
        11: "Nov",
        12: "Dic",
    }
    MONTH_NAME_ES = {
        1: "Enero",
        2: "Febrero",
        3: "Marzo",
        4: "Abril",
        5: "Mayo",
        6: "Junio",
        7: "Julio",
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre",
    }

    def get_sheet_names(self, excel_path: str | Path) -> list[str]:
        wb = load_workbook(excel_path, data_only=True)
        return wb.sheetnames

    def read(self, excel_path: str | Path, sheet_name: str | None = None) -> ExcelData:
        wb = load_workbook(excel_path, data_only=True)
        selected = sheet_name or wb.sheetnames[0]
        ws = wb[selected]

        header_row = next(ws.iter_rows(min_row=1, max_row=1))
        headers: list[str] = [normalize_column_name(cell.value) for cell in header_row]

        rows: list[ExcelRow] = []
        for row_idx, row_cells in enumerate(ws.iter_rows(min_row=2), start=2):
            row_raw: dict[str, Any] = {}
            row_display: dict[str, str] = {}
            has_content = False

            for i, cell in enumerate(row_cells):
                if i >= len(headers):
                    continue
                col = headers[i]
                if not col:
                    continue
                raw_val = cell.value
                disp_val = self.excel_display_value(cell)
                row_raw[col] = raw_val
                row_display[col] = disp_val
                if disp_val != "" or raw_val is not None:
                    has_content = True

            if has_content:
                rows.append(ExcelRow(raw=row_raw, display=row_display, source_row=row_idx))

        return ExcelData(rows=rows, columns=[h for h in headers if h], sheet_names=wb.sheetnames, selected_sheet=selected)

    def excel_display_value(self, cell) -> str:
        raw = cell.value
        if raw is None:
            return ""

        if isinstance(raw, str):
            return raw.strip()

        if isinstance(raw, bool):
            return "TRUE" if raw else "FALSE"

        fmt = (cell.number_format or "").strip()

        if isinstance(raw, (datetime, date)) or (is_date_format(fmt) and isinstance(raw, (int, float))):
            dt = raw if isinstance(raw, (datetime, date)) else self._excel_serial_to_datetime(raw)
            return self.format_excel_date(dt, fmt)

        if isinstance(raw, (int, float, Decimal)):
            return self._format_number(float(raw), fmt)

        return str(raw).strip()

    def _format_number(self, value: float, number_format: str) -> str:
        fmt = (number_format or "General").split(";")[0]

        if "%" in fmt:
            decimals = self._count_decimals(fmt)
            scaled = value * 100
            return f"{scaled:.{decimals}f}%"

        decimals = self._count_decimals(fmt)
        use_thousands = "," in fmt.split(".")[0]

        quantized = Decimal(str(value)).quantize(Decimal(f"1.{'0' * decimals}"), rounding=ROUND_HALF_UP) if decimals > 0 else Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)

        if decimals > 0:
            py_fmt = f",.{decimals}f" if use_thousands else f".{decimals}f"
            return format(float(quantized), py_fmt)

        py_fmt = ",.0f" if use_thousands else ".0f"
        return format(float(quantized), py_fmt)

    @staticmethod
    def _count_decimals(fmt: str) -> int:
        if "." not in fmt:
            return 0
        decimal_part = fmt.split(".", 1)[1]
        decimal_part = decimal_part.split("%", 1)[0]
        decimal_part = "".join(ch for ch in decimal_part if ch in {"0", "#"})
        return len(decimal_part)

    @staticmethod
    def _excel_serial_to_datetime(value: float) -> datetime:
        # Epoch 1899-12-30 for Excel serial dates
        base = datetime(1899, 12, 30)
        return base + timedelta(days=float(value))

    def format_excel_date(self, value: date | datetime, number_format: str) -> str:
        """
        Format a date/datetime value according to common Excel date display patterns.
        Falls back to dd/mm/yyyy for unknown date formats.
        """
        dt = value if isinstance(value, datetime) else datetime.combine(value, datetime.min.time())
        fmt = self._clean_excel_date_format(number_format)
        lower_fmt = fmt.lower()

        # Supported explicit patterns requested for this project.
        if re.fullmatch(r"d{1,2}/m{1,2}/y{4}", lower_fmt):
            day = str(dt.day) if lower_fmt.startswith("d/") else f"{dt.day:02d}"
            month = str(dt.month) if "/m/" in lower_fmt else f"{dt.month:02d}"
            return f"{day}/{month}/{dt.year:04d}"
        if re.fullmatch(r"mm/yyyy", lower_fmt):
            return f"{dt.month:02d}/{dt.year:04d}"
        if re.fullmatch(r"mmm-yy", lower_fmt):
            return f"{self._month_abbr(dt.month)}-{dt.year % 100:02d}"
        if re.fullmatch(r"mmmm-yy", lower_fmt):
            return f"{self._month_name(dt.month)}-{dt.year % 100:02d}"
        if re.fullmatch(r"mmm/yyyy", lower_fmt):
            return f"{self._month_abbr(dt.month)}/{dt.year:04d}"
        if re.fullmatch(r"mm-yy", lower_fmt):
            return f"{dt.month:02d}-{dt.year % 100:02d}"

        # Extra common variants
        if re.fullmatch(r"dd/mm/yyyy", lower_fmt):
            return f"{dt.day:02d}/{dt.month:02d}/{dt.year:04d}"
        if re.fullmatch(r"d/m/yyyy", lower_fmt):
            return f"{dt.day}/{dt.month}/{dt.year:04d}"

        # Fallback: preserve readable date in a stable default.
        return dt.strftime("%d/%m/%Y")

    @staticmethod
    def _clean_excel_date_format(number_format: str) -> str:
        """
        Keep the primary section and remove Excel-specific directives
        (colors/conditions/escaped chars/literals) to simplify matching.
        """
        base = (number_format or "").split(";")[0].strip()
        base = re.sub(r"\[[^\]]*\]", "", base)  # [Red], [$-es-PE], [>=1], etc.
        base = base.replace("\\", "")
        base = re.sub(r'"[^"]*"', "", base)
        return base.strip()

    @staticmethod
    def _month_abbr(month: int) -> str:
        return ExcelReader.MONTH_ABBR_ES[month]

    @staticmethod
    def _month_name(month: int) -> str:
        return ExcelReader.MONTH_NAME_ES[month]
