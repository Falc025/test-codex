from __future__ import annotations

from pathlib import Path

from docx import Document
from openpyxl import Workbook

BASE_DIR = Path(__file__).resolve().parents[1]
TEMPLATE_PATH = BASE_DIR / "templates" / "plantilla_base.docx"
EXCEL_PATH = BASE_DIR / "sample_input.xlsx"


def create_excel() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Datos"
    ws["A1"] = "campo"
    ws["B1"] = "valor"
    ws["A2"] = "expediente"
    ws["B2"] = "EXP-2026-001"
    ws["A3"] = "fecha"
    ws["B3"] = "27/03/2026"
    ws["A4"] = "administrado"
    ws["B4"] = "Acme S.A.C."

    wb.save(EXCEL_PATH)


def create_template() -> None:
    doc = Document()
    doc.add_heading("Documento de ejemplo", level=1)
    doc.add_paragraph("Expediente: {{expediente}}")
    doc.add_paragraph("Fecha: {{fecha}}")
    doc.add_paragraph("Administrado: {{administrado}}")
    doc.add_paragraph("\nTexto adicional para validar el formato base.")
    doc.save(TEMPLATE_PATH)


def main() -> None:
    TEMPLATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    create_excel()
    create_template()
    print(f"Excel de ejemplo creado en: {EXCEL_PATH}")
    print(f"Plantilla de ejemplo creada en: {TEMPLATE_PATH}")


if __name__ == "__main__":
    main()
