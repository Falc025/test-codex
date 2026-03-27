from __future__ import annotations

from datetime import datetime
from pathlib import Path

from docx import Document

from models.document_data import DocumentData
from services.validator import InputValidator, ValidationError
from utils.file_utils import ensure_directory


class DocumentBuilder:
    """Construye documentos Word reemplazando marcadores de plantilla."""

    def __init__(self, validator: InputValidator | None = None) -> None:
        self.validator = validator or InputValidator()

    def build(
        self,
        template_path: str | Path,
        output_dir: str | Path,
        data: DocumentData,
        filename_prefix: str = "documento",
    ) -> Path:
        template = self.validator.validate_template_path(template_path)
        out_dir = ensure_directory(self.validator.validate_output_dir(output_dir))

        try:
            document = Document(template)
        except Exception as exc:
            raise ValidationError(f"No se pudo abrir la plantilla Word: {exc}") from exc

        placeholders = data.to_placeholders()
        self._replace_in_document(document, placeholders)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_expediente = data.expediente.replace("/", "-").replace(" ", "_")
        output_name = f"{filename_prefix}_{safe_expediente}_{timestamp}.docx"
        output_path = out_dir / output_name

        try:
            document.save(output_path)
        except Exception as exc:
            raise ValidationError(f"No se pudo guardar el documento generado: {exc}") from exc

        return output_path

    def _replace_in_document(self, document: Document, placeholders: dict[str, str]) -> None:
        for paragraph in document.paragraphs:
            self._replace_in_paragraph(paragraph, placeholders)

        # Preparado para tablas dinámicas futuras: también se procesan celdas.
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_in_paragraph(paragraph, placeholders)

    @staticmethod
    def _replace_in_paragraph(paragraph, placeholders: dict[str, str]) -> None:
        full_text = paragraph.text
        if not full_text:
            return
        for token, value in placeholders.items():
            full_text = full_text.replace(token, value)

        if paragraph.runs:
            paragraph.runs[0].text = full_text
            for run in paragraph.runs[1:]:
                run.text = ""
