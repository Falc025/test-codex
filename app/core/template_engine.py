from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document

from app.utils.text_utils import extract_placeholders


@dataclass
class TemplateScan:
    placeholders: set[str]


class TemplateEngine:
    _token_pattern = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}")
    _fragment_pattern = re.compile(r"\{\{(?:[^{}]|<[^>]+>)+\}\}")

    def scan_placeholders(self, docx_path: str | Path) -> TemplateScan:
        doc = Document(docx_path)
        found: set[str] = set()
        for paragraph in self._iter_all_paragraphs(doc):
            found.update(extract_placeholders(paragraph.text))
        return TemplateScan(placeholders=found)

    def render(self, template_path: str | Path, data: dict[str, str], output_path: str | Path) -> None:
        doc = Document(template_path)
        self.replace_placeholders_in_document(doc, data)
        doc.save(output_path)
        self.replace_placeholders_in_docx_xml(output_path, data)

    def replace_placeholders_in_document(self, doc: Document, values: dict[str, str]) -> None:
        for container in self._iter_all_containers(doc):
            self.replace_placeholders_in_container(container, values)

    def replace_placeholders_in_container(self, container: Any, values: dict[str, str]) -> None:
        for paragraph in container.paragraphs:
            self._replace_in_paragraph(paragraph, values)
        for table in container.tables:
            self._replace_in_table(table, values)

    def _replace_in_table(self, table: Any, values: dict[str, str]) -> None:
        for row in table.rows:
            for cell in row.cells:
                self.replace_placeholders_in_container(cell, values)

    def _iter_all_containers(self, doc: Document):
        # Documento principal (body)
        yield doc
        # Todas las variantes de header/footer por sección.
        visited: set[int] = set()
        for section in doc.sections:
            containers = [
                section.header,
                section.footer,
                section.first_page_header,
                section.first_page_footer,
                section.even_page_header,
                section.even_page_footer,
            ]
            for container in containers:
                # En Word pueden estar vinculados entre secciones; evitar reprocesar el mismo objeto.
                if id(container) in visited:
                    continue
                visited.add(id(container))
                yield container

    def _iter_all_paragraphs(self, doc: Document):
        for container in self._iter_all_containers(doc):
            for p in container.paragraphs:
                yield p
            for table in container.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            yield p

    def _replace_in_paragraph(self, paragraph: Any, data: dict[str, str]) -> None:
        if not paragraph.runs:
            return
        original_text = "".join(run.text for run in paragraph.runs)
        replaced_text = self._token_pattern.sub(lambda m: data.get(m.group(1).strip(), ""), original_text)
        if replaced_text == original_text:
            return

        lengths = [len(run.text) for run in paragraph.runs]
        cursor = 0
        for i, run in enumerate(paragraph.runs):
            take = lengths[i]
            segment = replaced_text[cursor: cursor + take]
            run.text = segment
            cursor += len(segment)
        if cursor < len(replaced_text):
            paragraph.runs[-1].text += replaced_text[cursor:]

    def replace_placeholders_in_docx_xml(self, docx_path: str | Path, values: dict[str, str]) -> None:
        src = Path(docx_path)
        temp = src.with_suffix(".tmp.docx")
        with ZipFile(src, "r") as zin, ZipFile(temp, "w", compression=ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                payload = zin.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    xml = payload.decode("utf-8")
                    replaced = self._replace_placeholders_in_xml_text(xml, values)
                    payload = replaced.encode("utf-8")
                zout.writestr(item, payload)
        temp.replace(src)

    def find_remaining_placeholders(self, docx_path: str | Path) -> dict[str, list[str]]:
        found: dict[str, list[str]] = {}
        with ZipFile(docx_path, "r") as zf:
            for name in zf.namelist():
                if not (name.startswith("word/") and name.endswith(".xml")):
                    continue
                xml = zf.read(name).decode("utf-8")
                stripped = re.sub(r"<[^>]+>", "", xml)
                placeholders = sorted({m.group(1) for m in self._token_pattern.finditer(stripped)})
                if placeholders:
                    found[name] = [f"{{{{{p}}}}}" for p in placeholders]
        return found

    def _replace_placeholders_in_xml_text(self, xml_text: str, values: dict[str, str]) -> str:
        # 1) Reemplazo estándar.
        replaced = self._token_pattern.sub(lambda m: str(values.get(m.group(1).strip(), "")), xml_text)

        # 2) Respaldo para placeholders fragmentados entre nodos XML.
        def repl_fragment(match: re.Match) -> str:
            token_with_tags = match.group(0)
            plain = re.sub(r"<[^>]+>", "", token_with_tags)
            inner = self._token_pattern.fullmatch(plain)
            if not inner:
                return token_with_tags
            key = inner.group(1).strip()
            return str(values.get(key, ""))

        return self._fragment_pattern.sub(repl_fragment, replaced)
