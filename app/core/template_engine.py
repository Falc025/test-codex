from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document

from app.utils.text_utils import extract_placeholders


@dataclass
class TemplateScan:
    placeholders: set[str]


class TemplateEngine:
    _token_pattern = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}")
    _word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

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
                if self._should_process_xml_part(item.filename):
                    payload = self.replace_placeholders_in_xml_part(payload, values)
                zout.writestr(item, payload)
        temp.replace(src)

    def find_remaining_placeholders(self, docx_path: str | Path) -> dict[str, list[str]]:
        found: dict[str, list[str]] = {}
        with ZipFile(docx_path, "r") as zf:
            for name in zf.namelist():
                if not (name.startswith("word/") and name.endswith(".xml")):
                    continue
                xml = zf.read(name)
                placeholders = sorted(self._extract_placeholders_from_xml_part(xml))
                if placeholders:
                    found[name] = [f"{{{{{p}}}}}" for p in placeholders]
        return found

    def replace_placeholders_in_xml_part(self, xml_bytes: bytes, values: dict[str, str]) -> bytes:
        try:
            root = ET.fromstring(xml_bytes)
        except ET.ParseError:
            return xml_bytes

        changed = False
        paragraph_xpath = f".//{{{self._word_ns}}}p"
        text_xpath = f".//{{{self._word_ns}}}t"

        for paragraph in root.findall(paragraph_xpath):
            text_nodes = paragraph.findall(text_xpath)
            if not text_nodes:
                continue
            joined = "".join(node.text or "" for node in text_nodes)
            if "{{" not in joined:
                continue
            replaced = self._token_pattern.sub(lambda m: str(values.get(m.group(1).strip(), "")), joined)
            if replaced == joined:
                continue
            text_nodes[0].text = replaced
            for node in text_nodes[1:]:
                node.text = ""
            changed = True

        # Fallback sobre todo el XML (si no hubo reemplazo por párrafo) para casos especiales.
        if not changed:
            xml_text = xml_bytes.decode("utf-8")
            replaced_xml = self._token_pattern.sub(lambda m: str(values.get(m.group(1).strip(), "")), xml_text)
            if replaced_xml != xml_text:
                return replaced_xml.encode("utf-8")
            return xml_bytes

        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def _extract_placeholders_from_xml_part(self, xml_bytes: bytes) -> set[str]:
        try:
            root = ET.fromstring(xml_bytes)
            text_xpath = f".//{{{self._word_ns}}}t"
            text = "".join(node.text or "" for node in root.findall(text_xpath))
        except ET.ParseError:
            text = xml_bytes.decode("utf-8", errors="ignore")
        return {m.group(1).strip() for m in self._token_pattern.finditer(text)}

    def _should_process_xml_part(self, filename: str) -> bool:
        if not (filename.startswith("word/") and filename.endswith(".xml")):
            return False
        if filename in {"word/document.xml", "word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"}:
            return True
        return filename.startswith("word/header") or filename.startswith("word/footer")
