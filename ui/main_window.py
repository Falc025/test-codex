from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from models.document_data import DocumentData
from services.document_builder import DocumentBuilder
from services.excel_reader import ExcelReader
from services.validator import ValidationError
from utils.file_utils import resource_path


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Generador Documental Local")
        self.resize(840, 520)

        self.excel_reader = ExcelReader()
        self.document_builder = DocumentBuilder()
        self.current_data: DocumentData | None = None

        self.excel_input = QLineEdit()
        self.template_input = QLineEdit()
        self.output_input = QLineEdit()

        self.preview_expediente = QLabel("-")
        self.preview_fecha = QLabel("-")
        self.preview_administrado = QLabel("-")
        self.status_box = QTextEdit()
        self.status_box.setReadOnly(True)

        self._build_ui()
        self._set_defaults()

    def _build_ui(self) -> None:
        root = QWidget()
        main_layout = QVBoxLayout(root)
        main_layout.setSpacing(12)

        files_group = QGroupBox("Entradas y salida")
        files_layout = QFormLayout(files_group)

        files_layout.addRow("Archivo Excel:", self._line_with_button(self.excel_input, self._select_excel))
        files_layout.addRow(
            "Plantilla Word:",
            self._line_with_button(self.template_input, self._select_template),
        )
        files_layout.addRow(
            "Carpeta salida:",
            self._line_with_button(self.output_input, self._select_output_dir),
        )

        preview_group = QGroupBox("Previsualización de datos")
        preview_layout = QFormLayout(preview_group)
        preview_layout.addRow("Expediente:", self.preview_expediente)
        preview_layout.addRow("Fecha:", self.preview_fecha)
        preview_layout.addRow("Administrado:", self.preview_administrado)

        actions_layout = QHBoxLayout()
        btn_load = QPushButton("Cargar Excel")
        btn_load.clicked.connect(self._load_preview)

        btn_generate = QPushButton("Generar documento")
        btn_generate.clicked.connect(self._generate_document)

        actions_layout.addWidget(btn_load)
        actions_layout.addWidget(btn_generate)
        actions_layout.addStretch()

        status_group = QGroupBox("Mensajes")
        status_layout = QVBoxLayout(status_group)
        status_layout.addWidget(self.status_box)

        main_layout.addWidget(files_group)
        main_layout.addWidget(preview_group)
        main_layout.addLayout(actions_layout)
        main_layout.addWidget(status_group)

        self.setCentralWidget(root)

    def _line_with_button(self, line_edit: QLineEdit, callback) -> QWidget:
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)

        browse_btn = QPushButton("...")
        browse_btn.setFixedWidth(36)
        browse_btn.clicked.connect(callback)

        layout.addWidget(line_edit)
        layout.addWidget(browse_btn)
        return container

    def _set_defaults(self) -> None:
        default_template = resource_path("templates/plantilla_base.docx")
        default_output = resource_path("output")

        self.template_input.setText(str(default_template))
        self.output_input.setText(str(default_output))

    def _select_excel(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo Excel",
            "",
            "Excel (*.xlsx *.xlsm)",
        )
        if file_path:
            self.excel_input.setText(file_path)

    def _select_template(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar plantilla Word",
            "",
            "Word (*.docx)",
        )
        if file_path:
            self.template_input.setText(file_path)

    def _select_output_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta de salida")
        if folder:
            self.output_input.setText(folder)

    def _load_preview(self) -> None:
        try:
            data = self.excel_reader.read_document_data(self.excel_input.text())
        except ValidationError as exc:
            self._show_error(str(exc))
            return
        except Exception as exc:
            self._show_error(f"Error inesperado al leer Excel: {exc}")
            return

        self.current_data = data
        self.preview_expediente.setText(data.expediente)
        self.preview_fecha.setText(data.fecha)
        self.preview_administrado.setText(data.administrado)
        self._log("Datos cargados correctamente.")

    def _generate_document(self) -> None:
        if self.current_data is None:
            self._show_error("Primero cargue y valide datos desde el Excel.")
            return

        try:
            output_path = self.document_builder.build(
                template_path=self.template_input.text(),
                output_dir=self.output_input.text(),
                data=self.current_data,
            )
        except ValidationError as exc:
            self._show_error(str(exc))
            return
        except Exception as exc:
            self._show_error(f"Error inesperado al generar documento: {exc}")
            return

        self._log(f"Documento generado con éxito: {output_path}")
        QMessageBox.information(self, "Éxito", f"Documento generado:\n{output_path}")

    def _show_error(self, message: str) -> None:
        self._log(f"ERROR: {message}")
        QMessageBox.critical(self, "Error", message)

    def _log(self, message: str) -> None:
        self.status_box.append(message)
        self.status_box.moveCursor(self.status_box.textCursor().End)

    def keyPressEvent(self, event) -> None:  # noqa: N802
        if event.key() == Qt.Key_F5:
            self._load_preview()
            return
        super().keyPressEvent(event)
