from __future__ import annotations

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
        self.resize(980, 620)

        self.excel_reader = ExcelReader()
        self.document_builder = DocumentBuilder()
        self.current_records: list[DocumentData] = []

        self.excel_input = QLineEdit()
        self.template_zero_input = QLineEdit()
        self.template_positive_input = QLineEdit()
        self.template_negative_input = QLineEdit()
        self.output_input = QLineEdit()

        self.preview_count = QLabel("0")
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
            "Plantilla total = 0:",
            self._line_with_button(self.template_zero_input, lambda: self._select_template(self.template_zero_input)),
        )
        files_layout.addRow(
            "Plantilla total > 0:",
            self._line_with_button(self.template_positive_input, lambda: self._select_template(self.template_positive_input)),
        )
        files_layout.addRow(
            "Plantilla total < 0:",
            self._line_with_button(self.template_negative_input, lambda: self._select_template(self.template_negative_input)),
        )
        files_layout.addRow(
            "Carpeta salida:",
            self._line_with_button(self.output_input, self._select_output_dir),
        )

        preview_group = QGroupBox("Previsualización")
        preview_layout = QFormLayout(preview_group)
        preview_layout.addRow("Registros detectados:", self.preview_count)
        preview_layout.addRow("Primer expediente/SIGED:", self.preview_expediente)
        preview_layout.addRow("Primera fecha:", self.preview_fecha)
        preview_layout.addRow("Primer administrado:", self.preview_administrado)

        actions_layout = QHBoxLayout()
        btn_load = QPushButton("Cargar Excel")
        btn_load.clicked.connect(self._load_preview)

        btn_generate = QPushButton("Generar documentos")
        btn_generate.clicked.connect(self._generate_documents)

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
        self.template_zero_input.setText(str(resource_path("templates/plantilla_total_cero.docx")))
        self.template_positive_input.setText(str(resource_path("templates/plantilla_total_positivo.docx")))
        self.template_negative_input.setText(str(resource_path("templates/plantilla_total_negativo.docx")))
        self.output_input.setText(str(resource_path("output")))

    def _select_excel(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo Excel",
            "",
            "Excel (*.xlsx *.xlsm)",
        )
        if file_path:
            self.excel_input.setText(file_path)

    def _select_template(self, target_input: QLineEdit) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar plantilla Word",
            "",
            "Word (*.docx)",
        )
        if file_path:
            target_input.setText(file_path)

    def _select_output_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta de salida")
        if folder:
            self.output_input.setText(folder)

    def _load_preview(self) -> None:
        try:
            records = self.excel_reader.read_document_data(self.excel_input.text())
        except ValidationError as exc:
            self._show_error(str(exc))
            return
        except Exception as exc:
            self._show_error(f"Error inesperado al leer Excel: {exc}")
            return

        self.current_records = records
        first = records[0]
        preview = first.preview_values()

        self.preview_count.setText(str(len(records)))
        self.preview_expediente.setText(preview["expediente"] or "-")
        self.preview_fecha.setText(preview["fecha"] or "-")
        self.preview_administrado.setText(preview["administrado"] or "-")
        self._log(f"Excel cargado: {len(records)} registros listos para generar.")

    def _generate_documents(self) -> None:
        if not self.current_records:
            self._show_error("Primero cargue y valide el Excel.")
            return

        try:
            generated_paths = self.document_builder.build_many(
                records=self.current_records,
                template_zero=self.template_zero_input.text(),
                template_positive=self.template_positive_input.text(),
                template_negative=self.template_negative_input.text(),
                output_dir=self.output_input.text(),
            )
        except ValidationError as exc:
            self._show_error(str(exc))
            return
        except Exception as exc:
            self._show_error(f"Error inesperado al generar documentos: {exc}")
            return

        self._log(f"Generación completada: {len(generated_paths)} documentos creados.")
        self._log(f"Ejemplo de salida: {generated_paths[0]}")
        QMessageBox.information(
            self,
            "Éxito",
            f"Se generaron {len(generated_paths)} documentos en:\n{self.output_input.text()}",
        )

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
