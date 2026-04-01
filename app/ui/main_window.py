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
    QComboBox,
    QProgressBar,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.core.excel_reader import ExcelReader, ExcelData
from app.core.generator import DocumentGenerator
from app.core.template_engine import TemplateEngine
from app.core.validator import Validator
from app.services.config_service import ConfigService
from app.services.logging_service import LoggingService
from app.utils.file_utils import open_directory
from app.utils.text_utils import parse_numeric


class MainWindow(QMainWindow):
    def __init__(self, config_service: ConfigService) -> None:
        super().__init__()
        self.setWindowTitle("Generador masivo DOCX")
        self.resize(1300, 840)

        self.config_service = config_service
        self.reader = ExcelReader()
        self.validator = Validator()
        self.template_engine = TemplateEngine()
        self.generator = DocumentGenerator(self.template_engine)

        self.excel_data: ExcelData | None = None
        self.placeholders_by_tpl: dict[str, set[str]] = {"cero": set(), "negativo": set(), "positivo": set()}

        self.excel_edit = QLineEdit()
        self.tpl_zero_edit = QLineEdit()
        self.tpl_neg_edit = QLineEdit()
        self.tpl_pos_edit = QLineEdit()
        self.output_edit = QLineEdit()
        self.sheet_combo = QComboBox()

        self.columns_text = QTextEdit(); self.columns_text.setReadOnly(True)
        self.placeholders_text = QTextEdit(); self.placeholders_text.setReadOnly(True)
        self.mapping_text = QTextEdit(); self.mapping_text.setReadOnly(True)

        self.preview_table = QTableWidget()
        self.preview_row_placeholders = QTextEdit(); self.preview_row_placeholders.setReadOnly(True)

        self.log_text = QTextEdit(); self.log_text.setReadOnly(True)
        self.summary_label = QLabel("Resumen: -")
        self.progress = QProgressBar()

        self.btn_validate = QPushButton("Validar")
        self.btn_generate = QPushButton("Generar documentos")

        self._build_ui()
        self._load_config()

    def _build_ui(self) -> None:
        root = QWidget()
        main_layout = QVBoxLayout(root)

        files_box = QGroupBox("1) Archivos")
        files_form = QFormLayout(files_box)
        files_form.addRow("Excel:", self._path_row(self.excel_edit, self._pick_excel))
        files_form.addRow("Plantilla cero:", self._path_row(self.tpl_zero_edit, lambda: self._pick_docx(self.tpl_zero_edit)))
        files_form.addRow("Plantilla negativo:", self._path_row(self.tpl_neg_edit, lambda: self._pick_docx(self.tpl_neg_edit)))
        files_form.addRow("Plantilla positivo:", self._path_row(self.tpl_pos_edit, lambda: self._pick_docx(self.tpl_pos_edit)))
        files_form.addRow("Carpeta salida:", self._path_row(self.output_edit, self._pick_output))
        files_form.addRow("Hoja Excel:", self.sheet_combo)

        top_buttons = QHBoxLayout()
        btn_reload_config = QPushButton("Recargar config")
        btn_save_config = QPushButton("Guardar config")
        btn_open_out = QPushButton("Abrir salida")
        btn_reload_config.clicked.connect(self._load_config)
        btn_save_config.clicked.connect(self._save_config)
        btn_open_out.clicked.connect(self._open_output)
        self.btn_validate.clicked.connect(self.validate_all)
        self.btn_generate.clicked.connect(self.run_generation)
        self.btn_generate.setEnabled(False)

        top_buttons.addWidget(self.btn_validate)
        top_buttons.addWidget(self.btn_generate)
        top_buttons.addWidget(btn_save_config)
        top_buttons.addWidget(btn_reload_config)
        top_buttons.addWidget(btn_open_out)
        top_buttons.addStretch()

        valid_box = QGroupBox("2) Validación")
        valid_layout = QHBoxLayout(valid_box)
        valid_layout.addWidget(self._with_label("Columnas detectadas", self.columns_text))
        valid_layout.addWidget(self._with_label("Placeholders por plantilla", self.placeholders_text))
        valid_layout.addWidget(self._with_label("Mapeo / faltantes", self.mapping_text))

        preview_box = QGroupBox("3) Vista previa")
        preview_layout = QVBoxLayout(preview_box)
        preview_layout.addWidget(self.preview_table)
        preview_layout.addWidget(QLabel("Placeholders resueltos (fila seleccionada):"))
        preview_layout.addWidget(self.preview_row_placeholders)
        self.preview_table.itemSelectionChanged.connect(self._update_row_preview)

        exec_box = QGroupBox("4) Ejecución")
        exec_layout = QVBoxLayout(exec_box)
        exec_layout.addWidget(self.progress)
        exec_layout.addWidget(self.summary_label)
        exec_layout.addWidget(self.log_text)

        main_layout.addWidget(files_box)
        main_layout.addLayout(top_buttons)
        main_layout.addWidget(valid_box)
        main_layout.addWidget(preview_box)
        main_layout.addWidget(exec_box)

        self.setCentralWidget(root)

    def _with_label(self, title: str, widget: QWidget) -> QWidget:
        box = QGroupBox(title)
        lay = QVBoxLayout(box)
        lay.addWidget(widget)
        return box

    def _path_row(self, line: QLineEdit, cb) -> QWidget:
        w = QWidget()
        l = QHBoxLayout(w)
        l.setContentsMargins(0, 0, 0, 0)
        b = QPushButton("...")
        b.setFixedWidth(35)
        b.clicked.connect(cb)
        l.addWidget(line)
        l.addWidget(b)
        return w

    def _pick_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Excel", "", "Excel (*.xlsx *.xlsm)")
        if path:
            self.excel_edit.setText(path)
            try:
                sheets = self.reader.get_sheet_names(path)
                self.sheet_combo.clear(); self.sheet_combo.addItems(sheets)
            except Exception as exc:
                self._log(f"No se pudo leer hojas: {exc}")

    def _pick_docx(self, target: QLineEdit) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Plantilla", "", "Word (*.docx)")
        if path:
            target.setText(path)

    def _pick_output(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Salida")
        if path:
            self.output_edit.setText(path)

    def validate_all(self) -> None:
        self.btn_generate.setEnabled(False)
        self.log_text.clear()
        result = self.validator.validate_paths(
            self.excel_edit.text(), self.tpl_zero_edit.text(), self.tpl_neg_edit.text(), self.tpl_pos_edit.text(), self.output_edit.text()
        )
        if not result.ok:
            self._log("Validación fallida:")
            for msg in result.messages:
                self._log(f" - {msg}")
            return

        try:
            self.excel_data = self.reader.read(self.excel_edit.text(), self.sheet_combo.currentText() or None)
            if "apr_omitido" not in self.excel_data.columns:
                self._log("Error: columna apr_omitido no presente")
                return
            self._populate_preview_table(self.excel_data)
            available, missing = self.validator.validate_columns(self.excel_data.columns)
            self.columns_text.setPlainText("\n".join(available))

            self.placeholders_by_tpl["cero"] = self.template_engine.scan_placeholders(self.tpl_zero_edit.text()).placeholders
            self.placeholders_by_tpl["negativo"] = self.template_engine.scan_placeholders(self.tpl_neg_edit.text()).placeholders
            self.placeholders_by_tpl["positivo"] = self.template_engine.scan_placeholders(self.tpl_pos_edit.text()).placeholders
            self.placeholders_text.setPlainText(
                "\n".join(
                    [
                        f"cero: {sorted(self.placeholders_by_tpl['cero'])}",
                        f"negativo: {sorted(self.placeholders_by_tpl['negativo'])}",
                        f"positivo: {sorted(self.placeholders_by_tpl['positivo'])}",
                    ]
                )
            )

            warns: list[str] = []
            if missing:
                warns.append(f"Columnas mínimas faltantes: {', '.join(missing)}")
            for key in ["cero", "negativo", "positivo"]:
                ph = self.placeholders_by_tpl[key]
                absent = sorted([p for p in ph if p not in self.excel_data.columns])
                if absent:
                    warns.append(f"{key}: placeholders sin columna -> {', '.join(absent)}")
            self.mapping_text.setPlainText("\n".join(warns) if warns else "Mapeo OK")

            self.btn_generate.setEnabled(True)
            self._log("Validación OK. Puede generar documentos.")
        except Exception as exc:
            self._log(f"Validación con error: {exc}")

    def _populate_preview_table(self, data: ExcelData) -> None:
        cols = list(data.columns) + ["plantilla_asignada"]
        self.preview_table.clear()
        self.preview_table.setColumnCount(len(cols))
        self.preview_table.setHorizontalHeaderLabels(cols)

        sample = data.rows[:30]
        self.preview_table.setRowCount(len(sample))
        for r, row in enumerate(sample):
            for c, col in enumerate(data.columns):
                self.preview_table.setItem(r, c, QTableWidgetItem(row.display.get(col, "")))
            assigned = self._resolve_tpl_label(row)
            self.preview_table.setItem(r, len(cols) - 1, QTableWidgetItem(assigned))
        self.preview_table.resizeColumnsToContents()

    def _resolve_tpl_label(self, row) -> str:
        try:
            apr_source = row.raw.get("apr_omitido")
            n = parse_numeric(apr_source if apr_source is not None else row.display.get("apr_omitido", ""))
            if n == 0:
                return "cero"
            if n < 0:
                return "negativo"
            return "positivo"
        except Exception:
            return "error_apr_omitido"

    def _update_row_preview(self) -> None:
        if self.excel_data is None:
            return
        selected = self.preview_table.currentRow()
        if selected < 0 or selected >= len(self.excel_data.rows):
            return
        row = self.excel_data.rows[selected]
        tpl = self._resolve_tpl_label(row)
        placeholders = self.placeholders_by_tpl.get(tpl, set())
        lines = [f"{p} => {row.display.get(p, '')}" for p in sorted(placeholders)]
        self.preview_row_placeholders.setPlainText("\n".join(lines))

    def run_generation(self) -> None:
        if self.excel_data is None:
            QMessageBox.warning(self, "Atención", "Primero valide los archivos")
            return

        logger = LoggingService(Path(self.output_edit.text()))
        self.progress.setValue(0)

        summary = self.generator.generate(
            rows=self.excel_data.rows,
            tpl_zero=self.tpl_zero_edit.text(),
            tpl_neg=self.tpl_neg_edit.text(),
            tpl_pos=self.tpl_pos_edit.text(),
            output_dir=self.output_edit.text(),
            placeholders_by_tpl=self.placeholders_by_tpl,
            logger=logger,
            progress_cb=self._on_progress,
            log_cb=self._log,
        )
        self.summary_label.setText(
            f"Resumen: total={summary.total_filas} OK={summary.ok} error={summary.error} advertencias={summary.advertencias}"
        )
        self._log(f"Reporte: {summary.reporte_path}")
        self._log(f"Log: {summary.log_path}")
        QMessageBox.information(self, "Proceso completado", self.summary_label.text())

    def _on_progress(self, current: int, total: int) -> None:
        pct = int((current / total) * 100) if total else 0
        self.progress.setValue(pct)

    def _log(self, msg: str) -> None:
        self.log_text.append(msg)

    def _save_config(self) -> None:
        self.config_service.save(
            {
                "excel_path": self.excel_edit.text(),
                "template_zero": self.tpl_zero_edit.text(),
                "template_negative": self.tpl_neg_edit.text(),
                "template_positive": self.tpl_pos_edit.text(),
                "output_dir": self.output_edit.text(),
                "sheet_name": self.sheet_combo.currentText(),
            }
        )
        self._log("Configuración guardada")

    def _load_config(self) -> None:
        cfg = self.config_service.load()
        self.excel_edit.setText(cfg.get("excel_path", ""))
        self.tpl_zero_edit.setText(cfg.get("template_zero", ""))
        self.tpl_neg_edit.setText(cfg.get("template_negative", ""))
        self.tpl_pos_edit.setText(cfg.get("template_positive", ""))
        self.output_edit.setText(cfg.get("output_dir", ""))

        excel = cfg.get("excel_path", "")
        if excel and Path(excel).exists():
            try:
                sheets = self.reader.get_sheet_names(excel)
                self.sheet_combo.clear(); self.sheet_combo.addItems(sheets)
                wanted = cfg.get("sheet_name", "")
                if wanted in sheets:
                    self.sheet_combo.setCurrentText(wanted)
            except Exception:
                pass

    def _open_output(self) -> None:
        try:
            open_directory(self.output_edit.text())
        except Exception as exc:
            QMessageBox.warning(self, "Aviso", f"No se pudo abrir carpeta: {exc}")

    def keyPressEvent(self, event) -> None:  # noqa: N802
        if event.key() == Qt.Key_F5:
            self.validate_all()
            return
        super().keyPressEvent(event)
