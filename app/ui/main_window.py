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

from app.core.document_registry import MODULES, TEMPLATE_LABELS
from app.core.document_selector import DocumentSelector
from app.core.excel_reader import ExcelData, ExcelReader
from app.core.generator import DocumentGenerator
from app.core.template_engine import TemplateEngine
from app.core.validator import Validator
from app.services.config_service import ConfigService
from app.services.logging_service import LoggingService
from app.services.report_service import ReportService
from app.utils.file_utils import open_directory


class MainWindow(QMainWindow):
    TEMPLATE_EDIT_KEYS = [
        "rd_apr_cero",
        "rd_apr_negativo",
        "rd_apr_positivo",
        "rd_dge_cero",
        "rd_dge_negativo",
        "rd_dge_positivo",
        "rm_apr_176_1",
        "rm_apr_178_1",
        "rm_dge_176_1",
        "rm_dge_178_1",
    ]

    def __init__(self, config_service: ConfigService) -> None:
        super().__init__()
        self.setWindowTitle("Generador maestro DOCX (RD/RM APR-DGE)")
        self.resize(1500, 920)

        self.config_service = config_service
        self.reader = ExcelReader()
        self.validator = Validator()
        self.template_engine = TemplateEngine()
        self.generator = DocumentGenerator(self.template_engine, DocumentSelector())

        self.excel_data_by_module: dict[str, ExcelData] = {}
        self.placeholders_by_template: dict[str, set[str]] = {}

        self.excel_edit = QLineEdit()
        self.output_edit = QLineEdit()
        self.preview_sheet_combo = QComboBox()

        self.template_edits: dict[str, QLineEdit] = {k: QLineEdit() for k in self.TEMPLATE_EDIT_KEYS}

        self.columns_text = QTextEdit(); self.columns_text.setReadOnly(True)
        self.placeholders_text = QTextEdit(); self.placeholders_text.setReadOnly(True)
        self.mapping_text = QTextEdit(); self.mapping_text.setReadOnly(True)

        self.preview_table = QTableWidget()
        self.preview_row_placeholders = QTextEdit(); self.preview_row_placeholders.setReadOnly(True)

        self.log_text = QTextEdit(); self.log_text.setReadOnly(True)
        self.summary_label = QLabel("Resumen: -")
        self.progress = QProgressBar()

        self.btn_validate = QPushButton("Validar")
        self.btn_gen_rd_apr = QPushButton("Generar RD APR")
        self.btn_gen_rd_dge = QPushButton("Generar RD DGE")
        self.btn_gen_rm_apr = QPushButton("Generar RM APR")
        self.btn_gen_rm_dge = QPushButton("Generar RM DGE")
        self.btn_gen_all = QPushButton("Generar todo")

        self._build_ui()
        self._load_config()

    def _build_ui(self) -> None:
        root = QWidget()
        main_layout = QVBoxLayout(root)

        files_box = QGroupBox("1) Excel maestro, plantillas y salida")
        form = QFormLayout(files_box)
        form.addRow("Excel maestro:", self._path_row(self.excel_edit, self._pick_excel))
        form.addRow("Carpeta salida:", self._path_row(self.output_edit, self._pick_output))
        form.addRow("Preview hoja:", self.preview_sheet_combo)

        form.addRow("RD APR cero:", self._path_row(self.template_edits["rd_apr_cero"], lambda: self._pick_docx("rd_apr_cero")))
        form.addRow("RD APR negativo:", self._path_row(self.template_edits["rd_apr_negativo"], lambda: self._pick_docx("rd_apr_negativo")))
        form.addRow("RD APR positivo:", self._path_row(self.template_edits["rd_apr_positivo"], lambda: self._pick_docx("rd_apr_positivo")))
        form.addRow("RD DGE cero:", self._path_row(self.template_edits["rd_dge_cero"], lambda: self._pick_docx("rd_dge_cero")))
        form.addRow("RD DGE negativo:", self._path_row(self.template_edits["rd_dge_negativo"], lambda: self._pick_docx("rd_dge_negativo")))
        form.addRow("RD DGE positivo:", self._path_row(self.template_edits["rd_dge_positivo"], lambda: self._pick_docx("rd_dge_positivo")))
        form.addRow("RM APR 176-1:", self._path_row(self.template_edits["rm_apr_176_1"], lambda: self._pick_docx("rm_apr_176_1")))
        form.addRow("RM APR 178-1:", self._path_row(self.template_edits["rm_apr_178_1"], lambda: self._pick_docx("rm_apr_178_1")))
        form.addRow("RM DGE 176-1:", self._path_row(self.template_edits["rm_dge_176_1"], lambda: self._pick_docx("rm_dge_176_1")))
        form.addRow("RM DGE 178-1:", self._path_row(self.template_edits["rm_dge_178_1"], lambda: self._pick_docx("rm_dge_178_1")))

        controls = QHBoxLayout()
        btn_reload_config = QPushButton("Recargar config")
        btn_save_config = QPushButton("Guardar config")
        btn_open_out = QPushButton("Abrir salida")

        btn_reload_config.clicked.connect(self._load_config)
        btn_save_config.clicked.connect(self._save_config)
        btn_open_out.clicked.connect(self._open_output)
        self.btn_validate.clicked.connect(self.validate_all)

        controls.addWidget(self.btn_validate)
        controls.addWidget(self.btn_gen_rd_apr)
        controls.addWidget(self.btn_gen_rd_dge)
        controls.addWidget(self.btn_gen_rm_apr)
        controls.addWidget(self.btn_gen_rm_dge)
        controls.addWidget(self.btn_gen_all)
        controls.addWidget(btn_save_config)
        controls.addWidget(btn_reload_config)
        controls.addWidget(btn_open_out)
        controls.addStretch()

        self.btn_gen_rd_apr.clicked.connect(lambda: self.run_generation(["RD_APR"]))
        self.btn_gen_rd_dge.clicked.connect(lambda: self.run_generation(["RD_DGE"]))
        self.btn_gen_rm_apr.clicked.connect(lambda: self.run_generation(["RM_APR"]))
        self.btn_gen_rm_dge.clicked.connect(lambda: self.run_generation(["RM_DGE"]))
        self.btn_gen_all.clicked.connect(lambda: self.run_generation(list(MODULES.keys())))

        valid_box = QGroupBox("2) Validación")
        valid_layout = QHBoxLayout(valid_box)
        valid_layout.addWidget(self._with_label("Columnas detectadas (por hoja)", self.columns_text))
        valid_layout.addWidget(self._with_label("Placeholders por plantilla", self.placeholders_text))
        valid_layout.addWidget(self._with_label("Advertencias / mapeo", self.mapping_text))

        preview_box = QGroupBox("3) Vista previa")
        preview_layout = QVBoxLayout(preview_box)
        preview_layout.addWidget(self.preview_table)
        preview_layout.addWidget(QLabel("Placeholders resueltos (fila seleccionada):"))
        preview_layout.addWidget(self.preview_row_placeholders)
        self.preview_table.itemSelectionChanged.connect(self._update_row_preview)
        self.preview_sheet_combo.currentTextChanged.connect(self._refresh_preview_for_selected_sheet)

        exec_box = QGroupBox("4) Ejecución")
        exec_layout = QVBoxLayout(exec_box)
        exec_layout.addWidget(self.progress)
        exec_layout.addWidget(self.summary_label)
        exec_layout.addWidget(self.log_text)

        main_layout.addWidget(files_box)
        main_layout.addLayout(controls)
        main_layout.addWidget(valid_box)
        main_layout.addWidget(preview_box)
        main_layout.addWidget(exec_box)

        self.setCentralWidget(root)
        self._set_generate_enabled(False)

    def _set_generate_enabled(self, enabled: bool) -> None:
        self.btn_gen_rd_apr.setEnabled(enabled)
        self.btn_gen_rd_dge.setEnabled(enabled)
        self.btn_gen_rm_apr.setEnabled(enabled)
        self.btn_gen_rm_dge.setEnabled(enabled)
        self.btn_gen_all.setEnabled(enabled)

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
        path, _ = QFileDialog.getOpenFileName(self, "Excel maestro", "", "Excel (*.xlsx *.xlsm)")
        if path:
            self.excel_edit.setText(path)
            self._load_excel_sheets(path)

    def _pick_docx(self, key: str) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Plantilla", "", "Word (*.docx)")
        if path:
            self.template_edits[key].setText(path)

    def _pick_output(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Salida")
        if path:
            self.output_edit.setText(path)

    def _load_excel_sheets(self, excel_path: str) -> None:
        try:
            sheets = self.reader.get_sheet_names(excel_path)
            self.preview_sheet_combo.clear()
            self.preview_sheet_combo.addItems(sheets)
        except Exception as exc:
            self._log(f"No se pudo leer hojas: {exc}")

    def _collect_template_paths(self) -> dict[str, str]:
        return {key: edit.text() for key, edit in self.template_edits.items()}

    def validate_all(self) -> None:
        self._set_generate_enabled(False)
        self.log_text.clear()
        self.excel_data_by_module.clear()
        self.placeholders_by_template.clear()

        template_paths = self._collect_template_paths()

        path_result = self.validator.validate_master_paths(
            excel_path=self.excel_edit.text(),
            template_paths=template_paths,
            output_dir=self.output_edit.text(),
        )
        if not path_result.ok:
            for m in path_result.messages:
                self._log(m)
            return

        try:
            sheet_names = self.reader.get_sheet_names(self.excel_edit.text())
        except Exception as exc:
            self._log(f"No se pudo abrir Excel maestro: {exc}")
            return

        sheet_result = self.validator.validate_sheets(sheet_names)
        if not sheet_result.ok:
            for m in sheet_result.messages:
                self._log(m)
            return

        all_columns_lines: list[str] = []
        mapping_warnings: list[str] = []

        for module_key, module in MODULES.items():
            data = self.reader.read(self.excel_edit.text(), module.sheet_name)
            self.excel_data_by_module[module_key] = data
            all_columns_lines.append(f"[{module.sheet_name}] {', '.join(data.columns)}")
            for w in data.warnings:
                mapping_warnings.append(w)

        for t_key, path in template_paths.items():
            scan = self.template_engine.scan_placeholders(path)
            self.placeholders_by_template[t_key] = scan.placeholders

        placeholder_lines = [
            f"{TEMPLATE_LABELS.get(k, k)}: {sorted(v)}" for k, v in self.placeholders_by_template.items()
        ]

        for module_key, module in MODULES.items():
            cols = self.excel_data_by_module[module_key].columns
            for tpl_key in module.template_keys:
                warns = self.validator.validate_placeholders_vs_columns(
                    placeholders=self.placeholders_by_template.get(tpl_key, set()),
                    columns=cols,
                    template_key=tpl_key,
                )
                mapping_warnings.extend(warns)

        self.columns_text.setPlainText("\n".join(all_columns_lines))
        self.placeholders_text.setPlainText("\n".join(placeholder_lines))
        self.mapping_text.setPlainText("\n".join(mapping_warnings) if mapping_warnings else "Mapeo OK")

        # Actualizar combo de preview con hojas válidas del flujo
        self.preview_sheet_combo.clear()
        self.preview_sheet_combo.addItems([m.sheet_name for m in MODULES.values()])
        self._refresh_preview_for_selected_sheet()

        self._set_generate_enabled(True)
        self._log("Validación completa OK. Puede generar por módulo o todo.")

    def _module_key_from_sheet(self, sheet_name: str) -> str | None:
        for key, item in MODULES.items():
            if item.sheet_name == sheet_name:
                return key
        return None

    def _refresh_preview_for_selected_sheet(self) -> None:
        sheet = self.preview_sheet_combo.currentText()
        module_key = self._module_key_from_sheet(sheet)
        if not module_key or module_key not in self.excel_data_by_module:
            return

        data = self.excel_data_by_module[module_key]
        cols = list(data.columns) + ["plantilla_asignada"]
        sample = data.rows[:25]

        self.preview_table.clear()
        self.preview_table.setColumnCount(len(cols))
        self.preview_table.setHorizontalHeaderLabels(cols)
        self.preview_table.setRowCount(len(sample))

        selector = DocumentSelector()
        paths = self._collect_template_paths()
        for r, row in enumerate(sample):
            for c, col in enumerate(data.columns):
                self.preview_table.setItem(r, c, QTableWidgetItem(row.display.get(col, "")))
            try:
                sel = selector.select(module_key, row.raw, row.display, paths)
                assigned = TEMPLATE_LABELS.get(sel.template_key, sel.template_key)
            except Exception:
                assigned = "error_selector"
            self.preview_table.setItem(r, len(cols) - 1, QTableWidgetItem(assigned))

        self.preview_table.resizeColumnsToContents()

    def _update_row_preview(self) -> None:
        sheet = self.preview_sheet_combo.currentText()
        module_key = self._module_key_from_sheet(sheet)
        if not module_key or module_key not in self.excel_data_by_module:
            return

        data = self.excel_data_by_module[module_key]
        selected = self.preview_table.currentRow()
        if selected < 0 or selected >= len(data.rows):
            return

        row = data.rows[selected]
        try:
            sel = DocumentSelector().select(module_key, row.raw, row.display, self._collect_template_paths())
            placeholders = self.placeholders_by_template.get(sel.template_key, set())
        except Exception:
            placeholders = set()
        lines = [f"{p} => {row.display.get(p, '')}" for p in sorted(placeholders)]
        self.preview_row_placeholders.setPlainText("\n".join(lines))

    def run_generation(self, module_keys: list[str]) -> None:
        if not self.excel_data_by_module:
            QMessageBox.warning(self, "Atención", "Primero valide los archivos")
            return

        output = Path(self.output_edit.text())
        logger = LoggingService(output)
        report_service = ReportService(output)
        all_report_rows: list[dict[str, str]] = []
        total_ok = total_err = total_warn = total_rows = 0
        template_paths = self._collect_template_paths()

        for module_key in module_keys:
            data = self.excel_data_by_module.get(module_key)
            if data is None:
                continue
            self._log(f"Iniciando {module_key}...")
            self.progress.setValue(0)
            summary, rows = self.generator.generate_module(
                module_key=module_key,
                rows=data.rows,
                template_paths=template_paths,
                output_dir=self.output_edit.text(),
                placeholders_by_template=self.placeholders_by_template,
                logger=logger,
                progress_cb=self._on_progress,
                log_cb=self._log,
            )
            all_report_rows.extend(rows)
            total_ok += summary.ok
            total_err += summary.error
            total_warn += summary.advertencias + len(data.warnings)
            total_rows += summary.total_filas

        report_path = report_service.save(all_report_rows)
        self.summary_label.setText(
            f"Resumen: total={total_rows} OK={total_ok} error={total_err} advertencias={total_warn}"
        )
        self._log(f"Reporte: {report_path}")
        self._log(f"Log: {logger.log_file}")
        QMessageBox.information(self, "Proceso completado", self.summary_label.text())

    def _on_progress(self, current: int, total: int) -> None:
        pct = int((current / total) * 100) if total else 0
        self.progress.setValue(pct)

    def _log(self, msg: str) -> None:
        self.log_text.append(msg)

    def _save_config(self) -> None:
        config = {
            "excel_path": self.excel_edit.text(),
            "output_dir": self.output_edit.text(),
            "preview_sheet": self.preview_sheet_combo.currentText(),
        }
        for key, edit in self.template_edits.items():
            config[f"template_{key}"] = edit.text()
        self.config_service.save(config)
        self._log("Configuración guardada")

    def _load_config(self) -> None:
        cfg = self.config_service.load()
        self.excel_edit.setText(cfg.get("excel_path", ""))
        self.output_edit.setText(cfg.get("output_dir", ""))
        for key, edit in self.template_edits.items():
            edit.setText(cfg.get(f"template_{key}", ""))

        excel = cfg.get("excel_path", "")
        if excel and Path(excel).exists():
            self._load_excel_sheets(excel)
            preview = cfg.get("preview_sheet", "RD_APR")
            if self.preview_sheet_combo.findText(preview) >= 0:
                self.preview_sheet_combo.setCurrentText(preview)

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
