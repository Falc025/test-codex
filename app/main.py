from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication

from app.services.config_service import ConfigService
from app.ui.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Generador Masivo DOCX")

    base_dir = Path(__file__).resolve().parent
    config_service = ConfigService(base_dir / "config.json")
    window = MainWindow(config_service)
    window.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
