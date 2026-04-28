from __future__ import annotations

import logging
from pathlib import Path


class LoggingService:
    def __init__(self, output_dir: Path) -> None:
        self.output_dir = output_dir
        self.log_file = output_dir / "log_generacion.txt"
        self.logger = logging.getLogger("doc_generator")
        self.logger.setLevel(logging.INFO)
        self.logger.handlers.clear()

        file_handler = logging.FileHandler(self.log_file, encoding="utf-8")
        formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)

    def info(self, msg: str) -> None:
        self.logger.info(msg)

    def warning(self, msg: str) -> None:
        self.logger.warning(msg)

    def error(self, msg: str) -> None:
        self.logger.error(msg)
