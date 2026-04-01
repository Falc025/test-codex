from __future__ import annotations

import json
from pathlib import Path
from typing import Any


class ConfigService:
    def __init__(self, config_path: Path) -> None:
        self.config_path = config_path
        self.defaults: dict[str, Any] = {
            "excel_path": "",
            "template_zero": "",
            "template_negative": "",
            "template_positive": "",
            "output_dir": "",
            "sheet_name": "",
        }

    def load(self) -> dict[str, Any]:
        if not self.config_path.exists():
            return dict(self.defaults)
        try:
            data = json.loads(self.config_path.read_text(encoding="utf-8"))
        except Exception:
            return dict(self.defaults)
        cfg = dict(self.defaults)
        cfg.update({k: data.get(k, v) for k, v in self.defaults.items()})
        return cfg

    def save(self, config: dict[str, Any]) -> None:
        merged = dict(self.defaults)
        merged.update(config)
        self.config_path.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
