from __future__ import annotations

import json
from pathlib import Path
from typing import Any


class ConfigService:
    def __init__(self, config_path: Path) -> None:
        self.config_path = config_path
        self.defaults: dict[str, Any] = {
            "excel_path": "",
            "output_dir": "",
            "template_rd_apr_cero": "",
            "template_rd_apr_negativo": "",
            "template_rd_apr_positivo": "",
            "template_rd_dge_cero": "",
            "template_rd_dge_negativo": "",
            "template_rd_dge_positivo": "",
            "template_rm_apr_176_1": "",
            "template_rm_apr_178_1": "",
            "template_rm_dge_176_1": "",
            "template_rm_dge_178_1": "",
            "preview_sheet": "RD_APR",
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
