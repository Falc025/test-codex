from __future__ import annotations

import re
import subprocess
from pathlib import Path

INVALID_WIN_CHARS = r'[<>:"/\\|?*\x00-\x1F]'


def ensure_directory(path: str | Path) -> Path:
    target = Path(path)
    target.mkdir(parents=True, exist_ok=True)
    return target


def sanitize_filename(name: str, fallback: str = "archivo") -> str:
    clean = re.sub(INVALID_WIN_CHARS, "_", str(name or "").strip())
    clean = re.sub(r"\s+", "_", clean)
    clean = re.sub(r"_+", "_", clean).strip("._")
    return clean or fallback


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    i = 1
    while True:
        candidate = path.with_name(f"{stem}_{i}{suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def open_directory(path: str | Path) -> None:
    target = str(Path(path).resolve())
    subprocess.Popen(["explorer", target], shell=True)
