"""
config.py — Lectura y escritura de configuración local en JSON
"""

import json
import os
from pathlib import Path

CONFIG_DIR = Path(os.environ.get("APPDATA", Path.home())) / "OutlookArchiver"
CONFIG_FILE = CONFIG_DIR / "config.json"

DEFAULTS = {
    "threshold_gb": 4.0,
    "months_old": 12,
    "pst_path": str(Path.home() / "Documents" / "OutlookArchivo.pst"),
    "schedule_hour": 20,
    "schedule_minute": 0,
    "notify_email": "",
    "log_path": str(CONFIG_DIR / "archiver.log"),
    "enabled": True,
    "setup_done": False,   # True una vez completado el wizard inicial
    "autostart": True,
}


def load() -> dict:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if not CONFIG_FILE.exists():
        save(DEFAULTS)
        return DEFAULTS.copy()
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    # Rellenar claves faltantes con defaults
    for k, v in DEFAULTS.items():
        data.setdefault(k, v)
    return data


def save(config: dict):
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def get_config_path() -> Path:
    return CONFIG_FILE


def get_log_path(config: dict) -> Path:
    return Path(config.get("log_path", DEFAULTS["log_path"]))
