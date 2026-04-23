"""
config.py - Configuracion local en JSON
"""

import json
import os
from pathlib import Path

CONFIG_DIR  = Path(os.environ.get("APPDATA", Path.home())) / "OutlookArchiver"
CONFIG_FILE = CONFIG_DIR / "config.json"

DEFAULTS = {
    # Umbral del buzon principal para disparar el archivado
    "threshold_gb":  3.0,
    # Directorio base donde se crean los PST anuales
    "pst_base_dir":  str(Path.home() / "Documents" / "ArchivosOutlook"),
    # Limite de tamanio por PST antes de rotar (margen bajo el limite de 50 GB)
    "pst_max_gb":    47.0,
    # Hora y minuto de ejecucion diaria
    "schedule_hour":   20,
    "schedule_minute":  0,
    # Extras
    "notify_email":  "",
    "log_path":      str(CONFIG_DIR / "archiver.log"),
    "enabled":       True,
    "setup_done":    False,
    "autostart":     True,
}


def load() -> dict:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if not CONFIG_FILE.exists():
        save(DEFAULTS)
        return DEFAULTS.copy()
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
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
