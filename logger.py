"""
logger.py — Configuración del sistema de logging
"""

import logging
import os
from pathlib import Path
from logging.handlers import RotatingFileHandler


def setup(log_path: str = None):
    """Configura logging a archivo rotativo + consola."""
    if log_path is None:
        appdata = os.environ.get("APPDATA", Path.home())
        log_path = Path(appdata) / "OutlookArchiver" / "archiver.log"

    log_path = Path(log_path)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    fmt = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    root = logging.getLogger()
    root.setLevel(logging.INFO)

    # Archivo rotativo — máximo 2 MB, 3 backups
    fh = RotatingFileHandler(log_path, maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8")
    fh.setFormatter(fmt)
    root.addHandler(fh)

    # Consola (útil al correr desde terminal o en desarrollo)
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    root.addHandler(ch)

    return log_path
