"""
scheduler.py - Registra/actualiza/elimina la tarea en el Programador de Windows
Soporta frecuencia: diaria, semanal (dia configurable), mensual
"""

import subprocess
import sys
import os
import logging
from pathlib import Path

logger   = logging.getLogger(__name__)
TASK_NAME = "OutlookArchiverTask"

DAYS_ES = {
    "MON": "Lunes", "TUE": "Martes", "WED": "Miercoles",
    "THU": "Jueves", "FRI": "Viernes", "SAT": "Sabado", "SUN": "Domingo",
}


def get_executable_path() -> str:
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    return f'"{sys.executable}" "{Path(__file__).parent / "main.py"}"'


def register_task(config: dict) -> bool:
    """Crea o actualiza la tarea segun la configuracion completa."""
    exe_path = get_executable_path()
    hour     = int(config.get("schedule_hour",   20))
    minute   = int(config.get("schedule_minute",  0))
    freq     = config.get("schedule_freq",  "daily")
    day      = config.get("schedule_day",   "MON")
    time_str = f"{hour:02d}:{minute:02d}"

    cmd = [
        "schtasks", "/Create", "/F",
        "/TN", TASK_NAME,
        "/TR", f"{exe_path} --run",
        "/ST", time_str,
        "/RU", os.environ.get("USERNAME", ""),
    ]

    if freq == "weekly":
        cmd += ["/SC", "WEEKLY", "/D", day]
        logger.info("Registrando tarea: %s | semanal %s a las %s",
                    TASK_NAME, DAYS_ES.get(day, day), time_str)
    elif freq == "monthly":
        cmd += ["/SC", "MONTHLY", "/D", "1"]
        logger.info("Registrando tarea: %s | mensual (dia 1) a las %s",
                    TASK_NAME, time_str)
    else:
        cmd += ["/SC", "DAILY"]
        logger.info("Registrando tarea: %s | diaria a las %s", TASK_NAME, time_str)

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode == 0:
        logger.info("Tarea registrada correctamente.")
        return True
    logger.error("Error al registrar tarea: %s", result.stderr.strip())
    return False


def remove_task() -> bool:
    result = subprocess.run(
        ["schtasks", "/Delete", "/F", "/TN", TASK_NAME],
        capture_output=True, text=True,
    )
    return result.returncode == 0


def task_exists() -> bool:
    result = subprocess.run(
        ["schtasks", "/Query", "/TN", TASK_NAME],
        capture_output=True, text=True,
    )
    return result.returncode == 0


def run_task_now() -> bool:
    result = subprocess.run(
        ["schtasks", "/Run", "/TN", TASK_NAME],
        capture_output=True, text=True,
    )
    return result.returncode == 0
