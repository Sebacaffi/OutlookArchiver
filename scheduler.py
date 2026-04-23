"""
scheduler.py — Registra/actualiza/elimina la tarea en el Programador de Windows
usando schtasks.exe (no requiere permisos de administrador para tareas de usuario)
"""

import subprocess
import sys
import os
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

TASK_NAME = "OutlookArchiverTask"


def get_executable_path() -> str:
    """Retorna la ruta del ejecutable actual (.exe o script .py)."""
    if getattr(sys, "frozen", False):
        return sys.executable  # Empaquetado con PyInstaller
    return f'"{sys.executable}" "{Path(__file__).parent / "main.py"}" --run'


def register_task(hour: int, minute: int) -> bool:
    """Crea o actualiza la tarea programada en Windows."""
    exe_path = get_executable_path()
    time_str = f"{hour:02d}:{minute:02d}"

    cmd = [
        "schtasks", "/Create",
        "/F",                         # Sobreescribir si existe
        "/TN", TASK_NAME,
        "/TR", f'{exe_path} --run',
        "/SC", "DAILY",
        "/ST", time_str,
        "/RU", os.environ.get("USERNAME", ""),  # Usuario actual
    ]

    logger.info("Registrando tarea: %s a las %s", TASK_NAME, time_str)
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode == 0:
        logger.info("Tarea registrada correctamente.")
        return True
    else:
        logger.error("Error al registrar tarea: %s", result.stderr)
        return False


def remove_task() -> bool:
    """Elimina la tarea del Programador de Windows."""
    cmd = ["schtasks", "/Delete", "/F", "/TN", TASK_NAME]
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result.returncode == 0


def task_exists() -> bool:
    """Verifica si la tarea ya está registrada."""
    cmd = ["schtasks", "/Query", "/TN", TASK_NAME]
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result.returncode == 0


def run_task_now() -> bool:
    """Ejecuta la tarea inmediatamente."""
    cmd = ["schtasks", "/Run", "/TN", TASK_NAME]
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result.returncode == 0