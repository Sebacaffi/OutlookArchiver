"""
startup.py — Gestión del inicio automático con Windows
Usa el registro HKCU (sin permisos de administrador)
"""

import sys
import logging
import winreg
from pathlib import Path

logger = logging.getLogger(__name__)

REG_KEY  = r"Software\Microsoft\Windows\CurrentVersion\Run"
REG_NAME = "OutlookArchiver"


def get_executable_path() -> str:
    """Retorna la ruta del ejecutable actual."""
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    return f'"{sys.executable}" "{Path(__file__).parent / "main.py"}"'


def enable_autostart() -> bool:
    """Registra la app para que inicie con Windows (HKCU, sin admin)."""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, REG_KEY,
            0, winreg.KEY_SET_VALUE
        )
        winreg.SetValueEx(key, REG_NAME, 0, winreg.REG_SZ, get_executable_path())
        winreg.CloseKey(key)
        logger.info("Inicio automático habilitado en el registro de Windows.")
        return True
    except Exception as e:
        logger.error("Error al habilitar inicio automático: %s", e)
        return False


def disable_autostart() -> bool:
    """Elimina la entrada de inicio automático del registro."""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, REG_KEY,
            0, winreg.KEY_SET_VALUE
        )
        winreg.DeleteValue(key, REG_NAME)
        winreg.CloseKey(key)
        logger.info("Inicio automático deshabilitado.")
        return True
    except FileNotFoundError:
        logger.info("La entrada de inicio automático no existía.")
        return True
    except Exception as e:
        logger.error("Error al deshabilitar inicio automático: %s", e)
        return False


def autostart_enabled() -> bool:
    """Comprueba si la entrada de inicio automático existe."""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, REG_KEY,
            0, winreg.KEY_READ
        )
        winreg.QueryValueEx(key, REG_NAME)
        winreg.CloseKey(key)
        return True
    except FileNotFoundError:
        return False
    except Exception:
        return False
