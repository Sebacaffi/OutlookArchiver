"""
startup.py - Inicio automatico con Windows via registro HKCU
Soporta modo silencioso (--silent) para no mostrar ventana al arrancar
"""

import sys
import logging
import winreg
from pathlib import Path

logger   = logging.getLogger(__name__)
REG_KEY  = r"Software\Microsoft\Windows\CurrentVersion\Run"
REG_NAME = "OutlookArchiver"


def get_executable_path(silent: bool = True) -> str:
    flag = " --silent" if silent else ""
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"{flag}'
    return f'"{sys.executable}" "{Path(__file__).parent / "main.py"}"{flag}'


def enable_autostart(silent: bool = True) -> bool:
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, REG_KEY, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, REG_NAME, 0, winreg.REG_SZ,
                          get_executable_path(silent))
        winreg.CloseKey(key)
        mode = "silencioso" if silent else "normal"
        logger.info("Inicio automatico habilitado (modo %s).", mode)
        return True
    except Exception as e:
        logger.error("Error al habilitar inicio automatico: %s", e)
        return False


def disable_autostart() -> bool:
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, REG_KEY, 0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(key, REG_NAME)
        winreg.CloseKey(key)
        logger.info("Inicio automatico deshabilitado.")
        return True
    except FileNotFoundError:
        return True
    except Exception as e:
        logger.error("Error al deshabilitar inicio automatico: %s", e)
        return False


def autostart_enabled() -> bool:
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, REG_KEY, 0, winreg.KEY_READ)
        winreg.QueryValueEx(key, REG_NAME)
        winreg.CloseKey(key)
        return True
    except FileNotFoundError:
        return False
    except Exception:
        return False
