"""
build.py — Script para generar el instalador .exe con PyInstaller
Ejecutar desde la carpeta del proyecto:
    python build.py
"""

import subprocess
import sys
from pathlib import Path

PROJECT_DIR = Path(__file__).parent
DIST_DIR    = PROJECT_DIR / "dist"
BUILD_DIR   = PROJECT_DIR / "build"
ICON_PATH   = PROJECT_DIR / "icon.ico"

def build():
    print("=" * 55)
    print("  Outlook Archiver — Build")
    print("=" * 55)

    # Verificar que PyInstaller esté instalado
    try:
        import PyInstaller
        print(f"  PyInstaller {PyInstaller.__version__} encontrado.")
    except ImportError:
        print("  PyInstaller no encontrado. Instalando...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    args = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                         # Un solo .exe
        "--windowed",                        # Sin ventana de consola al abrir GUI
        "--name", "OutlookArchiver",
        "--add-data", f"{PROJECT_DIR / 'config.py'};.",
        "--add-data", f"{PROJECT_DIR / 'archiver.py'};.",
        "--add-data", f"{PROJECT_DIR / 'scheduler.py'};.",
        "--add-data", f"{PROJECT_DIR / 'logger.py'};.",
        "--add-data", f"{PROJECT_DIR / 'gui.py'};.",
        "--add-data", f"{PROJECT_DIR / 'startup.py'};.",
        "--add-data", f"{PROJECT_DIR / 'wizard.py'};.",
        "--hidden-import", "win32com.client",
        "--hidden-import", "win32com.shell",
        "--hidden-import", "pywintypes",
        "--hidden-import", "win32timezone",
        "--hidden-import", "winreg",
        "--hidden-import", "pystray",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        "--distpath", str(DIST_DIR),
        "--workpath", str(BUILD_DIR),
        "--specpath", str(PROJECT_DIR),
        "--clean",
        str(PROJECT_DIR / "main.py"),
    ]

    # Agregar ícono si existe
    if ICON_PATH.exists():
        args += ["--icon", str(ICON_PATH)]

    print("\n  Ejecutando PyInstaller...\n")
    result = subprocess.run(args, cwd=PROJECT_DIR)

    if result.returncode == 0:
        exe_path = DIST_DIR / "OutlookArchiver.exe"
        print("\n" + "=" * 55)
        print(f"  ✔ Build exitoso!")
        print(f"  Ejecutable: {exe_path}")
        print("=" * 55)
    else:
        print("\n  ✖ Error durante el build. Revisa los mensajes anteriores.")
        sys.exit(1)


if __name__ == "__main__":
    build()
