"""
archiver.py - Logica principal de archivado de Outlook

Logica de nombres de archivo PST por año:
  - El PST de destino se llama "Archivo YYYY.pst" y aparece en Outlook como "Archivo YYYY".
  - La fecha de corte es 1 mes antes de hoy (correos archivados hasta cutoff-1 dia inclusive).
  - Cuando un PST supera pst_max_gb, rota al año siguiente.
  - Al rotar, el PST lleno se copia a OneDrive (si esta configurado) y se cierra Outlook.
"""

import os
import shutil
import logging
import subprocess
from datetime import datetime, date, timedelta
from pathlib import Path

logger = logging.getLogger(__name__)

PST_HARD_LIMIT_GB = 47.0


# ── Utilidades de tamanio ─────────────────────────────────────────────────────

def get_ost_size_gb() -> float:
    outlook_dir = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Outlook"
    ost_files   = list(outlook_dir.glob("*.ost"))
    if not ost_files:
        logger.warning("No se encontro archivo .ost en %s", outlook_dir)
        return 0.0
    ost      = max(ost_files, key=lambda f: f.stat().st_mtime)
    size_gb  = ost.stat().st_size / (1024 ** 3)
    logger.info("Archivo .ost: %s | Tamanio: %.2f GB", ost.name, size_gb)
    return round(size_gb, 2)


def get_pst_size_gb(pst_path: str) -> float:
    p = Path(pst_path)
    if not p.exists():
        return 0.0
    return round(p.stat().st_size / (1024 ** 3), 2)


# ── Logica de fechas ──────────────────────────────────────────────────────────

def compute_cutoff_date(today: date = None) -> date:
    """Fecha de corte = mismo dia del mes anterior."""
    if today is None:
        today = date.today()
    import calendar
    month = today.month - 1
    year  = today.year
    if month == 0:
        month = 12
        year -= 1
    day = min(today.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def compute_archive_year(cutoff: date) -> int:
    """El año del PST es el año del ultimo dia que se archiva (cutoff - 1 dia)."""
    return (cutoff - timedelta(days=1)).year


def get_pst_path_for_year(base_dir: str, year: int) -> str:
    return str(Path(base_dir) / f"Archivo {year}.pst")


# ── OneDrive ──────────────────────────────────────────────────────────────────

ONEDRIVE_FOLDER_NAME = "OneDrive - Agencia de Aduanas I.P. Hardy y Cía. Ltda"


def find_onedrive_path() -> Path | None:
    """
    Retorna la ruta raiz de OneDrive corporativo si existe en el perfil del usuario.
    La carpeta siempre se llama "OneDrive - Agencia de Aduanas I.P. Hardy y Cia. Ltda".
    Si no existe, OneDrive no esta configurado en este equipo.
    """
    user_profile = Path(os.environ.get("USERPROFILE", Path.home()))
    onedrive_path = user_profile / ONEDRIVE_FOLDER_NAME
    if onedrive_path.exists():
        return onedrive_path
    return None


def get_onedrive_backup_path(config: dict) -> Path | None:
    """
    Construye la ruta de backup dentro de OneDrive.
    Retorna None si OneDrive no esta configurado en este equipo.
    """
    onedrive_root    = find_onedrive_path()
    onedrive_subpath = config.get("onedrive_subpath", "Respaldo Correo").strip()

    if not onedrive_root:
        return None

    backup_path = onedrive_root / onedrive_subpath if onedrive_subpath else onedrive_root
    return backup_path


# ── Outlook COM helpers ───────────────────────────────────────────────────────

def close_outlook() -> bool:
    """Cierra Outlook si esta abierto. Retorna True si se cerro."""
    try:
        import win32com.client
        try:
            ol = win32com.client.GetActiveObject("Outlook.Application")
            ol.Quit()
            import time
            time.sleep(3)  # Esperar a que cierre completamente
            logger.info("Outlook cerrado correctamente.")
            return True
        except Exception:
            return False  # Ya estaba cerrado
    except ImportError:
        return False


def open_outlook():
    """Reabre Outlook."""
    try:
        outlook_exe = Path(os.environ.get("ProgramFiles", "C:\\Program Files")) \
                      / "Microsoft Office" / "root" / "Office16" / "OUTLOOK.EXE"
        if not outlook_exe.exists():
            # Buscar en Program Files (x86)
            outlook_exe = Path(os.environ.get("ProgramFiles(x86)", "")) \
                          / "Microsoft Office" / "root" / "Office16" / "OUTLOOK.EXE"
        if outlook_exe.exists():
            subprocess.Popen([str(outlook_exe)])
            logger.info("Outlook reabierto.")
        else:
            # Fallback: abrir via shell
            subprocess.Popen(["outlook.exe"], shell=True)
            logger.info("Outlook reabierto via shell.")
    except Exception as e:
        logger.warning("No se pudo reabrir Outlook automaticamente: %s", e)


def set_store_display_name(namespace, pst_path: str, display_name: str):
    """Cambia el nombre que muestra Outlook para un store PST."""
    pst_path = str(Path(pst_path).resolve())
    for store in namespace.Stores:
        try:
            if store.FilePath and \
               str(Path(store.FilePath).resolve()).lower() == pst_path.lower():
                root = store.GetRootFolder()
                if root.Name != display_name:
                    root.Name = display_name
                    logger.info("Nombre del store actualizado a: %s", display_name)
                return
        except Exception:
            continue


def ensure_pst_store(namespace, pst_path: str, display_name: str):
    """Abre o crea el PST en Outlook, lo nombra correctamente y retorna el store."""
    pst_path = str(Path(pst_path).resolve())
    Path(pst_path).parent.mkdir(parents=True, exist_ok=True)

    # Verificar si ya esta abierto
    for store in namespace.Stores:
        try:
            if store.FilePath and \
               str(Path(store.FilePath).resolve()).lower() == pst_path.lower():
                logger.info("PST ya abierto: %s", Path(pst_path).name)
                # Asegurar que el nombre sea correcto
                try:
                    root = store.GetRootFolder()
                    if root.Name != display_name:
                        root.Name = display_name
                except Exception:
                    pass
                return store
        except Exception:
            continue

    # Agregar el PST
    logger.info("Abriendo/creando PST: %s", pst_path)
    namespace.AddStoreEx(pst_path, 3)  # 3 = olStoreUnicode

    # Buscar el store recien agregado y nombrarlo
    for store in namespace.Stores:
        try:
            if store.FilePath and \
               str(Path(store.FilePath).resolve()).lower() == pst_path.lower():
                try:
                    root = store.GetRootFolder()
                    root.Name = display_name
                    logger.info("Store nombrado como: %s", display_name)
                except Exception as e:
                    logger.warning("No se pudo nombrar el store: %s", e)
                return store
        except Exception:
            continue

    raise RuntimeError(f"No se pudo abrir el PST: {pst_path}")


def remove_pst_store(namespace, pst_path: str):
    """Cierra (desvincula) un PST de Outlook sin eliminarlo del disco."""
    pst_path = str(Path(pst_path).resolve())
    for store in namespace.Stores:
        try:
            if store.FilePath and \
               str(Path(store.FilePath).resolve()).lower() == pst_path.lower():
                namespace.RemoveStore(store.GetRootFolder())
                logger.info("PST desvinculado de Outlook: %s", Path(pst_path).name)
                return True
        except Exception as e:
            logger.warning("Error al desvincular PST: %s", e)
    return False


def get_or_create_subfolder(root_folder, name: str):
    for folder in root_folder.Folders:
        if folder.Name == name:
            return folder
    return root_folder.Folders.Add(name)


# ── Archivado de items ────────────────────────────────────────────────────────

def archive_folder_items(src_folder, dst_folder, cutoff_dt: datetime,
                         moved_count: list, pst_path: str, pst_max_gb: float):
    """
    Mueve items con ReceivedTime < cutoff_dt de src a dst.
    Retorna False si el PST alcanzo el limite.
    """
    try:
        items = src_folder.Items
        items.Sort("[ReceivedTime]")
        to_move = []

        for item in items:
            try:
                rt = item.ReceivedTime
                try:
                    rt_naive = rt.replace(tzinfo=None)
                except Exception:
                    rt_naive = datetime.strptime(str(rt)[:19], "%Y-%m-%d %H:%M:%S")
                if rt_naive < cutoff_dt:
                    to_move.append(item)
            except AttributeError:
                continue

        for item in to_move:
            if get_pst_size_gb(pst_path) >= pst_max_gb:
                logger.warning("PST alcanzo %.1f GB. Deteniendo archivado.", pst_max_gb)
                return False
            try:
                item.Move(dst_folder)
                moved_count[0] += 1
            except Exception as e:
                logger.warning("No se pudo mover elemento: %s", e)

        for subfolder in src_folder.Folders:
            dst_sub = get_or_create_subfolder(dst_folder, subfolder.Name)
            if not archive_folder_items(subfolder, dst_sub, cutoff_dt,
                                        moved_count, pst_path, pst_max_gb):
                return False

    except Exception as e:
        logger.error("Error procesando carpeta %s: %s", src_folder.Name, e)

    return True


# ── Copia a OneDrive (con cierre de Outlook) ──────────────────────────────────

def backup_pst_to_onedrive(pst_path: str, config: dict) -> dict:
    """
    Copia un PST lleno a OneDrive.
    Requiere cerrar Outlook, copiar, y reabrir.
    Retorna dict con status y message.
    """
    backup_dir = get_onedrive_backup_path(config)
    if not backup_dir:
        msg = "OneDrive no encontrado. No se realizo el respaldo."
        logger.warning(msg)
        return {"status": "warning", "message": msg}

    pst_file    = Path(pst_path)
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup_dest = backup_dir / pst_file.name

    if backup_dest.exists():
        logger.info("El PST ya existe en OneDrive: %s", backup_dest)
        return {"status": "ok", "message": f"Ya existia en OneDrive: {backup_dest.name}"}

    logger.info("Cerrando Outlook para copiar PST a OneDrive...")
    outlook_was_open = close_outlook()

    try:
        logger.info("Copiando %s -> %s", pst_file.name, backup_dir)
        shutil.copy2(str(pst_path), str(backup_dest))
        size_gb = get_pst_size_gb(str(backup_dest))
        msg = (f"PST copiado a OneDrive: {backup_dest.name} ({size_gb:.2f} GB). "
               f"Ruta: {backup_dir}")
        logger.info(msg)
        result = {"status": "ok", "message": msg}
    except Exception as e:
        msg = f"Error al copiar PST a OneDrive: {e}"
        logger.error(msg)
        result = {"status": "error", "message": msg}

    if outlook_was_open:
        import time
        time.sleep(2)
        open_outlook()

    return result


# ── Funcion principal ─────────────────────────────────────────────────────────

def run_archive(config: dict) -> dict:
    """
    Ejecuta el archivado completo segun configuracion.
    Retorna dict con: status, size_gb, moved, pst_path, cutoff, message.
    """
    try:
        import win32com.client
        import win32timezone  # requerido por pywin32 para fechas COM
    except ImportError:
        return {
            "status": "error",
            "message": "pywin32 no instalado. Ejecuta: pip install pywin32",
            "size_gb": 0, "moved": 0, "pst_path": "", "cutoff": "",
        }

    result = {
        "status": "ok", "size_gb": 0, "moved": 0,
        "pst_path": "", "cutoff": "", "message": "",
    }

    try:
        size_gb   = get_ost_size_gb()
        result["size_gb"] = size_gb
        threshold = float(config.get("threshold_gb", 4.0))

        if size_gb < threshold:
            result["message"] = (
                f"Buzon en {size_gb:.2f} GB — bajo umbral ({threshold} GB). Sin accion."
            )
            logger.info(result["message"])
            return result

        logger.info("Buzon supera umbral (%.2f GB > %.2f GB). Iniciando archivado...",
                    size_gb, threshold)

        today        = date.today()
        cutoff       = compute_cutoff_date(today)
        arc_year     = compute_archive_year(cutoff)
        cutoff_dt    = datetime(cutoff.year, cutoff.month, cutoff.day)
        pst_base_dir = config.get("pst_base_dir", str(Path.home() / "Documents"))
        pst_max_gb   = float(config.get("pst_max_gb", PST_HARD_LIMIT_GB))
        pst_path     = get_pst_path_for_year(pst_base_dir, arc_year)
        display_name = f"Archivo {arc_year}"

        # Si el PST del año actual ya esta lleno, respaldarlo y rotar
        rotated = False
        if get_pst_size_gb(pst_path) >= pst_max_gb:
            logger.warning("PST %s esta lleno. Rotando al siguiente año.", display_name)

            if config.get("onedrive_backup", False):
                backup_pst_to_onedrive(pst_path, config)

            arc_year     += 1
            pst_path      = get_pst_path_for_year(pst_base_dir, arc_year)
            display_name  = f"Archivo {arc_year}"
            rotated       = True

        result["pst_path"] = pst_path
        result["cutoff"]   = cutoff.strftime("%d/%m/%Y")

        logger.info("PST destino: %s (%s)", pst_path, display_name)
        logger.info("Fecha de corte: %s (archiva hasta %s inclusive)",
                    cutoff.strftime("%d/%m/%Y"),
                    (cutoff - timedelta(days=1)).strftime("%d/%m/%Y"))

        # Conectar a Outlook
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            outlook = win32com.client.Dispatch("Outlook.Application")

        namespace = outlook.GetNamespace("MAPI")
        store     = ensure_pst_store(namespace, pst_path, display_name)
        arc_root  = store.GetRootFolder()

        moved_count  = [0]
        pst_full     = False
        folders_todo = [
            "Bandeja de entrada", "Inbox",
            "Elementos enviados",  "Sent Items",
            "Elementos eliminados", "Deleted Items",
        ]

        for top_folder in namespace.Folders:
            for subfolder in top_folder.Folders:
                if subfolder.Name in folders_todo:
                    dst = get_or_create_subfolder(arc_root, subfolder.Name)
                    ok  = archive_folder_items(
                        subfolder, dst, cutoff_dt,
                        moved_count, pst_path, pst_max_gb
                    )
                    if not ok:
                        pst_full = True
                        break
            if pst_full:
                break

        result["moved"] = moved_count[0]

        # Si el PST se lleno durante este archivado, respaldarlo ahora
        if pst_full and config.get("onedrive_backup", False):
            logger.info("PST lleno durante archivado. Iniciando respaldo a OneDrive...")
            backup_result = backup_pst_to_onedrive(pst_path, config)
            logger.info("Respaldo: %s", backup_result["message"])

            # Reconectar Outlook tras el cierre del respaldo
            try:
                outlook   = win32com.client.GetActiveObject("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
            except Exception:
                pass

        suffix = f" | Rotado a {display_name}" if rotated else ""

        if pst_full:
            result["status"]  = "warning"
            result["message"] = (
                f"Archivado parcial: PST alcanzo el limite de {pst_max_gb:.0f} GB. "
                f"Movidos: {moved_count[0]} correos. "
                f"PST: {Path(pst_path).name}{suffix}"
            )
        else:
            result["message"] = (
                f"Archivado completado. Correos anteriores al {cutoff.strftime('%d/%m/%Y')} "
                f"movidos: {moved_count[0]}. "
                f"PST: {Path(pst_path).name} | Buzon previo: {size_gb:.2f} GB{suffix}"
            )

        logger.info(result["message"])

    except Exception as e:
        result["status"]  = "error"
        result["message"] = f"Error durante el archivado: {e}"
        logger.error(result["message"], exc_info=True)

    return result
