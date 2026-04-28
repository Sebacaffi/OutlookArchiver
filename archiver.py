"""
archiver.py - Logica principal de archivado de Outlook

Nombres de PST:
  - Archivo 2026.pst, Archivo 2026-2.pst, Archivo 2026-3.pst ...
  - El programa busca el primer PST del año con espacio disponible.
  - Si no existe ninguno, crea Archivo YYYY.pst.
  - Cambio de año: el 01/02/2027 la fecha de corte es 01/01/2027,
    el ultimo dia archivado es 31/12/2026 -> PST de 2026.
"""

import os
import shutil
import logging
import subprocess
from datetime import datetime, date, timedelta
from pathlib import Path

logger = logging.getLogger(__name__)
PST_HARD_LIMIT_GB    = 47.0
ONEDRIVE_FOLDER_NAME = "OneDrive - Agencia de Aduanas I.P. Hardy y Cía. Ltda"


# ── Tamanio ───────────────────────────────────────────────────────────────────

def get_ost_size_gb() -> float:
    outlook_dir = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Outlook"
    ost_files   = list(outlook_dir.glob("*.ost"))
    if not ost_files:
        logger.warning("No se encontro archivo .ost en %s", outlook_dir)
        return 0.0
    ost     = max(ost_files, key=lambda f: f.stat().st_mtime)
    size_gb = ost.stat().st_size / (1024 ** 3)
    logger.info("Archivo .ost: %s | Tamanio: %.2f GB", ost.name, size_gb)
    return round(size_gb, 2)


def get_pst_size_gb(pst_path: str) -> float:
    p = Path(pst_path)
    return round(p.stat().st_size / (1024 ** 3), 2) if p.exists() else 0.0


# ── Fechas ────────────────────────────────────────────────────────────────────

def compute_cutoff_date(today: date = None) -> date:
    """Mismo dia del mes, un mes antes de hoy."""
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
    """Año del ultimo dia archivado (cutoff - 1 dia)."""
    return (cutoff - timedelta(days=1)).year


# ── Rutas PST con sufijo numerico ─────────────────────────────────────────────

def get_pst_candidates(base_dir: str, year: int) -> list[Path]:
    """
    Retorna lista ordenada de PSTs existentes para un año:
    [Archivo 2026.pst, Archivo 2026-2.pst, Archivo 2026-3.pst, ...]
    """
    base = Path(base_dir)
    candidates = []
    # PST base
    p = base / f"Archivo {year}.pst"
    if p.exists():
        candidates.append(p)
    # PSTs con sufijo
    i = 2
    while True:
        p = base / f"Archivo {year}-{i}.pst"
        if p.exists():
            candidates.append(p)
            i += 1
        else:
            break
    return candidates


def get_active_pst_path(base_dir: str, year: int, pst_max_gb: float) -> tuple[Path, bool]:
    """
    Busca el PST activo para un año: el primero que NO este lleno.
    Si todos estan llenos, genera la ruta del siguiente sufijo.
    Retorna (path, es_nuevo).
    """
    base       = Path(base_dir)
    candidates = get_pst_candidates(base_dir, year)

    for pst in candidates:
        if get_pst_size_gb(str(pst)) < pst_max_gb:
            return pst, False

    # Todos llenos o no existe ninguno -> crear nuevo
    if not candidates:
        return base / f"Archivo {year}.pst", True
    # Siguiente sufijo
    last = candidates[-1].stem   # "Archivo 2026-3"
    parts = last.split("-")
    next_suffix = int(parts[-1]) + 1 if len(parts) > 1 and parts[-1].isdigit() else 2
    return base / f"Archivo {year}-{next_suffix}.pst", True


def get_next_pst_suffix(base_dir: str, year: int) -> Path:
    """Genera la ruta del siguiente PST de un año (para rotar cuando se llena)."""
    base       = Path(base_dir)
    candidates = get_pst_candidates(base_dir, year)
    if not candidates:
        return base / f"Archivo {year}.pst"
    last   = candidates[-1].stem
    parts  = last.split("-")
    suffix = int(parts[-1]) + 1 if len(parts) > 1 and parts[-1].isdigit() else 2
    return base / f"Archivo {year}-{suffix}.pst"


# ── OneDrive ──────────────────────────────────────────────────────────────────

def find_onedrive_path() -> Path | None:
    user_profile  = Path(os.environ.get("USERPROFILE", Path.home()))
    onedrive_path = user_profile / ONEDRIVE_FOLDER_NAME
    return onedrive_path if onedrive_path.exists() else None


def get_onedrive_backup_path(config: dict) -> Path | None:
    root     = find_onedrive_path()
    subpath  = config.get("onedrive_subpath", "Respaldo Correo").strip()
    if not root:
        return None
    return root / subpath if subpath else root


# ── Outlook COM ───────────────────────────────────────────────────────────────

def close_outlook() -> bool:
    try:
        import win32com.client
        ol = win32com.client.GetActiveObject("Outlook.Application")
        ol.Quit()
        import time; time.sleep(3)
        logger.info("Outlook cerrado.")
        return True
    except Exception:
        return False


def open_outlook():
    try:
        for pf in ("ProgramFiles", "ProgramFiles(x86)"):
            exe = Path(os.environ.get(pf, "")) \
                  / "Microsoft Office" / "root" / "Office16" / "OUTLOOK.EXE"
            if exe.exists():
                subprocess.Popen([str(exe)])
                logger.info("Outlook reabierto.")
                return
        subprocess.Popen(["outlook.exe"], shell=True)
    except Exception as e:
        logger.warning("No se pudo reabrir Outlook: %s", e)


def ensure_pst_store(namespace, pst_path: str, display_name: str):
    pst_path = str(Path(pst_path).resolve())
    Path(pst_path).parent.mkdir(parents=True, exist_ok=True)

    for store in namespace.Stores:
        try:
            if store.FilePath and \
               str(Path(store.FilePath).resolve()).lower() == pst_path.lower():
                try:
                    root = store.GetRootFolder()
                    if root.Name != display_name:
                        root.Name = display_name
                except Exception:
                    pass
                return store
        except Exception:
            continue

    logger.info("Abriendo/creando PST: %s", pst_path)
    namespace.AddStoreEx(pst_path, 3)

    for store in namespace.Stores:
        try:
            if store.FilePath and \
               str(Path(store.FilePath).resolve()).lower() == pst_path.lower():
                try:
                    store.GetRootFolder().Name = display_name
                    logger.info("Store nombrado: %s", display_name)
                except Exception as e:
                    logger.warning("No se pudo nombrar el store: %s", e)
                return store
        except Exception:
            continue

    raise RuntimeError(f"No se pudo abrir PST: {pst_path}")


def get_or_create_subfolder(root_folder, name: str):
    for folder in root_folder.Folders:
        if folder.Name == name:
            return folder
    return root_folder.Folders.Add(name)


# ── Archivado de items ────────────────────────────────────────────────────────

def archive_folder_items(src_folder, dst_folder, cutoff_dt: datetime,
                         moved_count: list, pst_path: str, pst_max_gb: float):
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
                logger.warning("PST alcanzo %.1f GB. Deteniendo.", pst_max_gb)
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


# ── Backup OneDrive ───────────────────────────────────────────────────────────

def backup_pst_to_onedrive(pst_path: str, config: dict) -> dict:
    backup_dir = get_onedrive_backup_path(config)
    if not backup_dir:
        msg = "OneDrive no encontrado. Respaldo omitido."
        logger.warning(msg)
        return {"status": "warning", "message": msg}

    pst_file    = Path(pst_path)
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup_dest = backup_dir / pst_file.name

    if backup_dest.exists():
        return {"status": "ok", "message": f"Ya existia en OneDrive: {backup_dest.name}"}

    logger.info("Cerrando Outlook para copiar PST a OneDrive...")
    was_open = close_outlook()

    try:
        shutil.copy2(str(pst_path), str(backup_dest))
        sz  = get_pst_size_gb(str(backup_dest))
        msg = f"PST copiado a OneDrive: {backup_dest.name} ({sz:.2f} GB)"
        logger.info(msg)
        result = {"status": "ok", "message": msg}
    except Exception as e:
        msg = f"Error al copiar PST a OneDrive: {e}"
        logger.error(msg)
        result = {"status": "error", "message": msg}

    if was_open:
        import time; time.sleep(2)
        open_outlook()

    return result


# ── Funcion principal ─────────────────────────────────────────────────────────

def run_archive(config: dict) -> dict:
    try:
        import win32com.client
        import win32timezone
    except ImportError:
        return {"status": "error",
                "message": "pywin32 no instalado.",
                "size_gb": 0, "moved": 0, "pst_path": "", "cutoff": ""}

    result = {"status": "ok", "size_gb": 0, "moved": 0,
              "pst_path": "", "cutoff": "", "message": ""}

    try:
        size_gb   = get_ost_size_gb()
        result["size_gb"] = size_gb
        threshold = float(config.get("threshold_gb", 4.0))

        if size_gb < threshold:
            result["message"] = (
                f"Buzon en {size_gb:.2f} GB — bajo umbral ({threshold} GB). Sin accion.")
            logger.info(result["message"])
            return result

        logger.info("Buzon supera umbral (%.2f > %.2f GB). Iniciando archivado...",
                    size_gb, threshold)

        today        = date.today()
        cutoff       = compute_cutoff_date(today)
        arc_year     = compute_archive_year(cutoff)
        cutoff_dt    = datetime(cutoff.year, cutoff.month, cutoff.day)
        base_dir     = config.get("pst_base_dir", str(Path.home() / "Documents"))
        pst_max_gb   = float(config.get("pst_max_gb", PST_HARD_LIMIT_GB))

        # Obtener PST activo (con logica de sufijos)
        pst_path, is_new = get_active_pst_path(base_dir, arc_year, pst_max_gb)

        # Si todos los PSTs del año estaban llenos, respaldar el ultimo y rotar
        if is_new and get_pst_candidates(base_dir, arc_year):
            last_full = get_pst_candidates(base_dir, arc_year)[-1]
            logger.warning("Todos los PSTs de %d estan llenos. Creando: %s",
                           arc_year, pst_path.name)
            if config.get("onedrive_backup", False):
                backup_pst_to_onedrive(str(last_full), config)

        display_name = pst_path.stem  # "Archivo 2026" o "Archivo 2026-2"
        result["pst_path"] = str(pst_path)
        result["cutoff"]   = cutoff.strftime("%d/%m/%Y")

        logger.info("PST destino: %s", pst_path.name)
        logger.info("Fecha de corte: %s (hasta %s inclusive)",
                    cutoff.strftime("%d/%m/%Y"),
                    (cutoff - timedelta(days=1)).strftime("%d/%m/%Y"))

        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            outlook = win32com.client.Dispatch("Outlook.Application")

        namespace = outlook.GetNamespace("MAPI")
        store     = ensure_pst_store(namespace, str(pst_path), display_name)
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
                        moved_count, str(pst_path), pst_max_gb)
                    if not ok:
                        pst_full = True
                        break
            if pst_full:
                break

        result["moved"] = moved_count[0]

        # PST se lleno durante el archivado -> respaldar
        if pst_full and config.get("onedrive_backup", False):
            logger.info("PST lleno. Respaldando a OneDrive...")
            br = backup_pst_to_onedrive(str(pst_path), config)
            logger.info("Respaldo: %s", br["message"])
            try:
                outlook   = win32com.client.GetActiveObject("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
            except Exception:
                pass

        if pst_full:
            result["status"]  = "warning"
            result["message"] = (
                f"Archivado parcial: PST {pst_path.name} alcanzo el limite. "
                f"Movidos: {moved_count[0]} correos.")
        else:
            result["message"] = (
                f"Completado. Correos anteriores al {cutoff.strftime('%d/%m/%Y')} "
                f"movidos: {moved_count[0]}. PST: {pst_path.name} | "
                f"Buzon previo: {size_gb:.2f} GB")

        logger.info(result["message"])

        # Apagar equipo si esta configurado
        if config.get("shutdown_after", False) and result["status"] == "ok":
            logger.info("Apagando equipo en 60 segundos...")
            subprocess.Popen(["shutdown", "/s", "/t", "60",
                               "/c", "Outlook Archiver: archivado completado."])

    except Exception as e:
        result["status"]  = "error"
        result["message"] = f"Error durante el archivado: {e}"
        logger.error(result["message"], exc_info=True)

    return result
