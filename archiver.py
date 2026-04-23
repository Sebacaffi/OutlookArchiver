"""
archiver.py - Logica principal de archivado de Outlook

Logica de nombres de archivo PST por año:
  - El PST de destino se llama "Archivo YYYY" donde YYYY es el año de los correos.
  - "Archivo XXXX" contiene correos desde el 01/01/XXXX hasta el 31/12/XXXX.
  - La fecha de corte es "1 mes antes de hoy" (ultimo dia inclusive = corte - 1 dia).
  - Si estamos en el mismo año que la fecha de corte -> PST del año actual.
  - Si ya pasamos al año siguiente (ej: 01/02/2027 archivando lo de 2026):
      el ultimo archivado del año anterior va al PST del año anterior,
      luego los del año nuevo van al PST del nuevo año.
  - El PST tiene limite ~50 GB. La herramienta rota automaticamente al siguiente
    archivo anual cuando el actual supera pst_max_gb.
"""

import os
import logging
from datetime import datetime, date, timedelta
from pathlib import Path

logger = logging.getLogger(__name__)

PST_HARD_LIMIT_GB = 47.0   # Margen de seguridad antes del limite de 50 GB


def get_ost_size_gb() -> float:
    """Retorna el tamanio del archivo .ost principal en GB."""
    outlook_dir = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Outlook"
    ost_files = list(outlook_dir.glob("*.ost"))
    if not ost_files:
        logger.warning("No se encontro archivo .ost en %s", outlook_dir)
        return 0.0
    ost = max(ost_files, key=lambda f: f.stat().st_mtime)
    size_gb = ost.stat().st_size / (1024 ** 3)
    logger.info("Archivo .ost: %s | Tamanio: %.2f GB", ost.name, size_gb)
    return round(size_gb, 2)


def get_pst_size_gb(pst_path: str) -> float:
    """Retorna el tamanio de un PST en GB (0.0 si no existe)."""
    p = Path(pst_path)
    if not p.exists():
        return 0.0
    return round(p.stat().st_size / (1024 ** 3), 2)


def compute_cutoff_date(today: date = None) -> date:
    """
    Fecha de corte = primer dia del mes actual - 1 mes.
    Es decir: se archivan correos con fecha < cutoff (hasta el dia anterior inclusive).

    Ejemplos:
      Hoy 22/04/2026  -> cutoff 22/03/2026  (archiva hasta 21/03/2026 inclusive)
      Hoy 01/02/2027  -> cutoff 01/01/2027  (archiva hasta 31/12/2026 inclusive)
    """
    if today is None:
        today = date.today()
    # Restar 1 mes manteniendo el dia (si el dia no existe en el mes anterior, usa el ultimo)
    month = today.month - 1
    year  = today.year
    if month == 0:
        month = 12
        year -= 1
    # Ajustar dia si necesario (ej: 31 de marzo -> 28/29 de febrero)
    import calendar
    last_day = calendar.monthrange(year, month)[1]
    day = min(today.day, last_day)
    return date(year, month, day)


def compute_archive_year(cutoff: date) -> int:
    """
    El año del archivo PST es el año de los correos que se van a archivar.

    Regla:
      - Si cutoff es dentro del mismo año que hoy -> año actual.
      - Si cutoff cae en el año anterior (ej: cutoff=01/01/2027 cuando hoy=01/02/2027)
        los correos archivados son del 2026, pero el ultimo lote completa el 2026.
        En este caso usamos el año de cutoff - 1 dia (ultimo dia archivado).
    """
    last_archived_day = cutoff - timedelta(days=1)
    return last_archived_day.year


def get_pst_path_for_year(base_dir: str, year: int) -> str:
    """Construye la ruta del PST para un año dado."""
    return str(Path(base_dir) / f"Archivo {year}.pst")


def ensure_pst_store(namespace, pst_path: str):
    """Abre o crea el PST en Outlook y retorna el store."""
    pst_path = str(Path(pst_path).resolve())
    Path(pst_path).parent.mkdir(parents=True, exist_ok=True)

    for store in namespace.Stores:
        try:
            if store.FilePath and \
               str(Path(store.FilePath).resolve()).lower() == pst_path.lower():
                logger.info("PST ya abierto: %s", Path(pst_path).name)
                return store
        except Exception:
            continue

    logger.info("Abriendo/creando PST: %s", pst_path)
    namespace.AddStoreEx(pst_path, 3)  # 3 = olStoreUnicode

    for store in namespace.Stores:
        try:
            if store.FilePath and \
               str(Path(store.FilePath).resolve()).lower() == pst_path.lower():
                return store
        except Exception:
            continue

    raise RuntimeError(f"No se pudo abrir el PST: {pst_path}")


def get_or_create_subfolder(root_folder, name: str):
    """Obtiene o crea una subcarpeta dentro del PST."""
    for folder in root_folder.Folders:
        if folder.Name == name:
            return folder
    return root_folder.Folders.Add(name)


def archive_folder_items(src_folder, dst_folder, cutoff_dt: datetime, moved_count: list,
                         pst_path: str, pst_max_gb: float):
    """
    Mueve items de src_folder a dst_folder cuyo ReceivedTime < cutoff_dt.
    Se detiene si el PST supera pst_max_gb.
    Procesa subcarpetas recursivamente.
    """
    try:
        items = src_folder.Items
        items.Sort("[ReceivedTime]")
        to_move = []

        for item in items:
            try:
                if item.ReceivedTime.replace(tzinfo=None) < cutoff_dt:
                    to_move.append(item)
            except AttributeError:
                continue

        for item in to_move:
            # Verificar tamanio del PST antes de cada movimiento
            if get_pst_size_gb(pst_path) >= pst_max_gb:
                logger.warning("PST alcanzo el limite de %.1f GB. Deteniendo archivado.", pst_max_gb)
                return False  # Senal de PST lleno

            try:
                item.Move(dst_folder)
                moved_count[0] += 1
            except Exception as e:
                logger.warning("No se pudo mover un elemento: %s", e)

        # Procesar subcarpetas
        for subfolder in src_folder.Folders:
            dst_sub = get_or_create_subfolder(dst_folder, subfolder.Name)
            if not archive_folder_items(subfolder, dst_sub, cutoff_dt,
                                        moved_count, pst_path, pst_max_gb):
                return False  # PST lleno, propagar senal

    except Exception as e:
        logger.error("Error procesando carpeta %s: %s", src_folder.Name, e)

    return True  # OK


def run_archive(config: dict) -> dict:
    """
    Ejecuta el archivado completo segun configuracion.
    Retorna dict con: status, size_gb, moved, pst_path, cutoff, message.
    """
    try:
        import win32com.client
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
        size_gb = get_ost_size_gb()
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

        # Calcular fecha de corte y año del archivo
        today     = date.today()
        cutoff    = compute_cutoff_date(today)
        arc_year  = compute_archive_year(cutoff)
        cutoff_dt = datetime(cutoff.year, cutoff.month, cutoff.day)  # sin tzinfo

        pst_base_dir = config.get("pst_base_dir",
                                   str(Path.home() / "Documents"))
        pst_max_gb   = float(config.get("pst_max_gb", PST_HARD_LIMIT_GB))
        pst_path     = get_pst_path_for_year(pst_base_dir, arc_year)

        # Si el PST del año actual ya esta lleno, rotar al siguiente
        if get_pst_size_gb(pst_path) >= pst_max_gb:
            logger.warning(
                "PST %s esta lleno (%.1f GB). Rotando a siguiente archivo.",
                Path(pst_path).name, get_pst_size_gb(pst_path)
            )
            pst_path = get_pst_path_for_year(pst_base_dir, arc_year + 1)

        result["pst_path"] = pst_path
        result["cutoff"]   = cutoff.strftime("%d/%m/%Y")

        logger.info("PST destino: %s", pst_path)
        logger.info("Fecha de corte: %s (archiva hasta %s inclusive)",
                    cutoff.strftime("%d/%m/%Y"),
                    (cutoff - timedelta(days=1)).strftime("%d/%m/%Y"))

        # Conectar a Outlook
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            outlook = win32com.client.Dispatch("Outlook.Application")

        namespace = outlook.GetNamespace("MAPI")
        store     = ensure_pst_store(namespace, pst_path)
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
                        subfolder, dst, cutoff_dt, moved_count,
                        pst_path, pst_max_gb
                    )
                    if not ok:
                        pst_full = True
                        break
            if pst_full:
                break

        result["moved"] = moved_count[0]

        if pst_full:
            result["status"]  = "warning"
            result["message"] = (
                f"Archivado parcial: PST alcanzo el limite de {pst_max_gb:.0f} GB. "
                f"Movidos: {moved_count[0]} correos hasta {cutoff.strftime('%d/%m/%Y')}. "
                f"PST: {Path(pst_path).name}"
            )
        else:
            result["message"] = (
                f"Archivado completado. Correos anteriores al {cutoff.strftime('%d/%m/%Y')} "
                f"movidos: {moved_count[0]}. "
                f"PST: {Path(pst_path).name} | Buzon previo: {size_gb:.2f} GB"
            )

        logger.info(result["message"])

    except Exception as e:
        result["status"]  = "error"
        result["message"] = f"Error durante el archivado: {e}"
        logger.error(result["message"], exc_info=True)

    return result
