"""
archiver.py — Lógica principal de archivado de Outlook
Usa win32com para conectarse a Outlook y mover correos antiguos a un .pst
"""

import os
import logging
from datetime import datetime, timedelta
from pathlib import Path

logger = logging.getLogger(__name__)


def get_ost_size_gb() -> float:
    """Retorna el tamaño del archivo .ost principal en GB."""
    outlook_dir = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Outlook"
    ost_files = list(outlook_dir.glob("*.ost"))
    if not ost_files:
        logger.warning("No se encontró archivo .ost en %s", outlook_dir)
        return 0.0
    # Tomar el más reciente
    ost = max(ost_files, key=lambda f: f.stat().st_mtime)
    size_gb = ost.stat().st_size / (1024 ** 3)
    logger.info("Archivo .ost: %s | Tamaño: %.2f GB", ost.name, size_gb)
    return round(size_gb, 2)


def ensure_pst_store(outlook_app, pst_path: str):
    """Agrega el .pst de archivo a Outlook si no está ya abierto."""
    namespace = outlook_app.GetNamespace("MAPI")
    pst_path = str(Path(pst_path).resolve())

    # Verificar si ya está abierto
    for store in namespace.Stores:
        try:
            if store.FilePath and str(Path(store.FilePath).resolve()).lower() == pst_path.lower():
                logger.info("El archivo .pst ya está abierto en Outlook.")
                return store
        except Exception:
            continue

    # Agregar el .pst (lo crea si no existe)
    logger.info("Agregando archivo .pst: %s", pst_path)
    namespace.AddStoreEx(pst_path, 3)  # 3 = olStoreUnicode

    # Volver a buscar el store recién añadido
    for store in namespace.Stores:
        try:
            if store.FilePath and str(Path(store.FilePath).resolve()).lower() == pst_path.lower():
                return store
        except Exception:
            continue

    raise RuntimeError(f"No se pudo acceder al archivo .pst: {pst_path}")


def archive_folder(folder, archive_root_folder, cutoff_date: datetime, moved_count: list):
    """Mueve recursivamente correos más antiguos que cutoff_date al archivo .pst."""
    try:
        items = folder.Items
        items.Sort("[ReceivedTime]")
        to_move = []

        for item in items:
            try:
                received = item.ReceivedTime
                # ReceivedTime viene como datetime con timezone info de COM
                if received.replace(tzinfo=None) < cutoff_date:
                    to_move.append(item)
            except AttributeError:
                # Algunos elementos (reuniones, etc.) no tienen ReceivedTime
                continue

        for item in to_move:
            try:
                item.Move(archive_root_folder)
                moved_count[0] += 1
            except Exception as e:
                logger.warning("No se pudo mover un elemento: %s", e)

        # Procesar subcarpetas recursivamente
        for subfolder in folder.Folders:
            archive_folder(subfolder, archive_root_folder, cutoff_date, moved_count)

    except Exception as e:
        logger.error("Error procesando carpeta %s: %s", folder.Name, e)


def run_archive(config: dict) -> dict:
    """
    Ejecuta el proceso de archivado según la configuración.
    Retorna un diccionario con el resultado: status, size_gb, moved, message.
    """
    try:
        import win32com.client
    except ImportError:
        return {
            "status": "error",
            "message": "pywin32 no está instalado. Ejecuta: pip install pywin32",
            "size_gb": 0,
            "moved": 0,
        }

    result = {"status": "ok", "size_gb": 0, "moved": 0, "message": ""}

    try:
        size_gb = get_ost_size_gb()
        result["size_gb"] = size_gb
        threshold = float(config.get("threshold_gb", 4))

        if size_gb < threshold:
            result["message"] = f"Buzón en {size_gb:.2f} GB — por debajo del umbral ({threshold} GB). Sin acción."
            logger.info(result["message"])
            return result

        logger.info("Buzón supera umbral (%.2f GB > %.2f GB). Iniciando archivado...", size_gb, threshold)

        months = int(config.get("months_old", 12))
        pst_path = config.get("pst_path", str(Path.home() / "Documents" / "OutlookArchivo.pst"))
        cutoff = datetime.now() - timedelta(days=months * 30)

        # Conectar a Outlook (debe estar abierto)
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            outlook = win32com.client.Dispatch("Outlook.Application")

        namespace = outlook.GetNamespace("MAPI")
        store = ensure_pst_store(outlook, pst_path)
        archive_root = store.GetRootFolder()

        moved_count = [0]

        # Archivar carpetas principales
        folders_to_archive = ["Bandeja de entrada", "Inbox", "Elementos enviados", "Sent Items"]
        for account in namespace.Accounts:
            for folder in namespace.Folders:
                for subfolder in folder.Folders:
                    if subfolder.Name in folders_to_archive:
                        logger.info("Archivando carpeta: %s", subfolder.Name)
                        archive_folder(subfolder, archive_root, cutoff, moved_count)

        result["moved"] = moved_count[0]
        result["message"] = (
            f"Archivado completado. Correos anteriores a {cutoff.strftime('%d/%m/%Y')} "
            f"movidos: {moved_count[0]}. Tamaño previo: {size_gb:.2f} GB."
        )
        logger.info(result["message"])

    except Exception as e:
        result["status"] = "error"
        result["message"] = f"Error durante el archivado: {e}"
        logger.error(result["message"], exc_info=True)

    return result
