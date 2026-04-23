"""
main.py — Punto de entrada principal

Modos de ejecución:
  Sin argumentos  → primera vez: wizard de configuración
                    ya configurado: abre GUI de ajustes
  --run           → archivado silencioso (llamado por el Programador de Windows)
  --setup         → forzar el wizard de configuración inicial
"""

import sys
import logging
import config as cfg
import logger as log_setup


def main():
    conf = cfg.load()
    log_setup.setup(conf.get("log_path"))

    # ── Modo silencioso: llamado por el Programador de Windows ───────────────
    if "--run" in sys.argv:
        import archiver
        logging.info("=" * 60)
        logging.info("Inicio de archivado programado")
        result = archiver.run_archive(conf)
        logging.info("Resultado: %s", result["message"])
        logging.info("=" * 60)
        return

    # ── Primera ejecución o --setup forzado ──────────────────────────────────
    first_run = not conf.get("setup_done", False)
    force_setup = "--setup" in sys.argv

    if first_run or force_setup:
        _run_wizard()
        return

    # ── Ejecución normal: abrir GUI de configuración ─────────────────────────
    import gui
    gui.run()


def _run_wizard():
    """Lanza el wizard, aplica la configuración resultante y abre la GUI."""
    import wizard
    import scheduler
    import startup

    result = wizard.run_wizard()

    if result is None:
        logging.info("Wizard cancelado por el usuario.")
        return

    # Guardar configuración
    cfg.save(result)
    logging.info("Configuración inicial guardada.")

    # Registrar tarea en el Programador de Windows
    ok_task = scheduler.register_task(result["schedule_hour"], result["schedule_minute"])
    if ok_task:
        logging.info("Tarea programada registrada correctamente.")
    else:
        logging.warning("No se pudo registrar la tarea programada.")

    # Registrar inicio automático con Windows (según preferencia del usuario)
    if result.get("autostart", True):
        ok_startup = startup.enable_autostart()
        if ok_startup:
            logging.info("Inicio automático con Windows habilitado.")
        else:
            logging.warning("No se pudo habilitar el inicio automático.")
    else:
        startup.disable_autostart()

    # Abrir GUI de configuración principal
    import gui
    gui.run()


if __name__ == "__main__":
    main()
