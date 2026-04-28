"""
main.py - Punto de entrada

Modos:
  Sin argumentos  -> primera vez: wizard. Ya configurado: GUI.
  --silent        -> inicia en bandeja sin mostrar ventana
  --run           -> archivado silencioso (Programador de Windows)
  --setup         -> fuerza el wizard
"""

import sys
import logging
import config as cfg
import logger as log_setup


def main():
    conf = cfg.load()
    log_setup.setup(conf.get("log_path"))

    if "--run" in sys.argv:
        import archiver
        logging.info("=" * 60)
        logging.info("Inicio de archivado programado")
        result = archiver.run_archive(conf)
        logging.info("Resultado: %s", result["message"])
        logging.info("=" * 60)
        return

    first_run   = not conf.get("setup_done", False)
    force_setup = "--setup" in sys.argv
    silent      = "--silent" in sys.argv

    if first_run or force_setup:
        _run_wizard_then_gui()
    else:
        import gui
        gui.run(start_hidden=silent)


def _run_wizard_then_gui():
    import wizard, scheduler, startup
    result = wizard.run_wizard()

    if result is None:
        logging.info("Wizard cancelado.")
        return

    cfg.save(result)
    scheduler.register_task(result)
    if result.get("autostart", True):
        startup.enable_autostart(silent=result.get("autostart_silent", True))
    else:
        startup.disable_autostart()

    import gui
    gui.run(start_hidden=False)


if __name__ == "__main__":
    main()
