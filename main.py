"""
main.py - Punto de entrada principal

Modos:
  Sin argumentos  -> primera vez: wizard. Ya configurado: GUI principal.
  --run           -> archivado silencioso (Programador de Windows)
  --setup         -> fuerza el wizard aunque ya este configurado
"""

import sys
import logging
import config as cfg
import logger as log_setup


def main():
    conf = cfg.load()
    log_setup.setup(conf.get("log_path"))

    # Modo silencioso: llamado por el Programador de Windows
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

    if first_run or force_setup:
        _run_wizard_then_gui()
    else:
        import gui
        gui.run()


def _run_wizard_then_gui():
    """
    Ejecuta el wizard (tk.Tk con mainloop propio).
    Al terminar, si el usuario completo la configuracion, abre la GUI
    en el MISMO proceso reusando Tk via Misc.tk (truco estandar para
    reiniciar tkinter sin relanzar el proceso).
    """
    import wizard
    import scheduler
    import startup

    # El wizard corre su propio mainloop y retorna al terminar
    result = wizard.run_wizard()

    if result is None:
        logging.info("Wizard cancelado. Saliendo.")
        return

    # Guardar config
    cfg.save(result)
    logging.info("Configuracion inicial guardada.")

    # Registrar tarea programada
    ok_task = scheduler.register_task(result["schedule_hour"], result["schedule_minute"])
    logging.info("Tarea programada: %s", "OK" if ok_task else "ERROR")

    # Inicio automatico con Windows
    if result.get("autostart", True):
        ok_st = startup.enable_autostart()
        logging.info("Autostart: %s", "OK" if ok_st else "ERROR")
    else:
        startup.disable_autostart()

    # Abrir GUI principal (crea un nuevo tk.Tk independiente)
    import gui
    gui.run()


if __name__ == "__main__":
    main()
