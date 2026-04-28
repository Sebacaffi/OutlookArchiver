"""
gui.py - Interfaz principal con pestanas + bandeja del sistema

Pestanas:
  Estado    - resumen, tamanios, proximo archivado, botones rapidos
  Archivado - umbral, carpeta PST, limite, frecuencia, apagado
  OneDrive  - backup, subcarpeta, estado de conexion
  Sistema   - autostart, silencioso, tarea, desinstalar, nota admin
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import logging
import os
import subprocess
from pathlib import Path
from datetime import date

import config as cfg
import scheduler
import startup
import archiver

logger = logging.getLogger(__name__)

# ── Paleta ────────────────────────────────────────────────────────────────────
BG      = "#1A1D23"
BG2     = "#22262F"
BG3     = "#2C3140"
ACCENT  = "#4F8EF7"
ACCENT2 = "#3A6ED4"
SUCCESS = "#3DBE7A"
WARNING = "#F0A500"
DANGER  = "#E05252"
TEXT    = "#E8EAF0"
TEXT2   = "#9AA0B2"
BORDER  = "#353A47"
FONT_H  = ("Segoe UI", 11, "bold")
FONT_B  = ("Segoe UI", 10)
FONT_S  = ("Segoe UI", 9)


def _lbl(parent, text, font=FONT_B, color=TEXT, **kw):
    return tk.Label(parent, text=text, font=font, fg=color,
                    bg=parent.cget("bg"), **kw)


def _entry(parent, var, width=28):
    return tk.Entry(
        parent, textvariable=var, width=width,
        font=FONT_B, fg=TEXT, bg=BG3, insertbackground=TEXT,
        relief="flat", bd=0, highlightthickness=1,
        highlightcolor=ACCENT, highlightbackground=BORDER,
    )


def _btn(parent, text, cmd, color=ACCENT, hover=ACCENT2, width=18):
    b = tk.Button(
        parent, text=text, command=cmd, font=FONT_B,
        fg=TEXT, bg=color, activeforeground=TEXT, activebackground=hover,
        relief="raised", bd=1, cursor="hand2", padx=12, pady=7, width=width,
    )
    b.bind("<Enter>", lambda e: b.config(bg=hover))
    b.bind("<Leave>", lambda e: b.config(bg=color))
    return b


def _card(parent, **kw):
    return tk.Frame(parent, bg=BG2, padx=16, pady=12,
                    highlightthickness=1, highlightbackground=BORDER, **kw)


def _section(parent, title):
    f = tk.Frame(parent, bg=BG, padx=24, pady=8)
    f.pack(fill="x")
    _lbl(f, title, font=FONT_H).pack(anchor="w", pady=(0, 6))
    return f


def _make_tray_icon():
    try:
        from PIL import Image, ImageDraw
        img  = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.ellipse([4, 4, 60, 60], fill="#4F8EF7")
        draw.rectangle([20, 28, 44, 36], fill="white")
        draw.rectangle([28, 20, 36, 44], fill="white")
        return img
    except ImportError:
        return None


# ── Pestana base con scroll ───────────────────────────────────────────────────
class ScrollTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        sb     = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = tk.Frame(canvas, bg=BG)
        win_id = canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        self.inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
            lambda e: canvas.itemconfig(win_id, width=e.width))
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))


# ── App principal ─────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self, start_hidden: bool = False):
        super().__init__()
        self._config     = cfg.load()
        self._tray_icon  = None
        self._hidden     = False
        self._start_hidden = start_hidden

        self._setup_window()
        self._init_vars()
        self._build_ui()
        self._load_values()
        self._refresh_status()
        self._start_systray()

        self.protocol("WM_DELETE_WINDOW", self._hide_to_tray)

        if start_hidden:
            self.after(100, self.withdraw)

    def _setup_window(self):
        self.title("Outlook Archiver")
        self.configure(bg=BG)
        self.resizable(True, True)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h   = min(620, sw - 40), min(680, sh - 80)
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{max(20,(sh-h)//2)}")
        self.minsize(560, 420)
        try:
            self.iconbitmap(Path(__file__).parent / "icon.ico")
        except Exception:
            pass

    def _init_vars(self):
        # Archivado
        self.threshold_var  = tk.StringVar(self)
        self.pst_dir_var    = tk.StringVar(self)
        self.pst_max_var    = tk.StringVar(self)
        self.freq_var       = tk.StringVar(self)
        self.day_var        = tk.StringVar(self)
        self.hour_var       = tk.StringVar(self)
        self.minute_var     = tk.StringVar(self)
        self.shutdown_var   = tk.BooleanVar(self)
        # OneDrive
        self.onedrive_var     = tk.BooleanVar(self)
        self.onedrive_sub_var = tk.StringVar(self)
        # Sistema
        self.autostart_var  = tk.BooleanVar(self)
        self.silent_var     = tk.BooleanVar(self)

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Header fijo
        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x", side="top")
        hdr = tk.Frame(self, bg=BG, padx=24, pady=12)
        hdr.pack(fill="x", side="top")
        _lbl(hdr, "Outlook Archiver", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        _lbl(hdr, "Archivado automatico por año con PST rotativo",
             color=TEXT2, font=FONT_S).pack(anchor="w")
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", side="top")

        # Barra de guardado fija abajo
        self._build_save_bar()
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", side="bottom")

        # Notebook con pestanas
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("TNotebook",        background=BG, borderwidth=0)
        style.configure("TNotebook.Tab",    background=BG3, foreground=TEXT2,
                        padding=[14, 6], font=FONT_B)
        style.map("TNotebook.Tab",
                  background=[("selected", BG2)],
                  foreground=[("selected", TEXT)])

        self._nb = ttk.Notebook(self)
        self._nb.pack(fill="both", expand=True, side="top")

        self._tab_estado    = ScrollTab(self._nb)
        self._tab_archivado = ScrollTab(self._nb)
        self._tab_onedrive  = ScrollTab(self._nb)
        self._tab_sistema   = ScrollTab(self._nb)

        self._nb.add(self._tab_estado,    text="  Estado  ")
        self._nb.add(self._tab_archivado, text="  Archivado  ")
        self._nb.add(self._tab_onedrive,  text="  OneDrive  ")
        self._nb.add(self._tab_sistema,   text="  Sistema  ")

        self._build_tab_estado()
        self._build_tab_archivado()
        self._build_tab_onedrive()
        self._build_tab_sistema()

    def _build_save_bar(self):
        bar = tk.Frame(self, bg=BG2, padx=24, pady=10)
        bar.pack(fill="x", side="bottom")
        row = tk.Frame(bar, bg=BG2)
        row.pack(fill="x")
        _btn(row, "Guardar y programar", self._save,
             color=ACCENT, hover=ACCENT2, width=20).pack(side="left")
        _btn(row, "Archivar ahora", self._run_now,
             color=SUCCESS, hover="#2EA865", width=14).pack(side="left", padx=(8, 0))
        self.progress = ttk.Progressbar(bar, mode="indeterminate", length=400)
        self.msg_lbl  = _lbl(bar, "", color=TEXT2, font=FONT_S)
        self.msg_lbl.pack(anchor="w", pady=(6, 0))

    # ── Pestana Estado ────────────────────────────────────────────────────────
    def _build_tab_estado(self):
        p = self._tab_estado.inner

        # Tarjeta estado general
        s = _section(p, "Estado del sistema")
        card = _card(s)
        card.pack(fill="x")

        r0 = tk.Frame(card, bg=BG2)
        r0.pack(fill="x")
        _lbl(r0, "Tarea programada:", color=TEXT2, font=FONT_S).pack(side="left")
        self.status_dot = tk.Label(r0, text="*", font=("Segoe UI", 13), bg=BG2, fg=TEXT2)
        self.status_dot.pack(side="right")
        self.status_lbl = _lbl(card, "Comprobando...", color=TEXT2, font=FONT_S)
        self.status_lbl.pack(anchor="w", pady=(2, 8))

        for label, attr in [
            ("Inicio con Windows:", "autostart_status_lbl"),
            ("OneDrive:",           "onedrive_status_lbl"),
            ("Tamanio buzon:",      "size_lbl"),
        ]:
            row = tk.Frame(card, bg=BG2)
            row.pack(fill="x", pady=2)
            _lbl(row, label, color=TEXT2, font=FONT_S).pack(side="left")
            lbl = _lbl(row, "-", color=TEXT2, font=("Segoe UI", 9, "bold"))
            lbl.pack(side="left", padx=(6, 0))
            setattr(self, attr, lbl)

        # Tarjeta proximo archivado
        s2 = _section(p, "Proximo archivado")
        card2 = _card(s2)
        card2.pack(fill="x")

        from archiver import compute_cutoff_date, compute_archive_year
        today    = date.today()
        cutoff   = compute_cutoff_date(today)
        arc_year = compute_archive_year(cutoff)

        for label, attr, default in [
            ("Archiva hasta (exclusive):", "cutoff_lbl",    cutoff.strftime("%d/%m/%Y")),
            ("PST activo:",                "pst_dest_lbl",  f"Archivo {arc_year}.pst"),
            ("Tamanio PST activo:",        "pst_size_lbl",  "-"),
        ]:
            row = tk.Frame(card2, bg=BG2)
            row.pack(fill="x", pady=2)
            _lbl(row, label, color=TEXT2, font=FONT_S).pack(side="left")
            lbl = _lbl(row, default, color=ACCENT, font=("Segoe UI", 9, "bold"))
            lbl.pack(side="left", padx=(6, 0))
            setattr(self, attr, lbl)

        # Botones rapidos
        s3 = _section(p, "Acciones rapidas")
        r = tk.Frame(s3, bg=BG)
        r.pack(fill="x")
        _btn(r, "Ver log", self._open_log, color=BG3, hover=BORDER, width=12).pack(side="left")
        _btn(r, "Abrir carpeta PST", self._open_pst_folder,
             color=BG3, hover=BORDER, width=16).pack(side="left", padx=(8, 0))

        tk.Frame(p, bg=BG, height=12).pack(fill="x")

    # ── Pestana Archivado ─────────────────────────────────────────────────────
    def _build_tab_archivado(self):
        p = self._tab_archivado.inner

        s = _section(p, "Configuracion de archivado")
        card = _card(s)
        card.pack(fill="x")

        self._form_row(card, "Archivar cuando el buzon supere (GB):", self.threshold_var)
        self._form_row(card, "Limite de tamanio por PST (GB):", self.pst_max_var)

        # Carpeta PST
        pf = tk.Frame(card, bg=BG2)
        pf.pack(fill="x", pady=4)
        _lbl(pf, "Carpeta de archivos PST:", color=TEXT2, font=FONT_S).pack(anchor="w")
        pr = tk.Frame(pf, bg=BG2)
        pr.pack(fill="x", pady=(4, 0))
        _entry(pr, self.pst_dir_var, width=32).pack(side="left", fill="x", expand=True)
        tk.Button(pr, text="...", command=self._browse_pst_dir,
                  font=FONT_B, fg=TEXT, bg=BG3, activeforeground=TEXT,
                  activebackground=BORDER, relief="raised", bd=1,
                  cursor="hand2", padx=8, pady=4).pack(side="left", padx=(6, 0))

        # Frecuencia
        s2 = _section(p, "Programacion")
        card2 = _card(s2)
        card2.pack(fill="x")

        ff = tk.Frame(card2, bg=BG2)
        ff.pack(fill="x", pady=(0, 6))
        _lbl(ff, "Frecuencia:", color=TEXT2, font=FONT_S).pack(anchor="w")
        fr = tk.Frame(card2, bg=BG2)
        fr.pack(anchor="w")
        for text, val in [("Diaria", "daily"), ("Semanal", "weekly"), ("Mensual (dia 1)", "monthly")]:
            tk.Radiobutton(fr, text=text, variable=self.freq_var, value=val,
                           font=FONT_B, fg=TEXT, bg=BG2,
                           activeforeground=TEXT, activebackground=BG2,
                           selectcolor=BG3, command=self._toggle_day_gui).pack(side="left", padx=(0, 14))

        self.day_frame_gui = tk.Frame(card2, bg=BG2)
        self.day_frame_gui.pack(fill="x", pady=(6, 0))
        _lbl(self.day_frame_gui, "Dia de la semana:", color=TEXT2, font=FONT_S).pack(side="left")
        for lbl, val in [("Lun","MON"),("Mar","TUE"),("Mie","WED"),("Jue","THU"),
                          ("Vie","FRI"),("Sab","SAT"),("Dom","SUN")]:
            tk.Radiobutton(self.day_frame_gui, text=lbl, variable=self.day_var, value=val,
                           font=FONT_S, fg=TEXT, bg=BG2,
                           activeforeground=TEXT, activebackground=BG2,
                           selectcolor=BG3).pack(side="left", padx=2)

        # Hora
        tf = tk.Frame(card2, bg=BG2)
        tf.pack(fill="x", pady=(10, 0))
        _lbl(tf, "Hora de ejecucion (HH : MM):", color=TEXT2, font=FONT_S).pack(anchor="w")
        tr = tk.Frame(tf, bg=BG2)
        tr.pack(anchor="w", pady=(4, 0))
        tk.Entry(tr, textvariable=self.hour_var, width=5,
                 font=FONT_B, fg=TEXT, bg=BG3, insertbackground=TEXT,
                 relief="flat", bd=0, highlightthickness=1,
                 highlightcolor=ACCENT, highlightbackground=BORDER).pack(side="left")
        _lbl(tr, "  :  ", color=TEXT, font=FONT_H).pack(side="left")
        tk.Entry(tr, textvariable=self.minute_var, width=5,
                 font=FONT_B, fg=TEXT, bg=BG3, insertbackground=TEXT,
                 relief="flat", bd=0, highlightthickness=1,
                 highlightcolor=ACCENT, highlightbackground=BORDER).pack(side="left")

        # Apagado
        s3 = _section(p, "Opciones adicionales")
        card3 = _card(s3)
        card3.pack(fill="x")
        tk.Checkbutton(card3,
            text="Apagar el equipo despues de archivar (60 seg de aviso)",
            variable=self.shutdown_var,
            font=FONT_B, fg=TEXT, bg=BG2,
            activeforeground=TEXT, activebackground=BG2,
            selectcolor=BG3, relief="flat", bd=0, cursor="hand2").pack(anchor="w")

        tk.Frame(p, bg=BG, height=12).pack(fill="x")

    def _toggle_day_gui(self):
        state = "normal" if self.freq_var.get() == "weekly" else "disabled"
        for w in self.day_frame_gui.winfo_children():
            try: w.config(state=state)
            except Exception: pass

    # ── Pestana OneDrive ──────────────────────────────────────────────────────
    def _build_tab_onedrive(self):
        p = self._tab_onedrive.inner

        # Estado
        s = _section(p, "Estado de OneDrive")
        card = _card(s)
        card.pack(fill="x")
        od_path = archiver.find_onedrive_path()
        if od_path:
            _lbl(card, "OneDrive detectado correctamente.", color=SUCCESS,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
            _lbl(card, str(od_path), color=TEXT2, font=FONT_S).pack(anchor="w", pady=(3, 0))
        else:
            _lbl(card, "OneDrive no encontrado en este equipo.",
                 color=DANGER, font=("Segoe UI", 10, "bold")).pack(anchor="w")
            _lbl(card,
                 f"Se busca la carpeta:\n{os.environ.get('USERPROFILE','~')}\\"
                 f"{archiver.ONEDRIVE_FOLDER_NAME}",
                 color=TEXT2, font=FONT_S).pack(anchor="w", pady=(4, 0))

        # Configuracion
        s2 = _section(p, "Configuracion de respaldo")
        card2 = _card(s2)
        card2.pack(fill="x")

        tk.Checkbutton(card2,
            text="Copiar PST a OneDrive al rotar (cierra Outlook temporalmente)",
            variable=self.onedrive_var,
            font=FONT_B, fg=TEXT, bg=BG2,
            activeforeground=TEXT, activebackground=BG2,
            selectcolor=BG3, relief="flat", bd=0, cursor="hand2").pack(anchor="w")

        # Subcarpeta
        sf = tk.Frame(card2, bg=BG2)
        sf.pack(fill="x", pady=(10, 0))
        _lbl(sf, "Subcarpeta dentro de OneDrive:", color=TEXT2, font=FONT_S).pack(anchor="w")
        sr = tk.Frame(sf, bg=BG2)
        sr.pack(fill="x", pady=(4, 0))
        _entry(sr, self.onedrive_sub_var, width=32).pack(side="left", fill="x", expand=True)
        tk.Button(sr, text="...", command=self._browse_onedrive_sub,
                  font=FONT_B, fg=TEXT, bg=BG3, activeforeground=TEXT,
                  activebackground=BORDER, relief="raised", bd=1,
                  cursor="hand2", padx=8, pady=4).pack(side="left", padx=(6, 0))

        if od_path:
            _lbl(card2,
                 f"Ruta completa: {od_path.name}\\{'{subcarpeta}'}",
                 color=TEXT2, font=("Segoe UI", 8)).pack(anchor="w", pady=(6, 0))

        tk.Frame(p, bg=BG, height=12).pack(fill="x")

    # ── Pestana Sistema ───────────────────────────────────────────────────────
    def _build_tab_sistema(self):
        p = self._tab_sistema.inner

        # Autostart
        s = _section(p, "Inicio con Windows")
        card = _card(s)
        card.pack(fill="x")
        tk.Checkbutton(card,
            text="Iniciar automaticamente con Windows",
            variable=self.autostart_var,
            font=FONT_B, fg=TEXT, bg=BG2,
            activeforeground=TEXT, activebackground=BG2,
            selectcolor=BG3, relief="flat", bd=0, cursor="hand2").pack(anchor="w")
        tk.Checkbutton(card,
            text="Iniciar en bandeja sin mostrar ventana",
            variable=self.silent_var,
            font=FONT_B, fg=TEXT, bg=BG2,
            activeforeground=TEXT, activebackground=BG2,
            selectcolor=BG3, relief="flat", bd=0, cursor="hand2").pack(anchor="w", pady=(6, 0))
        _lbl(card,
             "Con 'Iniciar en bandeja' el programa arranca con el equipo\n"
             "pero no muestra ninguna ventana. Se accede desde el icono\n"
             "en la bandeja del sistema (esquina inferior derecha).",
             color=TEXT2, font=("Segoe UI", 8)).pack(anchor="w", pady=(6, 0))

        # Tarea programada
        s2 = _section(p, "Tarea programada")
        card2 = _card(s2)
        card2.pack(fill="x")
        _lbl(card2, f"Nombre: OutlookArchiverTask", color=TEXT, font=FONT_S).pack(anchor="w")
        r = tk.Frame(card2, bg=BG2)
        r.pack(fill="x", pady=(8, 0))
        _btn(r, "Desactivar tarea", self._remove_task,
             color=BG3, hover=BORDER, width=16).pack(side="left")
        _btn(r, "Reconfigurar wizard", self._rerun_wizard,
             color=BG3, hover=BORDER, width=18).pack(side="left", padx=(8, 0))

        # Nota admin pantalla bloqueada
        s3 = _section(p, "Ejecucion con pantalla bloqueada")
        card3 = _card(s3)
        card3.pack(fill="x")
        nota = (
            "Para que el archivado funcione con la pantalla bloqueada,\n"
            "un administrador debe modificar la tarea manualmente:\n\n"
            "1. Abrir Programador de tareas de Windows\n"
            "2. Buscar 'OutlookArchiverTask'\n"
            "3. Propiedades -> General\n"
            "4. Seleccionar 'Ejecutar tanto si el usuario inicio sesion\n"
            "   como si no'\n"
            "5. Ingresar la contrasena del usuario cuando se solicite\n\n"
            "Nota: la sesion debe estar activa (pantalla bloqueada, no\n"
            "cerrada). Outlook no puede abrirse sin sesion activa."
        )
        _lbl(card3, nota, color=TEXT2, font=("Segoe UI", 8),
             justify="left").pack(anchor="w")

        # Desinstalar
        s4 = _section(p, "Desinstalacion")
        card4 = _card(s4)
        card4.pack(fill="x")
        _lbl(card4,
             "Elimina la tarea programada, el inicio con Windows\n"
             "y la configuracion guardada. Los archivos PST no se tocan.",
             color=TEXT2, font=FONT_S).pack(anchor="w", pady=(0, 8))
        _btn(card4, "Desinstalar Outlook Archiver", self._uninstall,
             color=DANGER, hover="#B83C3C", width=26).pack(anchor="w")

        tk.Frame(p, bg=BG, height=12).pack(fill="x")

    # ── Helpers de layout ─────────────────────────────────────────────────────
    def _form_row(self, parent, label_text, var):
        f = tk.Frame(parent, bg=BG2)
        f.pack(fill="x", pady=4)
        _lbl(f, label_text, color=TEXT2, font=FONT_S).pack(anchor="w")
        _entry(f, var, width=12).pack(anchor="w", pady=(4, 0))

    # ── Carga y recogida de valores ───────────────────────────────────────────
    def _load_values(self):
        d = self._config
        self.threshold_var.set(d.get("threshold_gb", 3.0))
        self.pst_dir_var.set(d.get("pst_base_dir", ""))
        self.pst_max_var.set(d.get("pst_max_gb", 47.0))
        self.freq_var.set(d.get("schedule_freq", "daily"))
        self.day_var.set(d.get("schedule_day", "MON"))
        self.hour_var.set(f"{int(d.get('schedule_hour', 20)):02d}")
        self.minute_var.set(f"{int(d.get('schedule_minute', 0)):02d}")
        self.shutdown_var.set(d.get("shutdown_after", False))
        self.onedrive_var.set(d.get("onedrive_backup", False))
        self.onedrive_sub_var.set(d.get("onedrive_subpath", "Respaldo Correo"))
        self.autostart_var.set(d.get("autostart", True))
        self.silent_var.set(d.get("autostart_silent", True))
        self._toggle_day_gui()

    def _collect(self):
        d = self._config.copy()
        try:
            d["threshold_gb"]    = float(self.threshold_var.get())
            d["pst_base_dir"]    = self.pst_dir_var.get().strip()
            d["pst_max_gb"]      = float(self.pst_max_var.get())
            d["schedule_freq"]   = self.freq_var.get()
            d["schedule_day"]    = self.day_var.get()
            d["schedule_hour"]   = int(self.hour_var.get())
            d["schedule_minute"] = int(self.minute_var.get())
            d["shutdown_after"]  = self.shutdown_var.get()
            d["onedrive_backup"]  = self.onedrive_var.get()
            d["onedrive_subpath"] = self.onedrive_sub_var.get().strip()
            d["autostart"]        = self.autostart_var.get()
            d["autostart_silent"] = self.silent_var.get()
        except ValueError as e:
            raise ValueError(f"Valor invalido: {e}")
        return d

    # ── Acciones ──────────────────────────────────────────────────────────────
    def _save(self):
        try:
            data = self._collect()
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return
        cfg.save(data)
        self._config = data
        ok = scheduler.register_task(data)
        if data.get("autostart"):
            startup.enable_autostart(silent=data.get("autostart_silent", True))
        else:
            startup.disable_autostart()
        color = SUCCESS if ok else WARNING
        msg   = "Configuracion guardada y tarea programada." if ok \
                else "Configuracion guardada, pero hubo un error al programar la tarea."
        self._set_msg(msg, color)
        self._refresh_status()

    def _run_now(self):
        try:
            data = self._collect()
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return
        cfg.save(data)
        self._config = data
        self._set_msg("Archivando correos...", TEXT2)
        self.progress.pack(anchor="w", pady=(6, 0))
        self.progress.start(10)
        threading.Thread(
            target=lambda: self.after(0, lambda: self._on_done(archiver.run_archive(data))),
            daemon=True,
        ).start()

    def _on_done(self, result):
        self.progress.stop()
        self.progress.pack_forget()
        color = SUCCESS if result["status"] == "ok" else (
                WARNING if result["status"] == "warning" else DANGER)
        self._set_msg(result["message"], color)
        self._refresh_status()

    def _remove_task(self):
        if messagebox.askyesno("Desactivar tarea",
                               "Eliminar la tarea del Programador de Windows?"):
            ok = scheduler.remove_task()
            self._set_msg("Tarea eliminada." if ok else "No se pudo eliminar.",
                          SUCCESS if ok else WARNING)
            self._refresh_status()

    def _open_log(self):
        lp = cfg.get_log_path(self._config)
        if lp.exists():
            os.startfile(lp)
        else:
            messagebox.showinfo("Log", f"No existe aun el log en:\n{lp}")

    def _open_pst_folder(self):
        base = self._config.get("pst_base_dir", "")
        if base and Path(base).exists():
            os.startfile(base)
        else:
            messagebox.showinfo("Carpeta PST",
                                f"La carpeta no existe aun:\n{base}")

    def _browse_pst_dir(self):
        path = filedialog.askdirectory(title="Carpeta para archivos PST")
        if path:
            self.pst_dir_var.set(path)

    def _browse_onedrive_sub(self):
        od = archiver.find_onedrive_path()
        initial = str(od) if od else os.environ.get("USERPROFILE", "")
        path = filedialog.askdirectory(
            title="Subcarpeta dentro de OneDrive",
            initialdir=initial,
        )
        if path and od:
            # Guardar solo la ruta relativa dentro de OneDrive
            try:
                rel = Path(path).relative_to(od)
                self.onedrive_sub_var.set(str(rel))
            except ValueError:
                self.onedrive_sub_var.set(path)
        elif path:
            self.onedrive_sub_var.set(path)

    def _rerun_wizard(self):
        import wizard
        self.withdraw()
        result = wizard.run_wizard()
        if result:
            cfg.save(result)
            self._config = result
            self._load_values()
            scheduler.register_task(result)
            startup.enable_autostart(result.get("autostart_silent", True)) \
                if result.get("autostart") else startup.disable_autostart()
            self._refresh_status()
            self._set_msg("Reconfigurado correctamente.", SUCCESS)
        self.deiconify()

    def _uninstall(self):
        if not messagebox.askyesno("Desinstalar",
            "Esto eliminara:\n"
            "  - Tarea OutlookArchiverTask\n"
            "  - Inicio automatico con Windows\n"
            "  - Configuracion guardada\n\n"
            "Los archivos PST NO se eliminan.\n\nContinuar?"):
            return
        scheduler.remove_task()
        startup.disable_autostart()
        try:
            cfg.get_config_path().unlink(missing_ok=True)
        except Exception:
            pass
        messagebox.showinfo("Desinstalado",
            "Outlook Archiver desinstalado.\n"
            "Puedes eliminar el .exe manualmente.")
        if self._tray_icon:
            self._tray_icon.stop()
        self.destroy()

    def _refresh_status(self):
        # Tarea
        exists = scheduler.task_exists()
        self.status_dot.config(fg=SUCCESS if exists else DANGER)
        self.status_lbl.config(
            text="Activa — OutlookArchiverTask" if exists
                 else "Sin tarea — haz clic en 'Guardar y programar'",
            fg=SUCCESS if exists else DANGER,
        )
        # Autostart
        self.autostart_status_lbl.config(
            text="Activo" if startup.autostart_enabled() else "Inactivo",
            fg=SUCCESS if startup.autostart_enabled() else TEXT2,
        )
        # OneDrive
        od = archiver.find_onedrive_path()
        self.onedrive_status_lbl.config(
            text="Configurado" if od else "No encontrado",
            fg=SUCCESS if od else DANGER,
        )
        # Buzon
        try:
            sz  = archiver.get_ost_size_gb()
            thr = float(self._config.get("threshold_gb", 3))
            self.size_lbl.config(
                text=f"{sz:.2f} GB",
                fg=DANGER if sz >= thr else SUCCESS,
            )
        except Exception:
            self.size_lbl.config(text="N/D", fg=TEXT2)
        # PST activo
        try:
            today    = date.today()
            cutoff   = archiver.compute_cutoff_date(today)
            arc_year = archiver.compute_archive_year(cutoff)
            base_dir = self._config.get("pst_base_dir", "")
            pst_max  = float(self._config.get("pst_max_gb", 47))
            pst_p, _ = archiver.get_active_pst_path(base_dir, arc_year, pst_max)
            pst_sz   = archiver.get_pst_size_gb(str(pst_p))
            self.pst_dest_lbl.config(text=pst_p.stem)
            self.pst_size_lbl.config(
                text=f"{pst_sz:.2f} GB",
                fg=DANGER if pst_sz >= pst_max * 0.9 else TEXT,
            )
        except Exception:
            pass

    def _set_msg(self, text, color=TEXT2):
        self.msg_lbl.config(text=text, fg=color)

    # ── Systray ───────────────────────────────────────────────────────────────
    def _start_systray(self):
        try:
            import pystray
        except ImportError:
            logger.warning("pystray no instalado — bandeja no disponible.")
            return
        icon_img = _make_tray_icon()
        if icon_img is None:
            return
        menu = pystray.Menu(
            pystray.MenuItem("Abrir configuracion", self._show_from_tray, default=True),
            pystray.MenuItem("Archivar ahora",      self._tray_archive_now),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Salir",               self._quit_app),
        )
        self._tray_icon = pystray.Icon(
            "OutlookArchiver", icon_img, "Outlook Archiver", menu)
        threading.Thread(target=self._tray_icon.run, daemon=True).start()

    def _show_from_tray(self, *_):
        self.after(0, self._do_show)

    def _do_show(self):
        self.deiconify()
        self.lift()
        self.focus_force()
        self._hidden = False

    def _hide_to_tray(self):
        if self._tray_icon:
            self.withdraw()
            self._hidden = True
        else:
            if messagebox.askyesno("Cerrar",
                "Cerrar? La tarea programada seguira activa."):
                self._quit_app()

    def _tray_archive_now(self, *_):
        threading.Thread(
            target=lambda: archiver.run_archive(self._config),
            daemon=True).start()

    def _quit_app(self, *_):
        if self._tray_icon:
            self._tray_icon.stop()
        self.after(0, self.destroy)


def run(start_hidden: bool = False):
    app = App(start_hidden=start_hidden)
    app.mainloop()
