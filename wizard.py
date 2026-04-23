"""
wizard.py — Wizard de primera ejecución
Se muestra cuando no existe configuración previa.
Guía al usuario en 3 pasos: bienvenida → configuración → finalizar.
"""

import tkinter as tk
from tkinter import filedialog
from pathlib import Path

# ── Paleta (misma que gui.py) ─────────────────────────────────────────────────
BG      = "#1A1D23"
BG2     = "#22262F"
BG3     = "#2C3140"
ACCENT  = "#4F8EF7"
ACCENT2 = "#3A6ED4"
SUCCESS = "#3DBE7A"
TEXT    = "#E8EAF0"
TEXT2   = "#9AA0B2"
BORDER  = "#353A47"
FONT_H  = ("Segoe UI", 13, "bold")
FONT_B  = ("Segoe UI", 10)
FONT_S  = ("Segoe UI", 9)
FONT_BIG = ("Segoe UI", 22, "bold")


def _label(parent, text, font=FONT_B, color=TEXT, **kw):
    return tk.Label(parent, text=text, font=font, fg=color, bg=parent["bg"], **kw)


def _entry(parent, textvariable, width=28):
    return tk.Entry(
        parent, textvariable=textvariable, width=width,
        font=FONT_B, fg=TEXT, bg=BG3, insertbackground=TEXT,
        relief="flat", bd=0, highlightthickness=1,
        highlightcolor=ACCENT, highlightbackground=BORDER,
    )


def _btn(parent, text, command, color=ACCENT, hover=ACCENT2, width=16):
    b = tk.Button(
        parent, text=text, command=command, font=FONT_B,
        fg=TEXT, bg=color, activeforeground=TEXT, activebackground=hover,
        relief="flat", bd=0, cursor="hand2", padx=14, pady=8, width=width,
    )
    b.bind("<Enter>", lambda e: b.config(bg=hover))
    b.bind("<Leave>", lambda e: b.config(bg=color))
    return b


class SetupWizard(tk.Toplevel):
    """
    Wizard de 3 pasos. Retorna la configuración completa al cerrarse.
    Acceder al resultado via wizard.result (dict o None si se canceló).
    """

    STEPS = ["Bienvenida", "Configuración", "Listo"]

    def __init__(self, parent=None):
        if parent:
            super().__init__(parent)
        else:
            # Ventana independiente si se llama sin parent
            self._root = tk.Tk()
            self._root.withdraw()
            super().__init__(self._root)

        self.result = None
        self._step = 0

        self._setup_window()
        self._init_vars()
        self._build()
        self._show_step(0)

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.grab_set()

    def _setup_window(self):
        self.title("Outlook Archiver — Configuración inicial")
        self.configure(bg=BG)
        self.resizable(False, False)
        w, h = 520, 520
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _init_vars(self):
        self.threshold_var = tk.StringVar(value="4")
        self.months_var    = tk.StringVar(value="12")
        self.hour_var      = tk.StringVar(value="20")
        self.minute_var    = tk.StringVar(value="00")
        self.pst_var       = tk.StringVar(value=str(Path.home() / "Documents" / "OutlookArchivo.pst"))
        self.autostart_var = tk.BooleanVar(value=True)

    # ── Estructura principal ─────────────────────────────────────────────────
    def _build(self):
        # Barra de color superior
        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x")

        # Indicador de pasos
        self.steps_frame = tk.Frame(self, bg=BG, pady=14)
        self.steps_frame.pack(fill="x")
        self._build_step_indicators()

        # Separador
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        # Contenedor de pasos (intercambiable)
        self.content_frame = tk.Frame(self, bg=BG, padx=40, pady=24)
        self.content_frame.pack(fill="both", expand=True)

        # Botones de navegación
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")
        nav = tk.Frame(self, bg=BG, padx=40, pady=16)
        nav.pack(fill="x")

        self.back_btn = _btn(nav, "← Atrás", self._prev, color=BG3, hover=BORDER, width=10)
        self.back_btn.pack(side="left")

        self.next_btn = _btn(nav, "Siguiente →", self._next, width=14)
        self.next_btn.pack(side="right")

    def _build_step_indicators(self):
        for widget in self.steps_frame.winfo_children():
            widget.destroy()

        inner = tk.Frame(self.steps_frame, bg=BG)
        inner.pack()

        for i, name in enumerate(self.STEPS):
            if i > 0:
                tk.Frame(inner, bg=BORDER, width=40, height=1).pack(side="left", pady=12)

            dot_color = ACCENT if i <= self._step else BORDER
            txt_color = TEXT if i <= self._step else TEXT2

            col = tk.Frame(inner, bg=BG)
            col.pack(side="left", padx=8)

            tk.Label(col, text="●", font=("Segoe UI", 14), fg=dot_color, bg=BG).pack()
            tk.Label(col, text=name, font=FONT_S, fg=txt_color, bg=BG).pack()

    # ── Pasos ────────────────────────────────────────────────────────────────
    def _clear_content(self):
        for w in self.content_frame.winfo_children():
            w.destroy()

    def _show_step(self, step: int):
        self._step = step
        self._build_step_indicators()
        self._clear_content()

        if step == 0:
            self._step_welcome()
            self.back_btn.config(state="disabled")
            self.next_btn.config(text="Comenzar →")
        elif step == 1:
            self._step_config()
            self.back_btn.config(state="normal")
            self.next_btn.config(text="Finalizar ✔")
        elif step == 2:
            self._step_done()
            self.back_btn.config(state="disabled")
            self.next_btn.config(text="Abrir configuración")

    def _step_welcome(self):
        f = self.content_frame
        _label(f, "📦", font=("Segoe UI", 36), color=ACCENT).pack(pady=(10, 6))
        _label(f, "Outlook Archiver", font=FONT_BIG, color=TEXT).pack()
        _label(f, "Archivado automático de correos", color=TEXT2, font=FONT_S).pack(pady=(4, 24))

        card = tk.Frame(f, bg=BG2, padx=20, pady=16,
                        highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="x")

        items = [
            ("📏", "Monitorea el tamaño de tu buzón"),
            ("📁", "Mueve correos antiguos a un archivo .pst"),
            ("⏰", "Corre automáticamente según el horario que elijas"),
            ("🚀", "Se inicia con Windows, sin intervención"),
        ]
        for icon, text in items:
            row = tk.Frame(card, bg=BG2)
            row.pack(fill="x", pady=3)
            _label(row, icon, font=("Segoe UI", 13), color=ACCENT).pack(side="left", padx=(0, 10))
            _label(row, text, color=TEXT, font=FONT_B).pack(side="left")

    def _step_config(self):
        f = self.content_frame
        _label(f, "Configuración", font=FONT_H, color=TEXT).pack(anchor="w", pady=(0, 16))

        def field(lbl, var, width=10):
            row = tk.Frame(f, bg=BG)
            row.pack(fill="x", pady=5)
            _label(row, lbl, color=TEXT2, font=FONT_S).pack(anchor="w")
            _entry(row, var, width=width).pack(anchor="w", pady=(4, 0))

        field("Archivar cuando el buzón supere (GB):", self.threshold_var)
        field("Correos más antiguos que (meses):", self.months_var)

        # Hora
        time_row_label = tk.Frame(f, bg=BG)
        time_row_label.pack(fill="x", pady=(5, 0))
        _label(time_row_label, "Ejecutar diariamente a las (HH:MM):", color=TEXT2, font=FONT_S).pack(anchor="w")

        time_row = tk.Frame(f, bg=BG)
        time_row.pack(anchor="w", pady=(4, 0))
        _entry(time_row, self.hour_var, width=5).pack(side="left")
        _label(time_row, ":", color=TEXT, font=FONT_H).pack(side="left", padx=4)
        _entry(time_row, self.minute_var, width=5).pack(side="left")

        # Ruta .pst
        pst_lbl = tk.Frame(f, bg=BG)
        pst_lbl.pack(fill="x", pady=(10, 0))
        _label(pst_lbl, "Ruta del archivo .pst:", color=TEXT2, font=FONT_S).pack(anchor="w")

        pst_row = tk.Frame(f, bg=BG)
        pst_row.pack(fill="x", pady=(4, 0))
        _entry(pst_row, self.pst_var, width=32).pack(side="left", fill="x", expand=True)
        browse = tk.Button(
            pst_row, text="…", command=self._browse_pst,
            font=FONT_B, fg=TEXT, bg=BG3, activeforeground=TEXT,
            activebackground=BORDER, relief="flat", bd=0,
            cursor="hand2", padx=10, pady=5,
        )
        browse.pack(side="left", padx=(6, 0))

        # Checkbox inicio con Windows
        chk_row = tk.Frame(f, bg=BG)
        chk_row.pack(fill="x", pady=(14, 0))
        tk.Checkbutton(
            chk_row, text="Iniciar automáticamente con Windows",
            variable=self.autostart_var,
            font=FONT_B, fg=TEXT, bg=BG,
            activeforeground=TEXT, activebackground=BG,
            selectcolor=BG3, relief="flat", bd=0, cursor="hand2",
        ).pack(anchor="w")

    def _step_done(self):
        f = self.content_frame
        _label(f, "✔", font=("Segoe UI", 42), color=SUCCESS).pack(pady=(20, 8))
        _label(f, "¡Todo listo!", font=FONT_BIG, color=TEXT).pack()
        _label(f, "La herramienta está configurada y activa.", color=TEXT2, font=FONT_S).pack(pady=(6, 28))

        conf = self._build_config()
        card = tk.Frame(f, bg=BG2, padx=20, pady=14,
                        highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="x")

        resumen = [
            ("Umbral", f"{conf['threshold_gb']} GB"),
            ("Antigüedad", f"{conf['months_old']} meses"),
            ("Horario diario", f"{conf['schedule_hour']:02d}:{conf['schedule_minute']:02d}"),
            ("Archivo .pst", Path(conf['pst_path']).name),
            ("Inicio con Windows", "Sí" if conf['autostart'] else "No"),
        ]
        for k, v in resumen:
            row = tk.Frame(card, bg=BG2)
            row.pack(fill="x", pady=2)
            _label(row, f"{k}:", color=TEXT2, font=FONT_S).pack(side="left")
            _label(row, v, color=TEXT, font=("Segoe UI", 9, "bold")).pack(side="left", padx=(6, 0))

    # ── Navegación ───────────────────────────────────────────────────────────
    def _next(self):
        if self._step == 0:
            self._show_step(1)
        elif self._step == 1:
            if not self._validate():
                return
            self.result = self._build_config()
            self._show_step(2)
        elif self._step == 2:
            self.destroy()

    def _prev(self):
        if self._step > 0:
            self._show_step(self._step - 1)

    def _on_close(self):
        self.result = None
        self.destroy()

    # ── Helpers ──────────────────────────────────────────────────────────────
    def _validate(self) -> bool:
        try:
            gb = float(self.threshold_var.get())
            assert 0.5 <= gb <= 50
        except Exception:
            tk.messagebox.showerror("Error", "El umbral en GB debe ser un número entre 0.5 y 50.", parent=self)
            return False
        try:
            m = int(self.months_var.get())
            assert 1 <= m <= 120
        except Exception:
            tk.messagebox.showerror("Error", "Los meses deben ser un número entre 1 y 120.", parent=self)
            return False
        try:
            h = int(self.hour_var.get())
            mm = int(self.minute_var.get())
            assert 0 <= h <= 23 and 0 <= mm <= 59
        except Exception:
            tk.messagebox.showerror("Error", "Hora inválida. Usa formato HH (00–23) y MM (00–59).", parent=self)
            return False
        if not self.pst_var.get().strip():
            tk.messagebox.showerror("Error", "Debes especificar la ruta del archivo .pst.", parent=self)
            return False
        return True

    def _build_config(self) -> dict:
        import os
        from config import CONFIG_DIR, DEFAULTS
        return {
            "threshold_gb":    float(self.threshold_var.get()),
            "months_old":      int(self.months_var.get()),
            "schedule_hour":   int(self.hour_var.get()),
            "schedule_minute": int(self.minute_var.get()),
            "pst_path":        self.pst_var.get().strip(),
            "autostart":       self.autostart_var.get(),
            "notify_email":    "",
            "log_path":        str(CONFIG_DIR / "archiver.log"),
            "enabled":         True,
            "setup_done":      True,
        }

    def _browse_pst(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".pst",
            filetypes=[("Archivo de datos de Outlook", "*.pst")],
            title="Seleccionar ubicación del archivo .pst",
            parent=self,
        )
        if path:
            self.pst_var.set(path)


def run_wizard() -> dict | None:
    """
    Lanza el wizard en modo standalone.
    Retorna la configuración si el usuario completó el wizard, None si canceló.
    """
    root = tk.Tk()
    root.withdraw()
    wizard = SetupWizard(root)
    root.wait_window(wizard)
    root.destroy()
    return wizard.result
