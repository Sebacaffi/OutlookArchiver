"""
wizard.py - Wizard de primera ejecucion (tk.Tk standalone, NO Toplevel)
3 pasos: Bienvenida -> Configuracion -> Listo
"""

import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import date

BG       = "#1A1D23"
BG2      = "#22262F"
BG3      = "#2C3140"
ACCENT   = "#4F8EF7"
ACCENT2  = "#3A6ED4"
SUCCESS  = "#3DBE7A"
TEXT     = "#E8EAF0"
TEXT2    = "#9AA0B2"
BORDER   = "#353A47"
FONT_H   = ("Segoe UI", 13, "bold")
FONT_B   = ("Segoe UI", 10)
FONT_S   = ("Segoe UI", 9)
FONT_BIG = ("Segoe UI", 18, "bold")


def _lbl(parent, text, font=FONT_B, color=TEXT, **kw):
    return tk.Label(parent, text=text, font=font, fg=color,
                    bg=parent.cget("bg"), **kw)


def _entry(parent, var, width=14):
    return tk.Entry(
        parent, textvariable=var, width=width,
        font=FONT_B, fg=TEXT, bg=BG3, insertbackground=TEXT,
        relief="flat", bd=4,
        highlightthickness=1, highlightcolor=ACCENT, highlightbackground=BORDER,
    )


def _btn(parent, text, cmd, color=ACCENT, hover=ACCENT2, width=14):
    b = tk.Button(
        parent, text=text, command=cmd, font=FONT_B,
        fg=TEXT, bg=color, activeforeground=TEXT, activebackground=hover,
        relief="raised", bd=1, cursor="hand2", padx=12, pady=6, width=width,
    )
    b.bind("<Enter>", lambda e: b.config(bg=hover))
    b.bind("<Leave>", lambda e: b.config(bg=color))
    return b


class SetupWizard(tk.Tk):
    STEPS = ["1. Bienvenida", "2. Configuracion", "3. Listo"]

    def __init__(self):
        super().__init__()
        self.result = None
        self._step  = 0

        # Preview de fecha de corte calculada
        from archiver import compute_cutoff_date, compute_archive_year
        today        = date.today()
        cutoff       = compute_cutoff_date(today)
        arc_year     = compute_archive_year(cutoff)
        self._cutoff_preview = cutoff.strftime("%d/%m/%Y")
        self._year_preview   = arc_year

        # Variables de formulario
        self.threshold_var   = tk.StringVar(self, value="3")
        self.pst_dir_var     = tk.StringVar(
            self, value=str(Path.home() / "Documents" / "ArchivosOutlook"))
        self.pst_max_var     = tk.StringVar(self, value="47")
        self.hour_var        = tk.StringVar(self, value="20")
        self.minute_var      = tk.StringVar(self, value="00")
        self.autostart_var   = tk.BooleanVar(self, value=True)
        self.onedrive_var    = tk.BooleanVar(self, value=False)
        self.onedrive_sub_var = tk.StringVar(self, value="Respaldo Correo")

        self._setup_window()
        self._build_shell()
        self._show_step(0)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _setup_window(self):
        self.title("Outlook Archiver - Configuracion inicial")
        self.configure(bg=BG)
        self.resizable(False, False)
        w, h = 540, 650
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _build_shell(self):
        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x", side="top")

        self.steps_bar = tk.Frame(self, bg=BG, pady=10)
        self.steps_bar.pack(fill="x", side="top")

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", side="top")
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", side="bottom")

        nav = tk.Frame(self, bg=BG, padx=40, pady=12)
        nav.pack(fill="x", side="bottom")

        self.back_btn = _btn(nav, "<- Atras", self._prev,
                             color=BG3, hover=BORDER, width=10)
        self.back_btn.pack(side="left")
        self.next_btn = _btn(nav, "Siguiente ->", self._next, width=16)
        self.next_btn.pack(side="right")

        self.content = tk.Frame(self, bg=BG, padx=40, pady=16)
        self.content.pack(fill="both", expand=True, side="top")

    def _refresh_steps_bar(self):
        for w in self.steps_bar.winfo_children():
            w.destroy()
        inner = tk.Frame(self.steps_bar, bg=BG)
        inner.pack()
        for i, name in enumerate(self.STEPS):
            if i > 0:
                tk.Frame(inner, bg=BORDER, width=36, height=1).pack(
                    side="left", padx=4, pady=8)
            col = tk.Frame(inner, bg=BG)
            col.pack(side="left", padx=6)
            dc = ACCENT if i <= self._step else BORDER
            tc = TEXT   if i <= self._step else TEXT2
            tk.Label(col, text="*", font=("Segoe UI", 11), fg=dc, bg=BG).pack()
            tk.Label(col, text=name, font=FONT_S, fg=tc, bg=BG).pack()

    def _clear_content(self):
        for w in self.content.winfo_children():
            w.destroy()

    def _show_step(self, step):
        self._step = step
        self._refresh_steps_bar()
        self._clear_content()
        if step == 0:
            self._page_welcome()
            self.back_btn.config(state="disabled")
            self.next_btn.config(text="Comenzar ->", bg=ACCENT)
        elif step == 1:
            self._page_config()
            self.back_btn.config(state="normal")
            self.next_btn.config(text="Finalizar", bg=ACCENT)
        elif step == 2:
            self._page_done()
            self.back_btn.config(state="disabled")
            self.next_btn.config(text="Abrir programa", bg=SUCCESS)
        self.update_idletasks()

    # ── Paginas ───────────────────────────────────────────────────────────────
    def _page_welcome(self):
        f = self.content
        _lbl(f, "Outlook Archiver", font=FONT_BIG).pack(pady=(8, 2))
        _lbl(f, "Archivado automatico por año con PST rotativo",
             color=TEXT2, font=FONT_S).pack()

        tk.Frame(f, bg=BORDER, height=1).pack(fill="x", pady=14)

        card = tk.Frame(f, bg=BG2, padx=18, pady=12)
        card.pack(fill="x")
        features = [
            "Archiva correos al PST del año correspondiente",
            f"Hoy archivaria hasta el {self._cutoff_preview} inclusive",
            f"PST de destino: Archivo {self._year_preview}.pst",
            "Rota automaticamente al llegar al limite de 47 GB",
            "Se inicia con Windows y corre en la bandeja del sistema",
        ]
        for feat in features:
            row = tk.Frame(card, bg=BG2)
            row.pack(fill="x", pady=2)
            tk.Label(row, text="OK", font=FONT_S, fg=SUCCESS, bg=BG2).pack(
                side="left", padx=(0, 8))
            _lbl(row, feat, color=TEXT, font=FONT_S).pack(side="left")

        tk.Frame(f, bg=BORDER, height=1).pack(fill="x", pady=14)
        _lbl(f, "Haz clic en 'Comenzar' para configurar.",
             color=TEXT2, font=FONT_S).pack()

    def _page_config(self):
        f = self.content
        _lbl(f, "Configuracion", font=FONT_H).pack(anchor="w", pady=(0, 10))

        # Umbral buzon
        self._field(f, "Archivar cuando el buzon supere (GB):",
                    self.threshold_var, width=8)

        # Directorio base de PSTs
        df = tk.Frame(f, bg=BG)
        df.pack(fill="x", pady=4)
        _lbl(df, "Carpeta donde guardar los archivos PST:",
             color=TEXT2, font=FONT_S).pack(anchor="w")
        dr = tk.Frame(df, bg=BG)
        dr.pack(fill="x", pady=(4, 0))
        _entry(dr, self.pst_dir_var, width=30).pack(side="left", fill="x", expand=True)
        tk.Button(
            dr, text="...", command=self._browse_dir,
            font=FONT_B, fg=TEXT, bg=BG3, activeforeground=TEXT,
            activebackground=BORDER, relief="raised", bd=1,
            cursor="hand2", padx=8, pady=4,
        ).pack(side="left", padx=(6, 0))
        _lbl(df, f"Se creara automaticamente: Archivo {self._year_preview}.pst",
             color=TEXT2, font=("Segoe UI", 8)).pack(anchor="w", pady=(3, 0))

        # Limite PST
        self._field(f, "Limite de tamanio por archivo PST (GB, max recomendado: 47):",
                    self.pst_max_var, width=8)

        # Hora
        tf = tk.Frame(f, bg=BG)
        tf.pack(fill="x", pady=4)
        _lbl(tf, "Ejecutar diariamente a las (HH : MM):",
             color=TEXT2, font=FONT_S).pack(anchor="w")
        tr = tk.Frame(tf, bg=BG)
        tr.pack(anchor="w", pady=(4, 0))
        _entry(tr, self.hour_var, width=5).pack(side="left")
        _lbl(tr, "  :  ", color=TEXT, font=FONT_H).pack(side="left")
        _entry(tr, self.minute_var, width=5).pack(side="left")

        # Autostart
        ck = tk.Frame(f, bg=BG)
        ck.pack(fill="x", pady=(10, 0))
        tk.Checkbutton(
            ck, text="Iniciar automaticamente con Windows",
            variable=self.autostart_var,
            font=FONT_B, fg=TEXT, bg=BG,
            activeforeground=TEXT, activebackground=BG,
            selectcolor=BG3, relief="flat", bd=0, cursor="hand2",
        ).pack(anchor="w")

        # OneDrive backup
        od = tk.Frame(f, bg=BG)
        od.pack(fill="x", pady=(8, 0))
        tk.Checkbutton(
            od, text="Copiar PST a OneDrive al rotar (requiere cerrar Outlook)",
            variable=self.onedrive_var,
            font=FONT_B, fg=TEXT, bg=BG,
            activeforeground=TEXT, activebackground=BG,
            selectcolor=BG3, relief="flat", bd=0, cursor="hand2",
        ).pack(anchor="w")
        self._field(f, "Subcarpeta dentro de OneDrive:", self.onedrive_sub_var, width=24)

    def _page_done(self):
        f = self.content
        _lbl(f, "Configuracion completada", font=FONT_BIG, color=SUCCESS).pack(
            pady=(12, 4))
        _lbl(f, "La herramienta quedara activa en la bandeja del sistema.",
             color=TEXT2, font=FONT_S).pack(pady=(0, 16))

        conf = self._collect()
        card = tk.Frame(f, bg=BG2, padx=18, pady=12)
        card.pack(fill="x")
        rows = [
            ("Umbral buzon",    f"{conf['threshold_gb']} GB"),
            ("Carpeta PST",     Path(conf['pst_base_dir']).name),
            ("PST activo",      f"Archivo {self._year_preview}.pst"),
            ("Limite por PST",  f"{conf['pst_max_gb']} GB"),
            ("Horario",         f"{conf['schedule_hour']:02d}:{conf['schedule_minute']:02d}"),
            ("Archiva hasta",   f"{self._cutoff_preview} (exclusive)"),
            ("Inicio con Win.", "Si" if conf['autostart'] else "No"),
            ("Backup OneDrive", "Si" if conf['onedrive_backup'] else "No"),
        ]
        for k, v in rows:
            row = tk.Frame(card, bg=BG2)
            row.pack(fill="x", pady=2)
            _lbl(row, f"{k}:", color=TEXT2, font=FONT_S).pack(side="left")
            _lbl(row, v, color=TEXT,
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=(8, 0))

    def _field(self, parent, label_text, var, width=10):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="x", pady=4)
        _lbl(f, label_text, color=TEXT2, font=FONT_S).pack(anchor="w")
        _entry(f, var, width=width).pack(anchor="w", pady=(4, 0))

    # ── Navegacion ────────────────────────────────────────────────────────────
    def _next(self):
        if self._step == 0:
            self._show_step(1)
        elif self._step == 1:
            if not self._validate():
                return
            self.result = self._collect()
            self._show_step(2)
        elif self._step == 2:
            self.destroy()

    def _prev(self):
        if self._step > 0:
            self._show_step(self._step - 1)

    def _on_close(self):
        self.result = None
        self.destroy()

    def _validate(self):
        try:
            assert 0.5 <= float(self.threshold_var.get()) <= 50
        except Exception:
            messagebox.showerror("Error", "Umbral GB debe ser entre 0.5 y 50.")
            return False
        if not self.pst_dir_var.get().strip():
            messagebox.showerror("Error", "Debes indicar la carpeta de destino.")
            return False
        try:
            assert 1.0 <= float(self.pst_max_var.get()) <= 50
        except Exception:
            messagebox.showerror("Error", "Limite PST debe ser entre 1 y 50 GB.")
            return False
        try:
            h  = int(self.hour_var.get())
            mm = int(self.minute_var.get())
            assert 0 <= h <= 23 and 0 <= mm <= 59
        except Exception:
            messagebox.showerror("Error", "Hora invalida (HH 00-23, MM 00-59).")
            return False
        return True

    def _collect(self):
        from config import CONFIG_DIR
        return {
            "threshold_gb":    float(self.threshold_var.get()),
            "pst_base_dir":    self.pst_dir_var.get().strip(),
            "pst_max_gb":      float(self.pst_max_var.get()),
            "schedule_hour":   int(self.hour_var.get()),
            "schedule_minute": int(self.minute_var.get()),
            "autostart":       bool(self.autostart_var.get()),
            "onedrive_backup":  bool(self.onedrive_var.get()),
            "onedrive_subpath": self.onedrive_sub_var.get().strip(),
            "notify_email":    "",
            "log_path":        str(CONFIG_DIR / "archiver.log"),
            "enabled":         True,
            "setup_done":      True,
        }

    def _browse_dir(self):
        path = filedialog.askdirectory(title="Seleccionar carpeta para archivos PST")
        if path:
            self.pst_dir_var.set(path)


def run_wizard():
    app = SetupWizard()
    app.mainloop()
    return app.result
