"""
gui.py — Interfaz gráfica principal con tkinter
Ventana de configuración limpia y profesional
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import logging
import os
from pathlib import Path

import config as cfg
import scheduler
import archiver

logger = logging.getLogger(__name__)

# ── Paleta de colores ──────────────────────────────────────────────────────────
BG        = "#1A1D23"
BG2       = "#22262F"
BG3       = "#2C3140"
ACCENT    = "#4F8EF7"
ACCENT2   = "#3A6ED4"
SUCCESS   = "#3DBE7A"
WARNING   = "#F0A500"
DANGER    = "#E05252"
TEXT      = "#E8EAF0"
TEXT2     = "#9AA0B2"
BORDER    = "#353A47"
FONT_H    = ("Segoe UI", 11, "bold")
FONT_B    = ("Segoe UI", 10)
FONT_S    = ("Segoe UI", 9)


def styled_frame(parent, **kw):
    return tk.Frame(parent, bg=BG2, **kw)


def label(parent, text, font=FONT_B, color=TEXT, **kw):
    return tk.Label(parent, text=text, font=font, fg=color, bg=parent["bg"], **kw)


def entry(parent, textvariable, width=30):
    return tk.Entry(
        parent, textvariable=textvariable, width=width,
        font=FONT_B, fg=TEXT, bg=BG3, insertbackground=TEXT,
        relief="flat", bd=0, highlightthickness=1,
        highlightcolor=ACCENT, highlightbackground=BORDER,
    )


def btn(parent, text, command, color=ACCENT, hover=ACCENT2, width=18):
    b = tk.Button(
        parent, text=text, command=command, font=FONT_B,
        fg=TEXT, bg=color, activeforeground=TEXT, activebackground=hover,
        relief="flat", bd=0, cursor="hand2", padx=12, pady=6, width=width,
    )
    b.bind("<Enter>", lambda e: b.config(bg=hover))
    b.bind("<Leave>", lambda e: b.config(bg=color))
    return b


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.config_data = cfg.load()
        self._setup_window()
        self._build_ui()
        self._load_values()
        self._refresh_status()

    # ── Ventana ─────────────────────────────────────────────────────────────
    def _setup_window(self):
        self.title("Outlook Archiver")
        self.configure(bg=BG)
        self.resizable(False, False)
        # Centrar en pantalla
        w, h = 560, 660
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        # Ícono (ignorar si no existe)
        try:
            self.iconbitmap(Path(__file__).parent / "icon.ico")
        except Exception:
            pass

    # ── UI ──────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header ──
        header = tk.Frame(self, bg=BG, pady=0)
        header.pack(fill="x")

        tk.Frame(header, bg=ACCENT, height=3).pack(fill="x")

        inner_h = tk.Frame(header, bg=BG, padx=28, pady=18)
        inner_h.pack(fill="x")
        label(inner_h, "📦 Outlook Archiver", font=("Segoe UI", 15, "bold"), color=TEXT).pack(anchor="w")
        label(inner_h, "Archivado automático de correos Outlook", color=TEXT2, font=FONT_S).pack(anchor="w")

        # ── Estado actual ──
        self._build_status_card()

        # ── Formulario de configuración ──
        self._build_config_form()

        # ── Botones de acción ──
        self._build_actions()

        # ── Footer ──
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", pady=(12, 0))
        footer = tk.Frame(self, bg=BG, padx=28, pady=10)
        footer.pack(fill="x")
        label(footer, "Los cambios se guardan automáticamente.", color=TEXT2, font=FONT_S).pack(anchor="w")

    def _build_status_card(self):
        outer = tk.Frame(self, bg=BG, padx=28, pady=8)
        outer.pack(fill="x")

        card = tk.Frame(outer, bg=BG2, padx=18, pady=14, highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="x")

        row = tk.Frame(card, bg=BG2)
        row.pack(fill="x")

        label(row, "Estado de la tarea programada", font=FONT_H, color=TEXT).pack(side="left")

        self.status_dot = tk.Label(row, text="●", font=("Segoe UI", 14), bg=BG2, fg=TEXT2)
        self.status_dot.pack(side="right")

        self.status_label = label(card, "Comprobando...", color=TEXT2, font=FONT_S)
        self.status_label.pack(anchor="w", pady=(4, 0))

        autostart_row = tk.Frame(card, bg=BG2)
        autostart_row.pack(fill="x", pady=(6, 0))
        label(autostart_row, "Inicio con Windows:", color=TEXT2, font=FONT_S).pack(side="left")
        self.autostart_status = label(autostart_row, "—", color=TEXT2, font=("Segoe UI", 9, "bold"))
        self.autostart_status.pack(side="left", padx=(6, 0))

        size_row = tk.Frame(card, bg=BG2)
        size_row.pack(fill="x", pady=(8, 0))

        label(size_row, "Tamaño actual del buzón:", color=TEXT2, font=FONT_S).pack(side="left")
        self.size_label = label(size_row, "—", color=ACCENT, font=("Segoe UI", 9, "bold"))
        self.size_label.pack(side="left", padx=(6, 0))

    def _build_config_form(self):
        outer = tk.Frame(self, bg=BG, padx=28, pady=4)
        outer.pack(fill="x")

        card = tk.Frame(outer, bg=BG2, padx=18, pady=16, highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="x")

        label(card, "Configuración", font=FONT_H, color=TEXT).pack(anchor="w", pady=(0, 12))

        # Umbral de tamaño
        self.threshold_var = tk.StringVar()
        self._form_row(card, "Archivar cuando el buzón supere (GB):", self.threshold_var)

        # Antigüedad de correos
        self.months_var = tk.StringVar()
        self._form_row(card, "Archivar correos más antiguos que (meses):", self.months_var)

        # Hora de ejecución
        time_frame = tk.Frame(card, bg=BG2)
        time_frame.pack(fill="x", pady=4)
        label(time_frame, "Ejecutar diariamente a las:", color=TEXT2, font=FONT_S).pack(anchor="w")

        time_inner = tk.Frame(time_frame, bg=BG2)
        time_inner.pack(anchor="w", pady=(4, 0))

        self.hour_var = tk.StringVar()
        self.minute_var = tk.StringVar()

        hour_entry = entry(time_inner, self.hour_var, width=5)
        hour_entry.pack(side="left")
        label(time_inner, ":", color=TEXT, font=FONT_H).pack(side="left", padx=4)
        min_entry = entry(time_inner, self.minute_var, width=5)
        min_entry.pack(side="left")
        label(time_inner, " (HH:MM)", color=TEXT2, font=FONT_S).pack(side="left", padx=(8, 0))

        # Ruta del .pst
        pst_frame = tk.Frame(card, bg=BG2)
        pst_frame.pack(fill="x", pady=(8, 4))
        label(pst_frame, "Ruta del archivo de archivo (.pst):", color=TEXT2, font=FONT_S).pack(anchor="w")

        pst_inner = tk.Frame(pst_frame, bg=BG2)
        pst_inner.pack(fill="x", pady=(4, 0))

        self.pst_var = tk.StringVar()
        pst_e = entry(pst_inner, self.pst_var, width=38)
        pst_e.pack(side="left", fill="x", expand=True)

        browse_btn = tk.Button(
            pst_inner, text="…", command=self._browse_pst,
            font=FONT_B, fg=TEXT, bg=BG3, activeforeground=TEXT,
            activebackground=BORDER, relief="flat", bd=0,
            cursor="hand2", padx=10, pady=5,
        )
        browse_btn.pack(side="left", padx=(6, 0))

        # Inicio automático con Windows
        self.autostart_var = tk.BooleanVar()
        chk_frame = tk.Frame(card, bg=BG2)
        chk_frame.pack(fill="x", pady=(10, 0))
        tk.Checkbutton(
            chk_frame,
            text="Iniciar automáticamente con Windows",
            variable=self.autostart_var,
            font=FONT_B, fg=TEXT, bg=BG2,
            activeforeground=TEXT, activebackground=BG2,
            selectcolor=BG3, relief="flat", bd=0, cursor="hand2",
        ).pack(anchor="w")

    def _form_row(self, parent, label_text, var):
        frame = tk.Frame(parent, bg=BG2)
        frame.pack(fill="x", pady=4)
        label(frame, label_text, color=TEXT2, font=FONT_S).pack(anchor="w")
        e = entry(frame, var, width=12)
        e.pack(anchor="w", pady=(4, 0))

    def _build_actions(self):
        outer = tk.Frame(self, bg=BG, padx=28, pady=10)
        outer.pack(fill="x")

        row1 = tk.Frame(outer, bg=BG)
        row1.pack(fill="x", pady=(0, 8))

        btn(row1, "💾  Guardar y programar", self._save_and_schedule, color=ACCENT, hover=ACCENT2).pack(side="left")
        btn(row1, "▶  Archivar ahora", self._run_now, color=SUCCESS, hover="#2EA865", width=16).pack(side="left", padx=(10, 0))

        row2 = tk.Frame(outer, bg=BG)
        row2.pack(fill="x")

        btn(row2, "🗑  Desactivar tarea", self._remove_task, color=BG3, hover=BORDER, width=18).pack(side="left")
        btn(row2, "📋  Ver log", self._open_log, color=BG3, hover=BORDER, width=14).pack(side="left", padx=(10, 0))
        btn(row2, "🔧  Reconfigurar", self._rerun_wizard, color=BG3, hover=BORDER, width=14).pack(side="left", padx=(10, 0))

        # Barra de progreso y mensaje
        self.progress_bar = ttk.Progressbar(outer, mode="indeterminate", length=500)
        self.msg_label = label(outer, "", color=TEXT2, font=FONT_S)
        self.msg_label.pack(anchor="w", pady=(10, 0))

    # ── Lógica de botones ───────────────────────────────────────────────────
    def _load_values(self):
        d = self.config_data
        self.threshold_var.set(d.get("threshold_gb", 4.0))
        self.months_var.set(d.get("months_old", 12))
        self.hour_var.set(f"{int(d.get('schedule_hour', 20)):02d}")
        self.minute_var.set(f"{int(d.get('schedule_minute', 0)):02d}")
        self.pst_var.set(d.get("pst_path", ""))
        self.autostart_var.set(d.get("autostart", True))

    def _collect_values(self) -> dict:
        d = self.config_data.copy()
        try:
            d["threshold_gb"] = float(self.threshold_var.get())
            d["months_old"] = int(self.months_var.get())
            d["schedule_hour"] = int(self.hour_var.get())
            d["schedule_minute"] = int(self.minute_var.get())
            d["pst_path"] = self.pst_var.get().strip()
            d["autostart"] = self.autostart_var.get()
        except ValueError as e:
            raise ValueError(f"Valor inválido en el formulario: {e}")
        return d

    def _save_and_schedule(self):
        try:
            data = self._collect_values()
        except ValueError as e:
            messagebox.showerror("Error de configuración", str(e))
            return

        cfg.save(data)
        self.config_data = data
        ok = scheduler.register_task(data["schedule_hour"], data["schedule_minute"])

        # Inicio automático con Windows
        import startup
        if data.get("autostart", True):
            startup.enable_autostart()
        else:
            startup.disable_autostart()

        if ok:
            self._set_msg(f"✔ Configuración guardada. Tarea programada a las {data['schedule_hour']:02d}:{data['schedule_minute']:02d}.", SUCCESS)
        else:
            self._set_msg("⚠ Configuración guardada, pero hubo un error al registrar la tarea.", WARNING)

        self._refresh_status()

    def _run_now(self):
        try:
            data = self._collect_values()
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return

        cfg.save(data)
        self.config_data = data
        self._set_msg("⏳ Archivando correos... por favor espera.", TEXT2)
        self.progress_bar.pack(anchor="w", pady=(6, 0))
        self.progress_bar.start(10)

        def worker():
            result = archiver.run_archive(data)
            self.after(0, lambda: self._on_archive_done(result))

        threading.Thread(target=worker, daemon=True).start()

    def _on_archive_done(self, result):
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        color = SUCCESS if result["status"] == "ok" else DANGER
        self._set_msg(result["message"], color)
        self._refresh_status()

    def _remove_task(self):
        if messagebox.askyesno("Desactivar tarea", "¿Deseas eliminar la tarea programada?"):
            ok = scheduler.remove_task()
            msg = "Tarea eliminada correctamente." if ok else "No se encontró la tarea o no se pudo eliminar."
            self._set_msg(("✔ " if ok else "⚠ ") + msg, SUCCESS if ok else WARNING)
            self._refresh_status()

    def _open_log(self):
        log_path = cfg.get_log_path(self.config_data)
        if log_path.exists():
            os.startfile(log_path)
        else:
            messagebox.showinfo("Log", f"No se encontró el archivo de log en:\n{log_path}")

    def _browse_pst(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".pst",
            filetypes=[("Archivo de datos de Outlook", "*.pst")],
            title="Seleccionar ubicación del archivo .pst",
        )
        if path:
            self.pst_var.set(path)

    def _refresh_status(self):
        exists = scheduler.task_exists()
        if exists:
            self.status_dot.config(fg=SUCCESS)
            self.status_label.config(text="Tarea activa en el Programador de Windows", fg=SUCCESS)
        else:
            self.status_dot.config(fg=DANGER)
            self.status_label.config(text="Sin tarea programada — haz clic en 'Guardar y programar'", fg=DANGER)

        # Estado inicio automático
        import startup
        if startup.autostart_enabled():
            self.autostart_status.config(text="Activo", fg=SUCCESS)
        else:
            self.autostart_status.config(text="Inactivo", fg=TEXT2)

        # Tamaño actual del buzón
        try:
            size = archiver.get_ost_size_gb()
            threshold = float(self.config_data.get("threshold_gb", 4))
            color = DANGER if size >= threshold else SUCCESS
            self.size_label.config(text=f"{size:.2f} GB", fg=color)
        except Exception:
            self.size_label.config(text="No disponible", fg=TEXT2)

    def _set_msg(self, text, color=TEXT2):
        self.msg_label.config(text=text, fg=color)


    def _rerun_wizard(self):
        import wizard, scheduler, startup
        result = wizard.run_wizard()
        if result is None:
            return
        import config as cfg
        cfg.save(result)
        self.config_data = result
        self._load_values()
        scheduler.register_task(result["schedule_hour"], result["schedule_minute"])
        if result.get("autostart", True):
            startup.enable_autostart()
        else:
            startup.disable_autostart()
        self._refresh_status()
        self._set_msg("✔ Reconfiguración completada.", SUCCESS)


def run():
    app = App()
    app.mainloop()
