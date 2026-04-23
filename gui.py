"""
gui.py - Ventana principal con scroll + bandeja del sistema
La ventana se adapta a la pantalla disponible y tiene scrollbar vertical.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import logging
import os
from pathlib import Path
from datetime import date

import config as cfg
import scheduler
import startup
import archiver

logger = logging.getLogger(__name__)

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


class ScrollableFrame(tk.Frame):
    """Frame con scrollbar vertical interna."""
    def __init__(self, parent, bg=BG, **kw):
        super().__init__(parent, bg=bg, **kw)

        self._canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        self._sb     = ttk.Scrollbar(self, orient="vertical",
                                     command=self._canvas.yview)
        self.inner   = tk.Frame(self._canvas, bg=bg)

        self._win_id = self._canvas.create_window(
            (0, 0), window=self.inner, anchor="nw"
        )
        self._canvas.configure(yscrollcommand=self._sb.set)

        self._sb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self.inner.bind("<Configure>", self._on_inner_configure)
        self._canvas.bind("<Configure>", self._on_canvas_configure)

        # Scroll con rueda del raton
        self._canvas.bind_all("<MouseWheel>",
            lambda e: self._canvas.yview_scroll(-1*(e.delta//120), "units"))

    def _on_inner_configure(self, _event):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self._canvas.itemconfig(self._win_id, width=event.width)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self._config    = cfg.load()
        self._tray_icon = None
        self._hidden    = False

        self._setup_window()
        self._init_vars()
        self._build_ui()
        self._load_values()
        self._refresh_status()
        self._start_systray()

        self.protocol("WM_DELETE_WINDOW", self._hide_to_tray)

    # ── Ventana ───────────────────────────────────────────────────────────────
    def _setup_window(self):
        self.title("Outlook Archiver")
        self.configure(bg=BG)
        self.resizable(True, True)

        # Adaptar al tamanio de pantalla disponible
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(600, sw - 40)
        h  = min(850, sh - 80)  
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{max(20,(sh-h)//2)}")
        self.minsize(520, 400)

        try:
            self.iconbitmap(Path(__file__).parent / "icon.ico")
        except Exception:
            pass

    def _init_vars(self):
        self.threshold_var = tk.StringVar(self)
        self.pst_dir_var   = tk.StringVar(self)
        self.pst_max_var   = tk.StringVar(self)
        self.hour_var      = tk.StringVar(self)
        self.minute_var    = tk.StringVar(self)
        self.autostart_var = tk.BooleanVar(self)

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Header fijo (fuera del scroll)
        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x", side="top")
        hdr = tk.Frame(self, bg=BG, padx=24, pady=14)
        hdr.pack(fill="x", side="top")
        _lbl(hdr, "Outlook Archiver",
             font=("Segoe UI", 14, "bold")).pack(anchor="w")
        _lbl(hdr, "Archivado automatico por año con PST rotativo",
             color=TEXT2, font=FONT_S).pack(anchor="w")
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", side="top")

        # Zona de botones de accion FIJA en la parte inferior
        self._build_action_bar()
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", side="bottom")

        # Area scrolleable con el resto del contenido
        self._scroll = ScrollableFrame(self)
        self._scroll.pack(fill="both", expand=True, side="top")
        f = self._scroll.inner

        self._build_status_card(f)
        self._build_next_archive_card(f)
        self._build_config_form(f)

    def _build_action_bar(self):
        """Barra de botones FIJA en la parte inferior, siempre visible."""
        bar = tk.Frame(self, bg=BG2, padx=24, pady=12)
        bar.pack(fill="x", side="bottom")

        r1 = tk.Frame(bar, bg=BG2)
        r1.pack(fill="x", pady=(0, 6))

        _btn(r1, "Guardar y programar", self._save,
             color=ACCENT, hover=ACCENT2, width=20).pack(side="left")
        _btn(r1, "Archivar ahora", self._run_now,
             color=SUCCESS, hover="#2EA865", width=14).pack(side="left", padx=(8, 0))

        r2 = tk.Frame(bar, bg=BG2)
        r2.pack(fill="x")

        _btn(r2, "Desactivar tarea", self._remove_task,
             color=BG3, hover=BORDER, width=16).pack(side="left")
        _btn(r2, "Ver log", self._open_log,
             color=BG3, hover=BORDER, width=10).pack(side="left", padx=(8, 0))
        _btn(r2, "Reconfigurar", self._rerun_wizard,
             color=BG3, hover=BORDER, width=12).pack(side="left", padx=(8, 0))
        _btn(r2, "Desinstalar", self._uninstall,
             color=DANGER, hover="#B83C3C", width=12).pack(side="right")

        self.progress = ttk.Progressbar(bar, mode="indeterminate", length=400)
        self.msg_lbl  = _lbl(bar, "", color=TEXT2, font=FONT_S)
        self.msg_lbl.pack(anchor="w", pady=(6, 0))

    def _build_status_card(self, parent):
        outer = tk.Frame(parent, bg=BG, padx=24, pady=6)
        outer.pack(fill="x")
        card = tk.Frame(outer, bg=BG2, padx=16, pady=12,
                        highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="x")

        row = tk.Frame(card, bg=BG2)
        row.pack(fill="x")
        _lbl(row, "Estado", font=FONT_H).pack(side="left")
        self.status_dot = tk.Label(row, text="*", font=("Segoe UI", 14),
                                   bg=BG2, fg=TEXT2)
        self.status_dot.pack(side="right")

        self.status_lbl = _lbl(card, "Comprobando...", color=TEXT2, font=FONT_S)
        self.status_lbl.pack(anchor="w", pady=(3, 0))

        info = tk.Frame(card, bg=BG2)
        info.pack(fill="x", pady=(8, 0))
        _lbl(info, "Inicio con Windows:", color=TEXT2, font=FONT_S).pack(side="left")
        self.autostart_lbl = _lbl(info, "-", color=TEXT2,
                                   font=("Segoe UI", 9, "bold"))
        self.autostart_lbl.pack(side="left", padx=(4, 20))
        _lbl(info, "Tamanio buzon:", color=TEXT2, font=FONT_S).pack(side="left")
        self.size_lbl = _lbl(info, "-", color=ACCENT,
                              font=("Segoe UI", 9, "bold"))
        self.size_lbl.pack(side="left", padx=(4, 0))

    def _build_next_archive_card(self, parent):
        outer = tk.Frame(parent, bg=BG, padx=24, pady=4)
        outer.pack(fill="x")
        card = tk.Frame(outer, bg=BG2, padx=16, pady=10,
                        highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="x")

        _lbl(card, "Proximo archivado", font=FONT_H).pack(anchor="w")

        from archiver import compute_cutoff_date, compute_archive_year
        today    = date.today()
        cutoff   = compute_cutoff_date(today)
        arc_year = compute_archive_year(cutoff)
        last_day = cutoff.strftime("%d/%m/%Y")

        r1 = tk.Frame(card, bg=BG2)
        r1.pack(fill="x", pady=(6, 2))
        _lbl(r1, "Archivara hasta:", color=TEXT2, font=FONT_S).pack(side="left")
        _lbl(r1, f"{last_day} (exclusive) = correos hasta el dia anterior",
             color=TEXT, font=("Segoe UI", 9, "bold")).pack(side="left", padx=(6, 0))

        r2 = tk.Frame(card, bg=BG2)
        r2.pack(fill="x", pady=2)
        _lbl(r2, "PST destino:", color=TEXT2, font=FONT_S).pack(side="left")
        self.pst_dest_lbl = _lbl(r2, f"Archivo {arc_year}.pst",
                                  color=ACCENT, font=("Segoe UI", 9, "bold"))
        self.pst_dest_lbl.pack(side="left", padx=(6, 0))

        r3 = tk.Frame(card, bg=BG2)
        r3.pack(fill="x", pady=2)
        _lbl(r3, "Tamanio actual PST:", color=TEXT2, font=FONT_S).pack(side="left")
        self.pst_size_lbl = _lbl(r3, "-", color=TEXT2,
                                  font=("Segoe UI", 9, "bold"))
        self.pst_size_lbl.pack(side="left", padx=(6, 0))

    def _build_config_form(self, parent):
        outer = tk.Frame(parent, bg=BG, padx=24, pady=6)
        outer.pack(fill="x")
        card = tk.Frame(outer, bg=BG2, padx=16, pady=14,
                        highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="x")
        _lbl(card, "Configuracion", font=FONT_H).pack(anchor="w", pady=(0, 8))

        self._form_row(card, "Archivar cuando el buzon supere (GB):",
                       self.threshold_var)

        # Carpeta PST
        pf = tk.Frame(card, bg=BG2)
        pf.pack(fill="x", pady=4)
        _lbl(pf, "Carpeta de archivos PST:", color=TEXT2, font=FONT_S).pack(anchor="w")
        pr = tk.Frame(pf, bg=BG2)
        pr.pack(fill="x", pady=(4, 0))
        _entry(pr, self.pst_dir_var, width=30).pack(
            side="left", fill="x", expand=True)
        tk.Button(
            pr, text="...", command=self._browse_dir,
            font=FONT_B, fg=TEXT, bg=BG3, activeforeground=TEXT,
            activebackground=BORDER, relief="raised", bd=1,
            cursor="hand2", padx=8, pady=4,
        ).pack(side="left", padx=(6, 0))

        self._form_row(card, "Limite de tamanio por PST (GB):",
                       self.pst_max_var)

        # Hora
        tf = tk.Frame(card, bg=BG2)
        tf.pack(fill="x", pady=4)
        _lbl(tf, "Hora de ejecucion diaria (HH : MM):",
             color=TEXT2, font=FONT_S).pack(anchor="w")
        tr = tk.Frame(tf, bg=BG2)
        tr.pack(anchor="w", pady=(4, 0))
        tk.Entry(
            tr, textvariable=self.hour_var, width=5,
            font=FONT_B, fg=TEXT, bg=BG3, insertbackground=TEXT,
            relief="flat", bd=0, highlightthickness=1,
            highlightcolor=ACCENT, highlightbackground=BORDER,
        ).pack(side="left")
        _lbl(tr, "  :  ", color=TEXT, font=FONT_H).pack(side="left")
        tk.Entry(
            tr, textvariable=self.minute_var, width=5,
            font=FONT_B, fg=TEXT, bg=BG3, insertbackground=TEXT,
            relief="flat", bd=0, highlightthickness=1,
            highlightcolor=ACCENT, highlightbackground=BORDER,
        ).pack(side="left")

        # Autostart
        ck = tk.Frame(card, bg=BG2)
        ck.pack(fill="x", pady=(10, 0))
        tk.Checkbutton(
            ck, text="Iniciar automaticamente con Windows",
            variable=self.autostart_var,
            font=FONT_B, fg=TEXT, bg=BG2,
            activeforeground=TEXT, activebackground=BG2,
            selectcolor=BG3, relief="flat", bd=0, cursor="hand2",
        ).pack(anchor="w")

        # Espaciado al final del scroll
        tk.Frame(parent, bg=BG, height=12).pack(fill="x")

    def _form_row(self, parent, label_text, var):
        f = tk.Frame(parent, bg=BG2)
        f.pack(fill="x", pady=4)
        _lbl(f, label_text, color=TEXT2, font=FONT_S).pack(anchor="w")
        _entry(f, var, width=12).pack(anchor="w", pady=(4, 0))

    # ── Logica de botones ─────────────────────────────────────────────────────
    def _load_values(self):
        d = self._config
        self.threshold_var.set(d.get("threshold_gb", 3.0))
        self.pst_dir_var.set(d.get("pst_base_dir", ""))
        self.pst_max_var.set(d.get("pst_max_gb", 47.0))
        self.hour_var.set(f"{int(d.get('schedule_hour', 20)):02d}")
        self.minute_var.set(f"{int(d.get('schedule_minute', 0)):02d}")
        self.autostart_var.set(d.get("autostart", True))

    def _collect(self):
        d = self._config.copy()
        try:
            d["threshold_gb"]    = float(self.threshold_var.get())
            d["pst_base_dir"]    = self.pst_dir_var.get().strip()
            d["pst_max_gb"]      = float(self.pst_max_var.get())
            d["schedule_hour"]   = int(self.hour_var.get())
            d["schedule_minute"] = int(self.minute_var.get())
            d["autostart"]       = self.autostart_var.get()
        except ValueError as e:
            raise ValueError(f"Valor invalido: {e}")
        return d

    def _save(self):
        try:
            data = self._collect()
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return
        cfg.save(data)
        self._config = data
        ok = scheduler.register_task(data["schedule_hour"], data["schedule_minute"])
        if data.get("autostart"):
            startup.enable_autostart()
        else:
            startup.disable_autostart()
        color = SUCCESS if ok else WARNING
        msg   = "Configuracion guardada y tarea programada." if ok \
                else "Configuracion guardada pero hubo un error al programar la tarea."
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

        def worker():
            result = archiver.run_archive(data)
            self.after(0, lambda: self._on_done(result))

        threading.Thread(target=worker, daemon=True).start()

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
            self._set_msg("Tarea eliminada." if ok else "No se pudo eliminar la tarea.",
                          SUCCESS if ok else WARNING)
            self._refresh_status()

    def _open_log(self):
        lp = cfg.get_log_path(self._config)
        if lp.exists():
            os.startfile(lp)
        else:
            messagebox.showinfo("Log", f"No existe aun el log en:\n{lp}")

    def _browse_dir(self):
        path = filedialog.askdirectory(title="Carpeta para archivos PST")
        if path:
            self.pst_dir_var.set(path)

    def _rerun_wizard(self):
        import wizard
        self.withdraw()
        result = wizard.run_wizard()
        if result:
            cfg.save(result)
            self._config = result
            self._load_values()
            scheduler.register_task(result["schedule_hour"], result["schedule_minute"])
            startup.enable_autostart() if result.get("autostart") \
                else startup.disable_autostart()
            self._refresh_status()
            self._set_msg("Reconfigurado correctamente.", SUCCESS)
        self.deiconify()

    def _uninstall(self):
        if not messagebox.askyesno(
            "Desinstalar Outlook Archiver",
            "Esto eliminara:\n"
            "  - La tarea del Programador de Windows\n"
            "  - El inicio automatico con Windows\n"
            "  - La configuracion guardada\n\n"
            "Los archivos PST y el .exe NO se eliminan.\n\n"
            "Continuar?",
        ):
            return

        scheduler.remove_task()
        startup.disable_autostart()

        # Eliminar config.json
        try:
            cfg.get_config_path().unlink(missing_ok=True)
        except Exception as e:
            logger.warning("No se pudo eliminar config: %s", e)

        messagebox.showinfo(
            "Desinstalado",
            "Outlook Archiver ha sido desinstalado.\n\n"
            "Puedes eliminar el archivo .exe manualmente.\n"
            "Los archivos PST permanecen intactos.",
        )

        if self._tray_icon:
            self._tray_icon.stop()
        self.destroy()

    def _refresh_status(self):
        exists = scheduler.task_exists()
        self.status_dot.config(fg=SUCCESS if exists else DANGER)
        self.status_lbl.config(
            text="Tarea activa en el Programador de Windows" if exists
                 else "Sin tarea — haz clic en 'Guardar y programar'",
            fg=SUCCESS if exists else DANGER,
        )

        self.autostart_lbl.config(
            text="Activo" if startup.autostart_enabled() else "Inactivo",
            fg=SUCCESS if startup.autostart_enabled() else TEXT2,
        )

        try:
            sz  = archiver.get_ost_size_gb()
            thr = float(self._config.get("threshold_gb", 3))
            self.size_lbl.config(
                text=f"{sz:.2f} GB",
                fg=DANGER if sz >= thr else SUCCESS,
            )
        except Exception:
            self.size_lbl.config(text="N/D", fg=TEXT2)

        try:
            from archiver import compute_cutoff_date, compute_archive_year, \
                                 get_pst_path_for_year
            today    = date.today()
            cutoff   = compute_cutoff_date(today)
            arc_year = compute_archive_year(cutoff)
            base_dir = self._config.get("pst_base_dir", "")
            pst_path = get_pst_path_for_year(base_dir, arc_year)
            pst_sz   = archiver.get_pst_size_gb(pst_path)
            pst_max  = float(self._config.get("pst_max_gb", 47))
            self.pst_dest_lbl.config(text=f"Archivo {arc_year}.pst")
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

    def _show_from_tray(self, icon=None, item=None):
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
                "Cerrar el programa? La tarea programada seguira activa."):
                self._quit_app()

    def _tray_archive_now(self, icon=None, item=None):
        threading.Thread(
            target=lambda: archiver.run_archive(self._config),
            daemon=True,
        ).start()

    def _quit_app(self, icon=None, item=None):
        if self._tray_icon:
            self._tray_icon.stop()
        self.after(0, self.destroy)


def run():
    app = App()
    app.mainloop()