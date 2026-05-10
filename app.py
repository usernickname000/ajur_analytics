# ============================================================
# app.py — Аналитика заказов, Фонтанка.ру  v3.0
# ============================================================

import sys
import os as _os
_HERE = _os.path.dirname(_os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import json
import os
from datetime import datetime, timezone, timedelta

from analytics import run_analytics
from comparison import run_comparison
from watcher import FolderWatcher
from dashboard import generate_dashboard

try:
    from external_income_editor import ExternalIncomeEditor
    HAS_EXT_EDITOR = True
except ImportError:
    HAS_EXT_EDITOR = False

try:
    from verified_figures_editor import VerifiedFiguresEditor
    HAS_VF_EDITOR = True
except ImportError:
    HAS_VF_EDITOR = False

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
ICON_PATH   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
APP_DIR     = os.path.dirname(os.path.abspath(__file__))

# ── Палитра ──────────────────────────────────────────────────
C_ORANGE   = "#F38120"
C_DARK     = "#D06A10"
C_GREEN    = "#22C55E"
C_GREEN_D  = "#16A34A"
C_RED      = "#EF4444"
C_RED_D    = "#DC2626"
C_AMBER    = "#F59E0B"
C_BLUE     = "#3B82F6"

THEMES = {
    "dark": {
        "bg":          "#0D0D0D",
        "sidebar":     "#111111",
        "surface":     "#181818",
        "surface2":    "#222222",
        "surface3":    "#2A2A2A",
        "border":      "#2E2E2E",
        "border2":     "#3A3A3A",
        "text":        "#F5F5F5",
        "text2":       "#AAAAAA",
        "muted":       "#555555",
        "log_bg":      "#0A0A0A",
        "log_fg":      "#CCCCCC",
        "entry_bg":    "#1E1E1E",
        "entry_fg":    "#F5F5F5",
        "tag_ok":      "#22C55E",
        "tag_err":     "#EF4444",
        "tag_warn":    "#F59E0B",
        "nav_active":  "#1E1E1E",
        "nav_hover":   "#161616",
        "accent_line": "#F38120",
    },
    "light": {
        "bg":          "#F4F4F5",
        "sidebar":     "#FAFAFA",
        "surface":     "#FFFFFF",
        "surface2":    "#F0F0F0",
        "surface3":    "#E8E8E8",
        "border":      "#E2E2E2",
        "border2":     "#CDCDCD",
        "text":        "#0A0A0A",
        "text2":       "#555555",
        "muted":       "#999999",
        "log_bg":      "#111111",
        "log_fg":      "#E8E8E8",
        "entry_bg":    "#FFFFFF",
        "entry_fg":    "#0A0A0A",
        "tag_ok":      "#16A34A",
        "tag_err":     "#DC2626",
        "tag_warn":    "#D97706",
        "nav_active":  "#FFF5EB",
        "nav_hover":   "#FFF8F2",
        "accent_line": "#F38120",
    },
}


def load_config():
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ============================================================
# TOAST
# ============================================================
class Toast(tk.Toplevel):
    _BG     = {"ok": "#0F2A1A", "err": "#2A0F0F", "warn": "#2A200A", "info": "#0A1830"}
    _BORDER = {"ok": "#22C55E", "err": "#EF4444", "warn": "#F59E0B", "info": "#3B82F6"}
    _ICON   = {"ok": "✓", "err": "✕", "warn": "!", "info": "i"}

    def __init__(self, parent, message, kind="ok",
                 action_text=None, action_cmd=None, duration=4000):
        super().__init__(parent)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.attributes("-alpha", 0.0)

        bg     = self._BG.get(kind, "#181818")
        border = self._BORDER.get(kind, "#444")
        icon   = self._ICON.get(kind, "·")

        self.configure(bg=border)
        inner = tk.Frame(self, bg=bg, padx=0, pady=0)
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        # Левая акцент-полоса
        accent = tk.Frame(inner, bg=border, width=3)
        accent.pack(side="left", fill="y")

        content = tk.Frame(inner, bg=bg, padx=14, pady=10)
        content.pack(side="left", fill="both", expand=True)

        top = tk.Frame(content, bg=bg)
        top.pack(fill="x")

        # Иконка в кружке
        ico_f = tk.Frame(top, bg=border, width=22, height=22)
        ico_f.pack(side="left", padx=(0, 10))
        ico_f.pack_propagate(False)
        tk.Label(ico_f, text=icon, font=("Segoe UI", 9, "bold"),
                 bg=border, fg="#FFFFFF").place(relx=.5, rely=.5, anchor="center")

        tk.Label(top, text=message,
                 font=("Segoe UI", 9), bg=bg, fg="#F0F0F0",
                 justify="left", wraplength=260).pack(side="left", fill="x", expand=True)

        tk.Button(top, text="×", font=("Segoe UI", 11),
                  bg=bg, fg="#666", relief="flat", bd=0,
                  activebackground=bg, activeforeground="#FFF",
                  cursor="hand2", command=self._dismiss).pack(side="right", padx=(8, 0))

        if action_text and action_cmd:
            btn_row = tk.Frame(content, bg=bg)
            btn_row.pack(fill="x", pady=(8, 0))
            tk.Button(btn_row, text=action_text,
                      font=("Segoe UI", 8, "bold"),
                      bg=border, fg="#FFF",
                      activebackground="#FFF", activeforeground=border,
                      relief="flat", bd=0, cursor="hand2",
                      padx=12, pady=4,
                      command=lambda: (action_cmd(), self._dismiss())
                      ).pack(side="left")

        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = sw - w - 20
        self._y_show = sh - h - 48
        self._y_hide = sh + 20
        self.geometry(f"{w}x{h}+{x}+{self._y_hide}")
        self._duration = duration
        self._alpha = 0.0
        self._fade_in()

    def _fade_in(self):
        self._alpha = min(self._alpha + 0.1, 1.0)
        cur_y = int(self._y_hide + (self._y_show - self._y_hide) * self._alpha)
        try:
            self.attributes("-alpha", self._alpha)
            sw = self.winfo_screenwidth()
            self.geometry(f"+{sw - self.winfo_reqwidth() - 20}+{cur_y}")
        except Exception:
            return
        if self._alpha < 1.0:
            self.after(14, self._fade_in)
        else:
            self.after(self._duration, self._fade_out)

    def _fade_out(self):
        self._alpha = max(self._alpha - 0.07, 0.0)
        cur_y = int(self._y_show + (self._y_hide - self._y_show) * (1 - self._alpha))
        try:
            self.attributes("-alpha", self._alpha)
            sw = self.winfo_screenwidth()
            self.geometry(f"+{sw - self.winfo_reqwidth() - 20}+{cur_y}")
        except Exception:
            return
        if self._alpha > 0:
            self.after(14, self._fade_out)
        else:
            try:
                self.destroy()
            except Exception:
                pass

    def _dismiss(self):
        self._duration = 0
        self._fade_out()


# ============================================================
# КНОПКА С АНИМАЦИЕЙ HOVER
# ============================================================
class AnimButton(tk.Button):
    STEPS = 6
    DELAY = 10

    def __init__(self, master, bg_n, bg_h, **kw):
        self._bg_n = bg_n
        self._bg_h = bg_h
        self._cur  = list(self._hex(bg_n))
        self._tgt  = list(self._hex(bg_n))
        self._job  = None
        super().__init__(master, bg=bg_n, relief="flat",
                         cursor="hand2", bd=0,
                         activebackground=bg_h,
                         activeforeground=kw.get("fg", "#FFFFFF"), **kw)
        self.bind("<Enter>", lambda e: self._go(self._hex(bg_h)))
        self.bind("<Leave>", lambda e: self._go(self._hex(bg_n)))

    def _hex(self, c):
        c = c.lstrip("#")
        return [int(c[i:i+2], 16) for i in (0, 2, 4)]

    def _go(self, tgt):
        self._tgt = tgt
        if self._job:
            try:
                self.after_cancel(self._job)
            except Exception:
                pass
        self._tick()

    def _tick(self):
        done = True
        for i in range(3):
            d = self._tgt[i] - self._cur[i]
            if d:
                done = False
                self._cur[i] += max(1, abs(d) // self.STEPS) * (1 if d > 0 else -1)
                self._cur[i] = max(0, min(255, self._cur[i]))
        try:
            self.config(bg="#{:02x}{:02x}{:02x}".format(*[int(x) for x in self._cur]))
        except Exception:
            return
        if not done:
            self._job = self.after(self.DELAY, self._tick)


# ============================================================
# ГЛАВНОЕ ОКНО
# ============================================================
class App(tk.Tk):
    # Навигационные пункты: (id, иконка, подпись)
    NAV_ITEMS = [
        ("analysis",    "▸",  "Анализ"),
        ("compare",     "⇄",  "Сравнение"),
        ("plan",        "≡",  "План"),
        ("watcher",     "◎",  "Наблюдатель"),
        ("dashboard",   "⊞",  "Дашборд"),
    ]

    def __init__(self):
        super().__init__()
        self.title("Аналитика — Фонтанка.ру")
        self.geometry("1040x720")
        self.minsize(900, 640)
        self.resizable(True, True)

        if os.path.exists(ICON_PATH):
            try:
                self.iconbitmap(default=ICON_PATH)
            except Exception:
                pass

        self._cfg         = load_config()
        self._tn          = self._cfg.get("theme", "dark")
        self._T           = THEMES[self._tn]
        self._tw          = []
        self._watcher     = None
        self._plan_vars   = {}
        self._anim_running= False
        self._anim_step   = 0
        self._date_by     = tk.StringVar(value=self._cfg.get("date_by", "order"))
        self._active_page = tk.StringVar(value="analysis")
        self._pages       = {}
        self._nav_btns    = {}

        self.configure(bg=self._T["bg"])
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Toast ────────────────────────────────────────────────
    def _toast(self, message, kind="ok", action_text=None, action_cmd=None, duration=4000):
        try:
            Toast(self, message, kind=kind,
                  action_text=action_text, action_cmd=action_cmd, duration=duration)
        except Exception:
            pass

    def _on_close(self):
        if self._watcher:
            self._watcher.stop()
        self._cfg["theme"]  = self._tn
        self._cfg["date_by"]= self._date_by.get()
        save_config(self._cfg)
        self.destroy()

    # ── Регистрация виджетов ─────────────────────────────────
    def _r(self, w, role):
        self._tw.append((w, role))
        return w

    # ============================================================
    # СБОРКА UI
    # ============================================================
    def _build_ui(self):
        # Корневой контейнер
        root = tk.Frame(self, bg=self._T["bg"])
        root.pack(fill="both", expand=True)

        # Боковая панель
        self._sidebar = tk.Frame(root, bg=self._T["sidebar"], width=220)
        self._sidebar.pack(side="left", fill="y")
        self._sidebar.pack_propagate(False)

        # Разделитель
        tk.Frame(root, bg=self._T["border"], width=1).pack(side="left", fill="y")

        # Рабочая область
        self._workspace = tk.Frame(root, bg=self._T["bg"])
        self._workspace.pack(side="left", fill="both", expand=True)

        self._build_sidebar()
        self._build_pages()
        self._build_statusbar()
        self._show_page("analysis")

    # ============================================================
    # САЙДБАР
    # ============================================================
    def _build_sidebar(self):
        T = self._T
        sb = self._sidebar

        # ── Логотип ──────────────────────────────────────────
        logo_frame = tk.Frame(sb, bg=T["sidebar"], height=80)
        logo_frame.pack(fill="x")
        logo_frame.pack_propagate(False)

        # Оранжевая вертикальная черта слева
        tk.Frame(logo_frame, bg=C_ORANGE, width=3).pack(side="left", fill="y")

        logo_inner = tk.Frame(logo_frame, bg=T["sidebar"])
        logo_inner.pack(side="left", fill="both", expand=True, padx=18)

        tk.Label(logo_inner, text="фонтанка.ру",
                 font=("Georgia", 16, "bold italic"),
                 bg=T["sidebar"], fg=C_ORANGE).pack(anchor="w", pady=(20, 0))
        tk.Label(logo_inner, text="КОММЕРЧЕСКИЙ ОТДЕЛ",
                 font=("Segoe UI", 7, "bold"),
                 bg=T["sidebar"], fg=T["muted"]).pack(anchor="w")

        # Тонкий разделитель
        tk.Frame(sb, bg=T["border"], height=1).pack(fill="x", pady=(4, 0))

        # ── Навигация ─────────────────────────────────────────
        nav_frame = tk.Frame(sb, bg=T["sidebar"])
        nav_frame.pack(fill="x", pady=(12, 0))

        for page_id, icon, label in self.NAV_ITEMS:
            self._nav_btns[page_id] = self._make_nav_btn(
                nav_frame, page_id, icon, label)

        # ── Нижний блок сайдбара ─────────────────────────────
        tk.Frame(sb, bg=T["border"], height=1).pack(fill="x", side="bottom", pady=(0, 10))

        bottom = tk.Frame(sb, bg=T["sidebar"])
        bottom.pack(side="bottom", fill="x", padx=14, pady=(0, 14))

        # Часы
        self._clock_frame = tk.Frame(bottom, bg=T["sidebar"])
        self._clock_frame.pack(fill="x", pady=(0, 12))

        self._clock_time = tk.StringVar()
        self._clock_date = tk.StringVar()
        tk.Label(self._clock_frame, textvariable=self._clock_time,
                 font=("Segoe UI", 20, "bold"),
                 bg=T["sidebar"], fg=T["text"]).pack(anchor="w")
        tk.Label(self._clock_frame, textvariable=self._clock_date,
                 font=("Segoe UI", 8),
                 bg=T["sidebar"], fg=T["muted"]).pack(anchor="w")
        self._update_clock()

        # Кнопка темы
        self._theme_btn = tk.Button(
            bottom,
            text="☀  Светлая тема" if self._tn == "dark" else "◑  Тёмная тема",
            font=("Segoe UI", 9),
            bg=T["surface2"], fg=T["text2"],
            activebackground=T["surface3"],
            activeforeground=T["text"],
            relief="flat", bd=0, cursor="hand2",
            pady=7, anchor="w", padx=12,
            command=self._toggle_theme)
        self._theme_btn.pack(fill="x")

        # Версия
        tk.Label(bottom, text="v 3.0",
                 font=("Segoe UI", 8),
                 bg=T["sidebar"], fg=T["muted"]).pack(anchor="e", pady=(8, 0))

    def _make_nav_btn(self, parent, page_id, icon, label):
        T = self._T

        frame = tk.Frame(parent, bg=T["sidebar"], cursor="hand2")
        frame.pack(fill="x", padx=8, pady=1)

        # Акцент-полоска слева (скрыта по умолчанию)
        accent = tk.Frame(frame, bg=C_ORANGE, width=3)
        accent.pack(side="left", fill="y")
        accent.pack_forget()

        inner = tk.Frame(frame, bg=T["sidebar"], pady=10, padx=14)
        inner.pack(side="left", fill="both", expand=True)

        icon_lbl = tk.Label(inner, text=icon,
                            font=("Segoe UI", 13),
                            bg=T["sidebar"], fg=T["muted"],
                            width=2, anchor="w")
        icon_lbl.pack(side="left")

        text_lbl = tk.Label(inner, text=label,
                            font=("Segoe UI", 10),
                            bg=T["sidebar"], fg=T["text2"],
                            anchor="w")
        text_lbl.pack(side="left", padx=(8, 0))

        def on_click(_=None):
            self._show_page(page_id)

        def on_enter(_=None):
            if self._active_page.get() != page_id:
                frame.configure(bg=T["nav_hover"])
                inner.configure(bg=T["nav_hover"])
                icon_lbl.configure(bg=T["nav_hover"])
                text_lbl.configure(bg=T["nav_hover"])

        def on_leave(_=None):
            if self._active_page.get() != page_id:
                frame.configure(bg=T["sidebar"])
                inner.configure(bg=T["sidebar"])
                icon_lbl.configure(bg=T["sidebar"])
                text_lbl.configure(bg=T["sidebar"])

        for w in (frame, inner, icon_lbl, text_lbl):
            w.bind("<Button-1>", on_click)
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)

        return {"frame": frame, "inner": inner,
                "icon": icon_lbl, "text": text_lbl,
                "accent": accent}

    def _show_page(self, page_id):
        T = self._T
        prev = self._active_page.get()

        # Снять выделение с предыдущего
        if prev in self._nav_btns:
            btn = self._nav_btns[prev]
            btn["frame"].configure(bg=T["sidebar"])
            btn["inner"].configure(bg=T["sidebar"])
            btn["icon"].configure(bg=T["sidebar"], fg=T["muted"],
                                   font=("Segoe UI", 13))
            btn["text"].configure(bg=T["sidebar"], fg=T["text2"],
                                   font=("Segoe UI", 10))
            btn["accent"].pack_forget()

        # Выделить новый
        self._active_page.set(page_id)
        if page_id in self._nav_btns:
            btn = self._nav_btns[page_id]
            btn["accent"].pack(side="left", fill="y")
            btn["frame"].configure(bg=T["nav_active"])
            btn["inner"].configure(bg=T["nav_active"])
            btn["icon"].configure(bg=T["nav_active"], fg=C_ORANGE,
                                   font=("Segoe UI", 13, "bold"))
            btn["text"].configure(bg=T["nav_active"], fg=T["text"],
                                   font=("Segoe UI", 10, "bold"))

        # Переключить страницу
        for pid, page in self._pages.items():
            if pid == page_id:
                page.pack(fill="both", expand=True)
            else:
                page.pack_forget()

    def _update_clock(self):
        try:
            msk = timezone(timedelta(hours=3))
            now = datetime.now(tz=msk)
        except Exception:
            now = datetime.now()
        self._clock_time.set(now.strftime("%H:%M"))
        self._clock_date.set(now.strftime("%d %B %Y").upper())
        self.after(60000, self._update_clock)

    # ============================================================
    # СТАТУС-БАР
    # ============================================================
    def _build_statusbar(self):
        T = self._T
        sb = tk.Frame(self._workspace, bg=T["surface"], height=28)
        sb.pack(fill="x", side="bottom")
        sb.pack_propagate(False)

        tk.Frame(sb, bg=T["border"], height=1).pack(fill="x", side="top")

        self._status_var = tk.StringVar(value="Готов к работе")
        tk.Label(sb, textvariable=self._status_var,
                 font=("Segoe UI", 8), bg=T["surface"], fg=T["muted"],
                 anchor="w").pack(side="left", padx=16)

        self._anim_lbl = tk.Label(sb, text="",
                                   font=("Segoe UI", 10),
                                   bg=T["surface"], fg=C_ORANGE)
        self._anim_lbl.pack(side="right", padx=16)

    def _start_anim(self):
        self._anim_running = True
        self._anim_step = 0
        self._tick_anim()

    def _stop_anim(self):
        self._anim_running = False
        try:
            self._anim_lbl.config(text="")
        except Exception:
            pass

    def _tick_anim(self):
        if not self._anim_running:
            return
        frames = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        try:
            self._anim_lbl.config(text=frames[self._anim_step % len(frames)])
        except Exception:
            return
        self._anim_step += 1
        self.after(100, self._tick_anim)

    # ============================================================
    # ТЕМА
    # ============================================================
    def _toggle_theme(self):
        self._tn = "light" if self._tn == "dark" else "dark"
        self._T  = THEMES[self._tn]
        self._cfg["theme"] = self._tn
        save_config(self._cfg)
        self._apply_theme()

    def _apply_theme(self):
        T = self._T

        # Обновляем базовые цвета окна
        self.configure(bg=T["bg"])

        role_map = {
            "bg":       lambda w: w.configure(bg=T["bg"]),
            "sidebar":  lambda w: w.configure(bg=T["sidebar"]),
            "surface":  lambda w: w.configure(bg=T["surface"]),
            "surface2": lambda w: w.configure(bg=T["surface2"]),
            "border":   lambda w: w.configure(bg=T["border"]),
            "text_bg":  lambda w: w.configure(bg=T["bg"],      fg=T["text"]),
            "text2_bg": lambda w: w.configure(bg=T["bg"],      fg=T["text2"]),
            "muted_bg": lambda w: w.configure(bg=T["bg"],      fg=T["muted"]),
            "text_sf":  lambda w: w.configure(bg=T["surface"], fg=T["text"]),
            "muted_sf": lambda w: w.configure(bg=T["surface"], fg=T["muted"]),
            "text2_sf": lambda w: w.configure(bg=T["surface"], fg=T["text2"]),
            "entry":    lambda w: w.configure(
                bg=T["entry_bg"], fg=T["entry_fg"],
                readonlybackground=T["entry_bg"],
                insertbackground=T["entry_fg"],
                disabledbackground=T["entry_bg"]),
        }
        for (w, role) in self._tw:
            try:
                if role in role_map:
                    role_map[role](w)
            except Exception:
                pass

        # Статус-бар
        try:
            self._anim_lbl.configure(bg=T["surface"])
        except Exception:
            pass

        # Логи
        for attr in ["log_box", "cmp_log_box", "watch_log_box", "dash_log_box"]:
            box = getattr(self, attr, None)
            if box:
                try:
                    box.configure(bg=T["log_bg"], fg=T["log_fg"])
                    box.tag_config("ok",   foreground=T["tag_ok"])
                    box.tag_config("err",  foreground=T["tag_err"])
                    box.tag_config("warn", foreground=T["tag_warn"])
                except Exception:
                    pass

        # KPI-карточки
        if hasattr(self, "_kpi_cards"):
            for key, (card, var, color) in self._kpi_cards.items():
                card.configure(bg=T["surface"])
                for w in card.winfo_children():
                    if isinstance(w, tk.Label):
                        w.configure(bg=T["surface"])

        # Навигация
        active = self._active_page.get()
        for pid, btn in self._nav_btns.items():
            if pid == active:
                btn["frame"].configure(bg=T["nav_active"])
                btn["inner"].configure(bg=T["nav_active"])
                btn["icon"].configure(bg=T["nav_active"])
                btn["text"].configure(bg=T["nav_active"])
            else:
                btn["frame"].configure(bg=T["sidebar"])
                btn["inner"].configure(bg=T["sidebar"])
                btn["icon"].configure(bg=T["sidebar"])
                btn["text"].configure(bg=T["sidebar"])

        # Сайдбар и его дочерние элементы
        try:
            self._sidebar.configure(bg=T["sidebar"])
            self._theme_btn.configure(
                bg=T["surface2"], fg=T["text2"],
                activebackground=T["surface3"],
                text="☀  Светлая тема" if self._tn == "dark" else "◑  Тёмная тема")
        except Exception:
            pass

        # Перестраиваем лого-блок
        self._refresh_sidebar_labels(T)

        # ttk прогресс
        style = ttk.Style()
        style.configure("F.Horizontal.TProgressbar",
                        troughcolor=T["border"], background=C_ORANGE)

        # Progressbar в плане
        if hasattr(self, "_plan_inner"):
            self._plan_inner.configure(bg=T["surface"])

    def _refresh_sidebar_labels(self, T):
        try:
            self._clock_frame.configure(bg=T["sidebar"])
            for w in self._clock_frame.winfo_children():
                if isinstance(w, tk.Label):
                    fg = T["text"] if w.cget("font") and "20" in str(w.cget("font")) else T["muted"]
                    w.configure(bg=T["sidebar"], fg=fg)
        except Exception:
            pass

    # ============================================================
    # СТРОИТЕЛИ UI-КОМПОНЕНТОВ
    # ============================================================

    def _page_header(self, parent, title, subtitle=""):
        T = self._T
        hdr = self._r(tk.Frame(parent, bg=T["bg"]), "bg")
        hdr.pack(fill="x", padx=32, pady=(28, 0))

        title_lbl = tk.Label(hdr, text=title,
                              font=("Segoe UI", 18, "bold"),
                              bg=T["bg"], fg=T["text"], anchor="w")
        self._r(title_lbl, "text_bg")
        title_lbl.pack(anchor="w")

        if subtitle:
            sub = tk.Label(hdr, text=subtitle,
                           font=("Segoe UI", 9),
                           bg=T["bg"], fg=T["muted"], anchor="w")
            self._r(sub, "muted_bg")
            sub.pack(anchor="w", pady=(2, 0))

        # Тонкая разделительная линия
        tk.Frame(hdr, bg=T["border"], height=1).pack(fill="x", pady=(16, 0))
        return hdr

    def _card(self, parent, title, icon=""):
        T = self._T
        wrapper = self._r(tk.Frame(parent, bg=T["bg"]), "bg")
        wrapper.pack(fill="x", padx=32, pady=(16, 0))

        # Заголовок карточки
        hdr = self._r(tk.Frame(wrapper, bg=T["bg"]), "bg")
        hdr.pack(fill="x", pady=(0, 6))
        if icon:
            tk.Label(hdr, text=icon, font=("Segoe UI", 10),
                     bg=T["bg"], fg=C_ORANGE).pack(side="left", padx=(0, 6))
        lbl = tk.Label(hdr, text=title.upper(),
                       font=("Segoe UI", 7, "bold"),
                       bg=T["bg"], fg=T["muted"])
        self._r(lbl, "muted_bg")
        lbl.pack(side="left")

        # Тело карточки
        body = tk.Frame(wrapper, bg=T["surface"],
                        highlightbackground=T["border"],
                        highlightthickness=1)
        self._r(body, "surface")
        body.pack(fill="x")

        pad = self._r(tk.Frame(body, bg=T["surface"]), "surface")
        pad.pack(fill="x", padx=16, pady=12)
        return pad

    def _file_row(self, parent, var, cmd, placeholder="Выберите файл..."):
        T = self._T
        row = self._r(tk.Frame(parent, bg=T["surface"]), "surface")
        row.pack(fill="x")

        # Поле с иконкой
        entry_frame = tk.Frame(row, bg=T["entry_bg"],
                               highlightbackground=T["border"],
                               highlightthickness=1)
        self._r(entry_frame, "surface")
        entry_frame.pack(side="left", fill="x", expand=True, padx=(0, 10))

        tk.Label(entry_frame, text="📁",
                 font=("Segoe UI", 9),
                 bg=T["entry_bg"], fg=T["muted"],
                 padx=8).pack(side="left")

        tk.Frame(entry_frame, bg=T["border"], width=1).pack(side="left", fill="y", pady=4)

        e = tk.Entry(entry_frame, textvariable=var,
                     font=("Segoe UI", 9),
                     state="readonly", relief="flat", bd=0,
                     bg=T["entry_bg"], fg=T["entry_fg"],
                     readonlybackground=T["entry_bg"],
                     insertbackground=T["entry_fg"])
        self._r(e, "entry")
        e.pack(side="left", fill="x", expand=True, ipady=6, padx=8)

        btn = AnimButton(row, C_ORANGE, C_DARK,
                         text="Обзор",
                         font=("Segoe UI", 9, "bold"),
                         fg="#FFFFFF", command=cmd,
                         padx=18, pady=6)
        btn.pack(side="left")
        return btn

    def _log_box(self, parent, height=7):
        T = self._T
        frame = self._r(tk.Frame(parent, bg=T["bg"]), "bg")
        frame.pack(fill="both", expand=True, padx=32, pady=(0, 4))

        box = scrolledtext.ScrolledText(
            frame, font=("Consolas", 8), height=height,
            state="disabled", bg=T["log_bg"], fg=T["log_fg"],
            relief="flat", selectbackground=C_ORANGE,
            insertbackground=T["log_fg"], padx=12, pady=8,
            borderwidth=0)
        box.pack(fill="both", expand=True)
        box.tag_config("ok",   foreground=T["tag_ok"])
        box.tag_config("err",  foreground=T["tag_err"])
        box.tag_config("warn", foreground=T["tag_warn"])
        box.tag_config("ts",   foreground=T["muted"])
        return box

    def _progress_row(self, parent):
        T = self._T
        row = self._r(tk.Frame(parent, bg=T["bg"]), "bg")
        row.pack(fill="x", padx=32, pady=(0, 4))

        style = ttk.Style()
        style.configure("F.Horizontal.TProgressbar",
                        troughcolor=T["border"],
                        background=C_ORANGE, thickness=4)
        bar = ttk.Progressbar(row, style="F.Horizontal.TProgressbar",
                               mode="determinate")
        bar.pack(side="left", fill="x", expand=True, padx=(0, 12), ipady=0)

        lbl = tk.Label(row, text="0%", font=("Segoe UI", 8, "bold"),
                       bg=T["bg"], fg=C_ORANGE, width=5, anchor="e")
        self._r(lbl, "bg")
        lbl.pack(side="left")
        return bar, lbl

    def _run_btn(self, parent, text, cmd, color=C_ORANGE, hover=C_DARK, icon="▶"):
        row = self._r(tk.Frame(parent, bg=self._T["bg"]), "bg")
        row.pack(fill="x", padx=32, pady=(12, 20))
        btn = AnimButton(row, color, hover,
                         text=f"  {text}  ",
                         font=("Segoe UI", 10, "bold"),
                         fg="#FFFFFF", command=cmd,
                         padx=24, pady=12)
        btn.pack(side="left")
        return btn

    def _section_lbl(self, parent, text):
        T = self._T
        f = self._r(tk.Frame(parent, bg=T["bg"]), "bg")
        f.pack(fill="x", padx=32, pady=(16, 6))
        tk.Frame(f, bg=C_ORANGE, width=2, height=12).pack(side="left", padx=(0, 8))
        lbl = tk.Label(f, text=text,
                       font=("Segoe UI", 7, "bold"),
                       bg=T["bg"], fg=T["muted"])
        self._r(lbl, "muted_bg")
        lbl.pack(side="left")

    def _write_log(self, box, msg):
        T = self._T
        tag = ("ok"   if "✅" in msg or "Готово" in msg else
               "err"  if "❌" in msg or "ОШИБКА" in msg else
               "warn" if "⚠" in msg else None)
        ts = datetime.now().strftime("%H:%M:%S")
        box.configure(state="normal")
        box.insert("end", f"[{ts}] ", "ts")
        line = f"{msg}\n"
        box.insert("end", line, tag) if tag else box.insert("end", line)
        box.see("end")
        box.configure(state="disabled")
        self._status_var.set(msg[:100])

    # ── Прогресс-лог ─────────────────────────────────────────
    STEPS = {
        "загрузк": 1, "качеств": 2, "клиент": 3, "менеджер": 4,
        "отрасл": 5, "сезонн": 6, "rfm": 7, "лояльн": 8,
        "план": 9, "номенклатур": 10, "сигнал": 11,
        "бухгалт": 12, "сохран": 13, "граф": 14,
    }
    TOTAL = 14

    def _progress_log(self, bar, pct_lbl):
        counter = [0]
        def fn(msg):
            lo = msg.lower()
            for kw, step in self.STEPS.items():
                if kw in lo and step > counter[0]:
                    counter[0] = step
                    pct = int(step / self.TOTAL * 100)
                    self.after(0, lambda p=pct: (
                        bar.config(value=p),
                        pct_lbl.config(text=f"{p}%")))
                    break
            self._log(msg)
        return fn

    # ============================================================
    # СТРАНИЦЫ
    # ============================================================
    def _build_pages(self):
        T = self._T
        for page_id, _, _ in self.NAV_ITEMS:
            frame = self._r(tk.Frame(self._workspace, bg=T["bg"]), "bg")
            self._pages[page_id] = frame

        self._build_page_analysis()
        self._build_page_compare()
        self._build_page_plan()
        self._build_page_watcher()
        self._build_page_dashboard()

    # ============================================================
    # СТРАНИЦА 1: АНАЛИЗ
    # ============================================================
    def _build_page_analysis(self):
        p = self._pages["analysis"]

        # Скроллируемая область
        canvas = tk.Canvas(p, bg=self._T["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(p, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = self._r(tk.Frame(canvas, bg=self._T["bg"]), "bg")
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_resize(e):
            canvas.itemconfig(win_id, width=e.width)
        canvas.bind("<Configure>", _on_resize)
        inner.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))

        self._bind_mousewheel(canvas)

        self._page_header(inner, "Анализ заказов",
                          "Загрузите выгрузку из CRM и сформируйте отчёт Excel")

        c1 = self._card(inner, "Входной файл", "📂")
        self.input_var = tk.StringVar(value=self._cfg.get("last_input", ""))
        self._file_row(c1, self.input_var, self._browse_input)

        c2 = self._card(inner, "Папка для отчёта", "📁")
        self.output_dir_var = tk.StringVar(value=self._cfg.get("last_output_dir", ""))
        self._file_row(c2, self.output_dir_var, self._browse_output)

        c3 = self._card(inner, "Настройки", "⚙")
        self._build_settings_card(c3)

        self._section_lbl(inner, "ПРОГРЕСС")
        self.main_bar, self.main_pct = self._progress_row(inner)

        self._section_lbl(inner, "СВЕРКА С БУХГАЛТЕРИЕЙ")
        self._build_recon_indicator(inner)

        self._section_lbl(inner, "ЖУРНАЛ")
        self.log_box = self._log_box(inner, height=5)

        self._build_kpi_strip(inner)
        self.run_btn = self._run_btn(inner, "Запустить анализ", self._start_analysis)

    def _bind_mousewheel(self, canvas):
        def _scroll(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _scroll)

    def _build_settings_card(self, parent):
        T = self._T
        row1 = self._r(tk.Frame(parent, bg=T["surface"]), "surface")
        row1.pack(fill="x", pady=(0, 8))

        lbl = tk.Label(row1, text="Дата для группировки:",
                       font=("Segoe UI", 9), bg=T["surface"], fg=T["text2"])
        self._r(lbl, "text2_sf")
        lbl.pack(side="left")

        for val, text in [("order", "дата заказа"), ("payment", "дата оплаты")]:
            rb = tk.Radiobutton(row1, text=text,
                                variable=self._date_by, value=val,
                                bg=T["surface"], fg=T["text"],
                                activebackground=T["surface"],
                                selectcolor=T["surface2"],
                                font=("Segoe UI", 9), cursor="hand2")
            self._r(rb, "text_sf")
            rb.pack(side="left", padx=(12, 0))

        row2 = self._r(tk.Frame(parent, bg=T["surface"]), "surface")
        row2.pack(fill="x")

        lbl2 = tk.Label(row2, text="Данные бухгалтерии:",
                        font=("Segoe UI", 9), bg=T["surface"], fg=T["text2"])
        self._r(lbl2, "text2_sf")
        lbl2.pack(side="left")

        if HAS_VF_EDITOR:
            AnimButton(row2, T["surface2"], T["border"],
                       text="Верифицированные цифры",
                       font=("Segoe UI", 9), fg=T["text"],
                       command=self._open_verified_editor,
                       padx=10, pady=4).pack(side="left", padx=(12, 6))

        if HAS_EXT_EDITOR:
            AnimButton(row2, T["surface2"], T["border"],
                       text="Внешние доходы",
                       font=("Segoe UI", 9), fg=T["text"],
                       command=self._open_external_editor,
                       padx=10, pady=4).pack(side="left")

    def _build_recon_indicator(self, parent):
        T = self._T
        f = self._r(tk.Frame(parent, bg=T["bg"]), "bg")
        f.pack(fill="x", padx=32, pady=(0, 4))

        self._recon_label = tk.Label(
            f, text="  Сверка появится после запуска анализа",
            font=("Segoe UI", 9), bg=T["surface"], fg=T["muted"],
            anchor="w", justify="left", padx=12, pady=10,
            relief="flat")
        self._r(self._recon_label, "muted_sf")
        self._recon_label.pack(fill="x")

    # ── KPI-полоска ──────────────────────────────────────────
    def _build_kpi_strip(self, parent):
        T = self._T
        self._kpi_frame = self._r(tk.Frame(parent, bg=T["bg"]), "bg")

        defs = [
            ("crm",      "CRM",       "тыс. руб.",  C_ORANGE),
            ("external", "Внешние",   "тыс. руб.",  C_BLUE),
            ("grand",    "Итого",     "тыс. руб.",  "#8B5CF6"),
            ("dev",      "Отклонение","от плана",   C_GREEN),
        ]
        self._kpi_cards = {}
        for key, title, unit, color in defs:
            card = tk.Frame(self._kpi_frame, bg=T["surface"],
                            highlightbackground=T["border"],
                            highlightthickness=1)
            card.pack(side="left", fill="x", expand=True, padx=(0, 8), ipady=8, ipadx=12)

            # Верхняя цветная полоска
            tk.Frame(card, bg=color, height=2).pack(fill="x")

            tk.Label(card, text=title,
                     font=("Segoe UI", 7, "bold"),
                     bg=T["surface"], fg=color).pack(anchor="w", padx=10, pady=(8, 0))

            val_var = tk.StringVar(value="—")
            tk.Label(card, textvariable=val_var,
                     font=("Segoe UI", 16, "bold"),
                     bg=T["surface"], fg=T["text"]).pack(anchor="w", padx=10)

            tk.Label(card, text=unit,
                     font=("Segoe UI", 7),
                     bg=T["surface"], fg=T["muted"]).pack(anchor="w", padx=10, pady=(0, 8))

            self._kpi_cards[key] = (card, val_var, color)

    def _show_kpi_cards(self, result: dict):
        if not hasattr(self, "_kpi_frame"):
            return
        T = self._T

        crm_k   = result.get("crm_total", 0)
        ext_k   = result.get("full_external_total", result.get("external_total", 0))
        grand_k = result.get("grand_total", 0)
        pct     = result.get("deviation_pct")

        def _fmt(v):
            return f"{v:,.0f}".replace(",", " ")

        card, var, _ = self._kpi_cards["crm"]
        var.set(_fmt(crm_k))

        card, var, _ = self._kpi_cards["external"]
        var.set(_fmt(ext_k))

        card, var, _ = self._kpi_cards["grand"]
        var.set(_fmt(grand_k))

        card, var, _ = self._kpi_cards["dev"]
        if pct is not None:
            sign = "+" if pct >= 0 else ""
            var.set(f"{sign}{pct:.2f}%")
            dev_color = C_GREEN if abs(pct) < 2 else (C_AMBER if abs(pct) < 5 else C_RED)
        else:
            var.set("—")
            dev_color = T["muted"]
        card.configure(highlightbackground=dev_color)
        for w in card.winfo_children():
            if isinstance(w, tk.Label) and "16" in str(w.cget("font")):
                w.configure(fg=dev_color)
                break

        for key, (c, v, color) in self._kpi_cards.items():
            c.configure(bg=T["surface"])
            for w in c.winfo_children():
                if isinstance(w, tk.Label):
                    w.configure(bg=T["surface"])

        if not self._kpi_frame.winfo_ismapped():
            self._kpi_frame.pack(fill="x", padx=32, pady=(0, 4),
                                 before=self.run_btn.master)

    def _open_external_editor(self):
        path = os.path.join(APP_DIR, "external_income.json")
        if not os.path.exists(path):
            messagebox.showerror("Нет файла",
                f"Файл external_income.json не найден:\n{path}")
            return
        ExternalIncomeEditor(self, path, theme=self._T)

    def _open_verified_editor(self):
        path = os.path.join(APP_DIR, "verified_figures.json")
        if not os.path.exists(path):
            default = {
                "_описание": "Верифицированные годовые итоги из бухгалтерии.",
                "_год": 2025,
                "total_with_prog": 482404000,
                "total_with_barter_no_prog": 363000166,
                "advertising_no_events": 243005820,
                "programmatic_external": 119403897,
                "other_external_income": 41057102,
            }
            try:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(default, f, ensure_ascii=False, indent=2)
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))
                return
        VerifiedFiguresEditor(self, path, theme=self._T)

    # ============================================================
    # СТРАНИЦА 2: СРАВНЕНИЕ
    # ============================================================
    def _build_page_compare(self):
        p = self._pages["compare"]

        canvas = tk.Canvas(p, bg=self._T["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(p, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = self._r(tk.Frame(canvas, bg=self._T["bg"]), "bg")
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(win_id, width=e.width))
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        self._page_header(inner, "Сравнение периодов",
                          "Сравните показатели двух произвольных периодов")

        def period_block(title, path_var, label_var):
            c = self._card(inner, title)
            self._file_row(c, path_var,
                           lambda v=path_var: self._browse_cmp(v))
            row = self._r(tk.Frame(c, bg=self._T["surface"]), "surface")
            row.pack(fill="x", pady=(8, 0))
            lbl = tk.Label(row, text="Метка периода:",
                           font=("Segoe UI", 9),
                           bg=self._T["surface"], fg=self._T["text2"])
            self._r(lbl, "text2_sf")
            lbl.pack(side="left")
            e = tk.Entry(row, textvariable=label_var,
                         font=("Segoe UI", 9), width=22,
                         relief="flat", bd=0,
                         highlightbackground=self._T["border"],
                         highlightthickness=1,
                         bg=self._T["entry_bg"], fg=self._T["entry_fg"])
            self._r(e, "entry")
            e.pack(side="left", padx=12, ipady=4)

        self.cmp_path_a  = tk.StringVar()
        self.cmp_label_a = tk.StringVar(value="Период А")
        self.cmp_path_b  = tk.StringVar()
        self.cmp_label_b = tk.StringVar(value="Период Б")
        period_block("Период А  ←", self.cmp_path_a, self.cmp_label_a)
        period_block("Период Б  →", self.cmp_path_b, self.cmp_label_b)

        co = self._card(inner, "Папка для отчёта", "📁")
        self.cmp_output_dir = tk.StringVar(
            value=self._cfg.get("last_output_dir", ""))
        self._file_row(co, self.cmp_output_dir, self._browse_cmp_output)

        self._section_lbl(inner, "ЖУРНАЛ")
        self.cmp_log_box = self._log_box(inner, height=5)
        self.cmp_run_btn = self._run_btn(inner, "Сравнить периоды",
                                          self._start_comparison)

    # ============================================================
    # СТРАНИЦА 3: ПЛАН
    # ============================================================
    def _build_page_plan(self):
        p = self._pages["plan"]
        T = self._T

        self._page_header(p, "Планы менеджеров",
                          "Плановые суммы в тыс. руб. — сохраняются автоматически")

        outer = self._r(tk.Frame(p, bg=T["bg"]), "bg")
        outer.pack(fill="both", expand=True, padx=32, pady=(12, 0))

        canvas = tk.Canvas(outer, bg=T["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)

        self._plan_inner = self._r(
            tk.Frame(canvas, bg=T["surface"]), "surface")
        self._plan_inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._plan_inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # Заголовок таблицы
        hdr = tk.Frame(self._plan_inner, bg=T["surface2"])
        hdr.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Frame(hdr, bg=C_ORANGE, width=3).pack(side="left", fill="y")
        tk.Label(hdr, text="  Менеджер",
                 font=("Segoe UI", 8, "bold"),
                 bg=T["surface2"], fg=T["text2"],
                 width=32, anchor="w", pady=8).pack(side="left")
        tk.Label(hdr, text="План, тыс. руб.",
                 font=("Segoe UI", 8, "bold"),
                 bg=T["surface2"], fg=T["text2"],
                 width=18, pady=8).pack(side="left")

        saved    = self._cfg.get("manager_plan", {})
        managers = list(dict.fromkeys([
            "Петрова Э.", "Егорова И.", "Демахина Е.",
            "Дубровская Т.", "Долгополова Н.", "Демидова Г.",
            "Сосновик А.", "Киселёва Е.", "Лея К.",
            "Борисенко Е.", "Сорокина Т.", "Шевчук Е.", "Карпус О.",
        ] + list(saved.keys())))
        for i, name in enumerate(managers, 1):
            self._add_plan_row(name, str(saved.get(name, "")), i)

        # Строка добавления
        add_frame = self._r(tk.Frame(p, bg=T["bg"]), "bg")
        add_frame.pack(fill="x", padx=32, pady=(8, 20))

        tk.Label(add_frame, text="Добавить менеджера:",
                 font=("Segoe UI", 9),
                 bg=T["bg"], fg=T["text2"]).pack(side="left")
        self._r(add_frame.winfo_children()[-1], "text2_bg")

        self._new_mgr = tk.StringVar()
        ne = tk.Entry(add_frame, textvariable=self._new_mgr,
                      font=("Segoe UI", 9), width=24,
                      relief="flat", bd=0,
                      highlightbackground=T["border"],
                      highlightthickness=1,
                      bg=T["entry_bg"], fg=T["entry_fg"])
        self._r(ne, "entry")
        ne.pack(side="left", padx=12, ipady=5)

        AnimButton(add_frame, C_ORANGE, C_DARK,
                   text="+ Добавить",
                   font=("Segoe UI", 9, "bold"),
                   fg="#FFFFFF", command=self._add_manager_ui,
                   padx=14, pady=5).pack(side="left")

    def _add_plan_row(self, name, val="", idx=None):
        T = self._T
        if idx is None:
            idx = len(self._plan_vars) + 1
        even = idx % 2 == 0
        bg = T["surface"] if even else T["surface2"]

        lbl = tk.Label(self._plan_inner, text=f"  {name}",
                       font=("Segoe UI", 9), bg=bg,
                       fg=T["text"], anchor="w", pady=6)
        lbl.grid(row=idx, column=0, sticky="ew")
        self._plan_inner.columnconfigure(0, weight=1)

        var = tk.StringVar(value=val)
        var.trace_add("write",
                      lambda *_, n=name, v=var: self._save_plan(n, v))
        self._plan_vars[name] = var

        ent = tk.Entry(self._plan_inner, textvariable=var,
                       font=("Segoe UI", 9), width=16,
                       relief="flat", bd=0,
                       highlightbackground=T["border"],
                       highlightthickness=1,
                       bg=bg, fg=T["text"],
                       insertbackground=T["text"])
        ent.grid(row=idx, column=1, pady=2, padx=8)

    def _add_manager_ui(self):
        name = self._new_mgr.get().strip()
        if not name:
            return
        if name in self._plan_vars:
            messagebox.showinfo("", f"«{name}» уже есть.")
            return
        self._add_plan_row(name)
        self._new_mgr.set("")

    def _save_plan(self, name, var):
        plans = self._cfg.setdefault("manager_plan", {})
        val = var.get().strip()
        if val:
            plans[name] = val
        else:
            plans.pop(name, None)
        save_config(self._cfg)

    def _get_plan_rub(self):
        result = {}
        for name, var in self._plan_vars.items():
            val = var.get().strip().replace(",", ".")
            if val:
                try:
                    result[name] = float(val) * 1000
                except ValueError:
                    pass
        return result or None

    # ============================================================
    # СТРАНИЦА 4: НАБЛЮДАТЕЛЬ
    # ============================================================
    def _build_page_watcher(self):
        p = self._pages["watcher"]
        T = self._T

        self._page_header(p, "Наблюдатель",
                          "Авто-анализ через 5 минут после появления нового .xlsx")

        cw = self._card(p, "Папка для наблюдения", "👁")
        self.watch_dir_var = tk.StringVar(
            value=self._cfg.get("watch_dir", ""))
        self._file_row(cw, self.watch_dir_var, self._browse_watch_dir)

        co = self._card(p, "Папка для отчётов", "📁")
        self.watch_out_var = tk.StringVar(
            value=self._cfg.get("watch_output_dir", ""))
        self._file_row(co, self.watch_out_var, self._browse_watch_out)

        # Статус-блок
        sf = self._r(tk.Frame(p, bg=T["bg"]), "bg")
        sf.pack(fill="x", padx=32, pady=(20, 8))

        status_card = self._r(
            tk.Frame(sf, bg=T["surface"],
                     highlightbackground=T["border"],
                     highlightthickness=1), "surface")
        status_card.pack(fill="x")

        status_inner = self._r(
            tk.Frame(status_card, bg=T["surface"], padx=16, pady=12), "surface")
        status_inner.pack(fill="x")

        self._status_dot = tk.Label(
            status_inner, text="●",
            font=("Segoe UI", 16),
            bg=T["surface"], fg=T["muted"])
        self._r(self._status_dot, "surface")
        self._status_dot.pack(side="left", padx=(0, 12))

        status_text = self._r(tk.Frame(status_inner, bg=T["surface"]), "surface")
        status_text.pack(side="left")

        self.watch_status_var = tk.StringVar(value="Остановлен")
        self.watch_status_lbl = tk.Label(
            status_text, textvariable=self.watch_status_var,
            font=("Segoe UI", 12, "bold"),
            bg=T["surface"], fg=T["muted"])
        self._r(self.watch_status_lbl, "surface")
        self.watch_status_lbl.pack(anchor="w")

        tk.Label(status_text, text="Нажмите «Запустить» для начала слежения",
                 font=("Segoe UI", 8),
                 bg=T["surface"], fg=T["muted"]).pack(anchor="w")
        self._r(status_text.winfo_children()[-1], "muted_sf")

        self._section_lbl(p, "ЖУРНАЛ")
        self.watch_log_box = self._log_box(p, height=8)

        # Кнопки
        bf = self._r(tk.Frame(p, bg=T["bg"]), "bg")
        bf.pack(fill="x", padx=32, pady=(8, 20))

        self.watch_start_btn = AnimButton(
            bf, C_GREEN, C_GREEN_D,
            text="  ▶  Запустить",
            font=("Segoe UI", 10, "bold"),
            fg="#FFFFFF", command=self._start_watcher,
            padx=20, pady=11)
        self.watch_start_btn.pack(side="left", padx=(0, 10))

        self.watch_stop_btn = AnimButton(
            bf, C_RED, C_RED_D,
            text="  ⏹  Остановить",
            font=("Segoe UI", 10, "bold"),
            fg="#FFFFFF", command=self._stop_watcher,
            padx=20, pady=11, state="disabled")
        self.watch_stop_btn.pack(side="left")

    # ============================================================
    # СТРАНИЦА 5: ДАШБОРД
    # ============================================================
    def _build_page_dashboard(self):
        p = self._pages["dashboard"]
        T = self._T

        self._page_header(p, "HTML-Дашборд",
                          "Интерактивный дашборд с графиками — открывается в браузере")

        c1 = self._card(p, "Входной файл", "📂")
        self.dash_input_var = tk.StringVar(
            value=self._cfg.get("last_input", ""))
        self._file_row(c1, self.dash_input_var, self._browse_dash_input)

        # Превью-карточки
        preview_wrap = self._r(tk.Frame(p, bg=T["bg"]), "bg")
        preview_wrap.pack(fill="x", padx=32, pady=(20, 0))

        self._section_lbl(p, "СОДЕРЖИМОЕ ДАШБОРДА")

        grid = self._r(tk.Frame(p, bg=T["bg"]), "bg")
        grid.pack(fill="x", padx=32, pady=(4, 0))

        items = [
            ("📈", "Выручка по месяцам",  "Линейный с градиентом",  C_ORANGE),
            ("🏆", "Топ-10 клиентов",     "Горизонтальный bar",     C_BLUE),
            ("👤", "Топ-10 менеджеров",   "Вертикальный bar",       C_GREEN),
            ("🍩", "Выручка по отраслям", "Doughnut-диаграмма",     "#8B5CF6"),
        ]
        for i, (icon, title, desc, color) in enumerate(items):
            card = tk.Frame(grid, bg=T["surface"],
                            highlightbackground=T["border"],
                            highlightthickness=1)
            card.grid(row=0, column=i, sticky="nsew",
                      padx=(0, 8) if i < 3 else 0, pady=4, ipady=12)
            grid.columnconfigure(i, weight=1)

            tk.Frame(card, bg=color, height=2).pack(fill="x")
            tk.Label(card, text=icon, font=("Segoe UI", 24),
                     bg=T["surface"]).pack(pady=(12, 4))
            tk.Label(card, text=title,
                     font=("Segoe UI", 9, "bold"),
                     bg=T["surface"], fg=T["text"]).pack()
            tk.Label(card, text=desc,
                     font=("Segoe UI", 8),
                     bg=T["surface"], fg=T["muted"]).pack(pady=(2, 8))

        self._section_lbl(p, "ЖУРНАЛ")
        self.dash_log_box = self._log_box(p, height=5)

        self.dash_btn = self._run_btn(
            p, "Открыть дашборд в браузере",
            self._start_dashboard,
            color="#1D4ED8", hover="#1E40AF")

    # ============================================================
    # BROWSE / LOG
    # ============================================================
    def _browse_input(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.input_var.set(path)
            self._cfg["last_input"] = path
            if not self.output_dir_var.get():
                d = os.path.dirname(path)
                self.output_dir_var.set(d)
                self._cfg["last_output_dir"] = d
            save_config(self._cfg)

    def _browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir_var.set(path)
            self._cfg["last_output_dir"] = path
            save_config(self._cfg)

    def _browse_cmp(self, var):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            var.set(path)
            if not self.cmp_output_dir.get():
                self.cmp_output_dir.set(os.path.dirname(path))

    def _browse_cmp_output(self):
        path = filedialog.askdirectory()
        if path:
            self.cmp_output_dir.set(path)

    def _browse_watch_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.watch_dir_var.set(path)
            self._cfg["watch_dir"] = path
            if not self.watch_out_var.get():
                self.watch_out_var.set(path)
                self._cfg["watch_output_dir"] = path
            save_config(self._cfg)

    def _browse_watch_out(self):
        path = filedialog.askdirectory()
        if path:
            self.watch_out_var.set(path)
            self._cfg["watch_output_dir"] = path
            save_config(self._cfg)

    def _browse_dash_input(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.dash_input_var.set(path)
            self._cfg["last_input"] = path
            self.input_var.set(path)
            save_config(self._cfg)

    def _log(self, msg):
        self._write_log(self.log_box, msg)

    def _cmp_log(self, msg):
        self._write_log(self.cmp_log_box, msg)

    def _watch_log(self, msg):
        self.after(0, lambda: self._write_log(self.watch_log_box, msg))

    def _dash_log(self, msg):
        self._write_log(self.dash_log_box, msg)

    # ============================================================
    # АНАЛИЗ
    # ============================================================
    def _start_analysis(self):
        inp = self.input_var.get().strip()
        out = self.output_dir_var.get().strip()
        if not inp:
            messagebox.showwarning("", "Выберите входной файл.")
            return
        if not out:
            messagebox.showwarning("", "Укажите папку для отчёта.")
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(out, f"аналитика_заказов_{ts}.xlsx")
        plan = self._get_plan_rub()
        self.main_bar.config(value=0)
        self.main_pct.config(text="0%")
        self.run_btn.configure(state="disabled", text="  Выполняется...  ")
        self._start_anim()
        self._recon_label.config(text="  Считаю...", fg=self._T["muted"])
        log_fn = self._progress_log(self.main_bar, self.main_pct)
        threading.Thread(target=self._thread_analysis,
                         args=(inp, out_path, plan, log_fn),
                         daemon=True).start()

    def _thread_analysis(self, inp, out_path, plan, log_fn):
        try:
            log_fn(f"Файл: {os.path.basename(inp)}")
            date_by = self._date_by.get()
            result = run_analytics(inp, out_path,
                                   log=log_fn,
                                   manager_plan=plan,
                                   date_by=date_by)
            log_fn(f"✅ Готово: {out_path}")
            self.after(0, lambda: (
                self.main_bar.config(value=100),
                self.main_pct.config(text="100%")))

            if isinstance(result, dict):
                path = result.get("output_path", out_path)
                self.after(0, lambda r=result: self._update_recon(r))
                self.after(0, lambda r=result: self._show_kpi_cards(r))
                folder = os.path.dirname(path)
                self.after(200, lambda p=path, f=folder: self._toast(
                    f"Отчёт готов: {os.path.basename(p)}",
                    kind="ok",
                    action_text="Открыть папку",
                    action_cmd=lambda: os.startfile(f),
                    duration=6000))
            else:
                path = result or out_path
                folder = os.path.dirname(path)
                self.after(200, lambda p=path, f=folder: self._toast(
                    f"Отчёт готов: {os.path.basename(p)}",
                    kind="ok",
                    action_text="Открыть папку",
                    action_cmd=lambda: os.startfile(f),
                    duration=6000))

            self.after(0, lambda: self._on_success(path))
        except Exception as e:
            err = str(e)
            log_fn(f"❌ ОШИБКА: {err}")
            self.after(200, lambda m=err: self._toast(
                f"Ошибка: {m[:80]}", kind="err", duration=7000))
            self.after(0, lambda: messagebox.showerror("Ошибка", err))
        finally:
            self._stop_anim()
            self.after(0, lambda: self.run_btn.configure(
                state="normal", text="  Запустить анализ  "))

    def _update_recon(self, result):
        pct      = result.get("deviation_pct")
        grand    = result.get("grand_total", 0)
        verified = result.get("verified_total", 0)
        date_by  = result.get("date_by", "order")

        if pct is None or verified <= 0:
            self._recon_label.config(
                text=f"  CRM + внешние: {grand:,.0f} тыс. руб.".replace(",", " "),
                fg=self._T["text"])
            return

        if abs(pct) < 2:
            color, icon, verdict = C_GREEN, "✓", "отлично"
        elif abs(pct) < 5:
            color, icon, verdict = C_AMBER, "!", "в норме"
        else:
            color, icon, verdict = C_RED, "✕", "большое расхождение"

        date_label = "дата оплаты" if date_by == "payment" else "дата заказа"
        text = (f"  {icon}  {date_label.capitalize()}: "
                f"CRM+внешние {grand:,.0f} тыс.  │  "
                f"Цель {verified:,.0f} тыс.  │  "
                f"{pct:+.2f}% — {verdict}").replace(",", " ")
        self._recon_label.config(text=text, fg=color)

    # ============================================================
    # СРАВНЕНИЕ
    # ============================================================
    def _start_comparison(self):
        pa = self.cmp_path_a.get().strip()
        pb = self.cmp_path_b.get().strip()
        od = self.cmp_output_dir.get().strip()
        la = self.cmp_label_a.get().strip() or "Период А"
        lb = self.cmp_label_b.get().strip() or "Период Б"
        if not pa:
            messagebox.showwarning("", "Выберите файл периода А.")
            return
        if not pb:
            messagebox.showwarning("", "Выберите файл периода Б.")
            return
        if not od:
            messagebox.showwarning("", "Укажите папку для отчёта.")
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(od, f"сравнение_{ts}.xlsx")
        self.cmp_run_btn.configure(state="disabled", text="  Выполняется...  ")
        self._start_anim()
        threading.Thread(target=self._thread_comparison,
                         args=(pa, pb, out_path, la, lb),
                         daemon=True).start()

    def _thread_comparison(self, pa, pb, out_path, la, lb):
        try:
            self._cmp_log(f"Сравниваю «{la}» и «{lb}»")
            run_comparison(pa, pb, out_path, la, lb, log=self._cmp_log)
            self._cmp_log(f"✅ Готово: {out_path}")
            folder = os.path.dirname(out_path)
            self.after(200, lambda f=folder: self._toast(
                f"Сравнение готово: {os.path.basename(out_path)}",
                kind="ok",
                action_text="Открыть папку",
                action_cmd=lambda: os.startfile(f),
                duration=6000))
            self.after(0, lambda: self._on_success(out_path))
        except Exception as e:
            err = str(e)
            self._cmp_log(f"❌ ОШИБКА: {err}")
            self.after(200, lambda m=err: self._toast(
                f"Ошибка: {m[:80]}", kind="err", duration=7000))
            self.after(0, lambda: messagebox.showerror("Ошибка", err))
        finally:
            self._stop_anim()
            self.after(0, lambda: self.cmp_run_btn.configure(
                state="normal", text="  Сравнить периоды  "))

    # ============================================================
    # НАБЛЮДАТЕЛЬ
    # ============================================================
    def _start_watcher(self):
        wd = self.watch_dir_var.get().strip()
        wo = self.watch_out_var.get().strip()
        if not wd or not wo:
            messagebox.showwarning("", "Укажите обе папки.")
            return
        if not os.path.isdir(wd):
            messagebox.showerror("", f"Папка не существует:\n{wd}")
            return
        self._watcher = FolderWatcher(wd, wo, self._watch_log)
        self._watcher.start()
        self.watch_status_var.set("Наблюдатель активен")
        self.watch_status_lbl.config(fg=C_GREEN)
        self._status_dot.config(fg=C_GREEN)
        self.watch_start_btn.configure(state="disabled")
        self.watch_stop_btn.configure(state="normal")
        self._toast("Наблюдатель запущен", kind="info", duration=3000)

    def _stop_watcher(self):
        if self._watcher:
            self._watcher.stop()
            self._watcher = None
        T = self._T
        self.watch_status_var.set("Остановлен")
        self.watch_status_lbl.config(fg=T["muted"])
        self._status_dot.config(fg=T["muted"])
        self.watch_start_btn.configure(state="normal")
        self.watch_stop_btn.configure(state="disabled")

    # ============================================================
    # ДАШБОРД
    # ============================================================
    def _start_dashboard(self):
        inp = self.dash_input_var.get().strip()
        if not inp:
            messagebox.showwarning("", "Выберите входной файл.")
            return
        self.dash_btn.configure(state="disabled", text="  Генерирую...  ")
        self._start_anim()
        threading.Thread(target=self._thread_dashboard,
                         args=(inp,), daemon=True).start()

    def _thread_dashboard(self, inp):
        try:
            generate_dashboard(inp, log=self._dash_log)
        except Exception as e:
            self._dash_log(f"❌ ОШИБКА: {e}")
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self._stop_anim()
            self.after(0, lambda: self.dash_btn.configure(
                state="normal", text="  Открыть дашборд в браузере  "))

    def _on_success(self, path):
        self._status_var.set(f"Готово: {os.path.basename(path)}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
