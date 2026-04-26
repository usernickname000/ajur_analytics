# ============================================================
# app.py — Аналитика заказов, Фонтанка.ру  v2.1
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

# Опциональные модули — редакторы JSON (подключаем если файлы есть)
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

C_ORANGE  = "#F38120"
C_DARK    = "#D06A10"
C_GREEN   = "#22C55E"
C_GREEN_D = "#16A34A"
C_RED     = "#EF4444"
C_RED_D   = "#DC2626"
C_AMBER   = "#F39C12"

THEMES = {
    "light": {
        "bg":         "#F8F8F8",
        "surface":    "#FFFFFF",
        "surface2":   "#F0F0F0",
        "border":     "#E2E2E2",
        "text":       "#0A0A0A",
        "text2":      "#555555",
        "muted":      "#888888",
        "log_bg":     "#111111",
        "log_fg":     "#E8E8E8",
        "entry_bg":   "#FFFFFF",
        "entry_fg":   "#0A0A0A",
        "sb_bg":      "#EFEFEF",
        "plan_alt":   "#F8F0EA",
        "tag_ok":     "#16A34A",
        "tag_err":    "#DC2626",
        "tag_warn":   "#D97706",
    },
    "dark": {
        "bg":         "#0A0A0A",
        "surface":    "#141414",
        "surface2":   "#1E1E1E",
        "border":     "#2A2A2A",
        "text":       "#F0F0F0",
        "text2":      "#AAAAAA",
        "muted":      "#666666",
        "log_bg":     "#000000",
        "log_fg":     "#DDDDDD",
        "entry_bg":   "#1E1E1E",
        "entry_fg":   "#F0F0F0",
        "sb_bg":      "#000000",
        "plan_alt":   "#1A1A1A",
        "tag_ok":     "#22C55E",
        "tag_err":    "#EF4444",
        "tag_warn":   "#F59E0B",
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


# ── Анимированная кнопка ─────────────────────────────────────
class AnimButton(tk.Button):
    STEPS = 8
    DELAY = 12

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
    def __init__(self):
        super().__init__()
        self.title("Аналитика заказов — Фонтанка.ру")
        self.geometry("820x780")
        self.resizable(False, False)

        if os.path.exists(ICON_PATH):
            try:
                self.iconbitmap(default=ICON_PATH)
            except Exception:
                pass

        self._cfg  = load_config()
        self._tn   = self._cfg.get("theme", "light")
        self._T    = THEMES[self._tn]
        self._tw   = []  # themed widgets list
        self._watcher = None
        self._plan_vars = {}
        self._anim_running = False
        self._anim_step = 0
        # НОВОЕ: выбор базы даты
        self._date_by = tk.StringVar(value=self._cfg.get("date_by", "order"))

        self.configure(bg=self._T["bg"])
        self._build_header()
        self._build_tabs()
        self._build_statusbar()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        if self._watcher:
            self._watcher.stop()
        self._cfg["theme"] = self._tn
        self._cfg["date_by"] = self._date_by.get()
        save_config(self._cfg)
        self.destroy()

    # ── Регистрация виджета для перекраски ───────────────────
    def _r(self, w, role):
        self._tw.append((w, role))
        return w

    # ── Применить тему ───────────────────────────────────────
    def _apply_theme(self):
        T = self._T
        self.configure(bg=T["bg"])

        role_map = {
            "bg":      lambda w: w.configure(bg=T["bg"]),
            "surface": lambda w: w.configure(bg=T["surface"]),
            "surface2":lambda w: w.configure(bg=T["surface2"]),
            "border":  lambda w: w.configure(bg=T["border"]),
            "text_bg": lambda w: w.configure(bg=T["bg"], fg=T["text"]),
            "text2_bg":lambda w: w.configure(bg=T["bg"], fg=T["text2"]),
            "muted_bg":lambda w: w.configure(bg=T["bg"], fg=T["muted"]),
            "text_sf": lambda w: w.configure(bg=T["surface"], fg=T["text"]),
            "muted_sf":lambda w: w.configure(bg=T["surface"], fg=T["muted"]),
            "entry":   lambda w: w.configure(bg=T["entry_bg"], fg=T["entry_fg"],
                                              readonlybackground=T["entry_bg"],
                                              insertbackground=T["entry_fg"],
                                              disabledbackground=T["entry_bg"]),
            "sb":      lambda w: w.configure(bg=T["sb_bg"], fg=T["muted"]),
            "plan_e":  lambda w: w.configure(bg=T["surface"],  fg=T["text"]),
            "plan_o":  lambda w: w.configure(bg=T["plan_alt"], fg=T["text"]),
            "plan_le": lambda w: w.configure(bg=T["surface"],  fg=T["text"]),
            "plan_lo": lambda w: w.configure(bg=T["plan_alt"], fg=T["text"]),
        }

        for (w, role) in self._tw:
            try:
                if role in role_map:
                    role_map[role](w)
            except Exception:
                pass

        # Статус-бар
        try:
            self._sb_f.configure(bg=T["sb_bg"])
            self._sb_lbl.configure(bg=T["sb_bg"], fg=T["muted"])
            self._sb_ver.configure(bg=T["sb_bg"], fg=T["muted"])
            self._anim_lbl.configure(bg=T["sb_bg"])
        except Exception:
            pass

        # Логи — обновляем теги
        for box in [self.log_box, self.cmp_log_box, self.watch_log_box]:
            try:
                box.configure(bg=T["log_bg"], fg=T["log_fg"])
                box.tag_config("ok",   foreground=T["tag_ok"])
                box.tag_config("err",  foreground=T["tag_err"])
                box.tag_config("warn", foreground=T["tag_warn"])
            except Exception:
                pass

        # ttk стили
        style = ttk.Style()
        style.configure("F.TNotebook", background=T["bg"])
        style.configure("F.TNotebook.Tab",
                        background=T["surface2"],
                        foreground=T["muted"])
        style.map("F.TNotebook.Tab",
                  background=[("selected", T["surface"])],
                  foreground=[("selected", C_ORANGE)])
        style.configure("F.Horizontal.TProgressbar",
                        troughcolor=T["border"],
                        background=C_ORANGE)

    def _toggle_theme(self):
        self._tn = "dark" if self._tn == "light" else "light"
        self._T  = THEMES[self._tn]
        self._cfg["theme"] = self._tn
        save_config(self._cfg)
        self._apply_theme()
        self._theme_btn.config(
            text="☀️" if self._tn == "dark" else "🌙",
            bg=C_ORANGE, activebackground=C_DARK)

    # ============================================================
    # ШАПКА
    # ============================================================
    def _build_header(self):
        hdr = tk.Frame(self, bg=C_ORANGE, height=68)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        left = tk.Frame(hdr, bg=C_ORANGE)
        left.pack(side="left", fill="y")

        tk.Label(left, text="фонтанка.ру",
                 font=("Georgia", 19, "bold italic"),
                 bg=C_ORANGE, fg="#FFFFFF",
                 padx=20).pack(side="left", pady=16)

        sep = tk.Frame(hdr, bg="#D8711A", width=1)
        sep.pack(side="left", fill="y", pady=12, padx=2)

        tk.Label(hdr, text="АНАЛИТИКА  /  КОММЕРЧЕСКИЙ ОТДЕЛ",
                 font=("Segoe UI", 8, "bold"),
                 bg=C_ORANGE, fg="#FFD8A8",
                 padx=14).pack(side="left", pady=22)

        right = tk.Frame(hdr, bg=C_ORANGE)
        right.pack(side="right", fill="y", padx=16)

        icon = "🌙" if self._tn == "light" else "☀️"
        self._theme_btn = tk.Button(
            right, text=icon,
            font=("Segoe UI", 16),
            bg=C_ORANGE, fg="#FFFFFF",
            activebackground=C_DARK,
            activeforeground="#FFFFFF",
            relief="flat", bd=0, cursor="hand2",
            command=self._toggle_theme)
        self._theme_btn.pack(side="right", padx=(8, 0), pady=18)

        clock_frame = tk.Frame(right, bg=C_ORANGE)
        clock_frame.pack(side="right", fill="y", pady=10)

        self._clock_time = tk.StringVar()
        self._clock_date = tk.StringVar()

        tk.Label(clock_frame, textvariable=self._clock_time,
                 font=("Segoe UI", 15, "bold"),
                 bg=C_ORANGE, fg="#FFFFFF").pack(anchor="e")
        tk.Label(clock_frame, textvariable=self._clock_date,
                 font=("Segoe UI", 8),
                 bg=C_ORANGE, fg="#FFD8A8").pack(anchor="e")

        self._update_clock()

    def _update_clock(self):
        try:
            msk = timezone(timedelta(hours=3))
            now = datetime.now(tz=msk)
        except Exception:
            now = datetime.now()
        self._clock_time.set(now.strftime("%H:%M  МСК"))
        self._clock_date.set(now.strftime("%d  %B  %Y").upper())
        self.after(60000, self._update_clock)

    # ============================================================
    # ВКЛАДКИ
    # ============================================================
    def _build_tabs(self):
        T = self._T
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("F.TNotebook", background=T["bg"], borderwidth=0)
        style.configure("F.TNotebook.Tab",
                        font=("Segoe UI", 10), padding=[16, 8],
                        background=T["surface2"], foreground=T["muted"])
        style.map("F.TNotebook.Tab",
                  background=[("selected", T["surface"])],
                  foreground=[("selected", C_ORANGE)],
                  font=[("selected", ("Segoe UI", 10, "bold"))])

        self.notebook = ttk.Notebook(self, style="F.TNotebook")
        self.notebook.pack(fill="both", expand=True)

        self.tab_main    = self._r(tk.Frame(self.notebook, bg=T["bg"]), "bg")
        self.tab_compare = self._r(tk.Frame(self.notebook, bg=T["bg"]), "bg")
        self.tab_plan    = self._r(tk.Frame(self.notebook, bg=T["bg"]), "bg")
        self.tab_watch   = self._r(tk.Frame(self.notebook, bg=T["bg"]), "bg")
        self.tab_dash    = self._r(tk.Frame(self.notebook, bg=T["bg"]), "bg")

        self.notebook.add(self.tab_main,    text="  Анализ  ")
        self.notebook.add(self.tab_compare, text="  Сравнение  ")
        self.notebook.add(self.tab_plan,    text="  План  ")
        self.notebook.add(self.tab_watch,   text="  Наблюдатель  ")
        self.notebook.add(self.tab_dash,    text="  Дашборд  ")

        self._build_tab_main()
        self._build_tab_compare()
        self._build_tab_plan()
        self._build_tab_watch()
        self._build_tab_dash()

    # ============================================================
    # СТАТУС-БАР
    # ============================================================
    def _build_statusbar(self):
        T = self._T
        self._sb_f = tk.Frame(self, bg=T["sb_bg"], height=26)
        self._sb_f.pack(fill="x", side="bottom")
        self._sb_f.pack_propagate(False)

        tk.Frame(self._sb_f, bg=C_ORANGE, height=1).pack(fill="x", side="top")

        self._status_var = tk.StringVar(value="Готов к работе")
        self._sb_lbl = tk.Label(
            self._sb_f, textvariable=self._status_var,
            font=("Segoe UI", 8), bg=T["sb_bg"], fg=T["muted"], anchor="w")
        self._sb_lbl.pack(side="left", padx=12)

        self._anim_lbl = tk.Label(
            self._sb_f, text="", font=("Segoe UI", 9),
            bg=T["sb_bg"], fg=C_ORANGE)
        self._anim_lbl.pack(side="right", padx=12)

        self._sb_ver = tk.Label(
            self._sb_f, text="v2.1  ·  Фонтанка.ру",
            font=("Segoe UI", 8), bg=T["sb_bg"], fg=T["muted"])
        self._sb_ver.pack(side="right", padx=12)

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
    # СТРОИТЕЛИ ВИДЖЕТОВ
    # ============================================================
    def _card(self, parent, title):
        T = self._T
        outer = self._r(tk.Frame(parent, bg=T["bg"]), "bg")
        outer.pack(fill="x", padx=16, pady=(4, 2))

        hdr_row = self._r(tk.Frame(outer, bg=T["surface"]), "surface")
        hdr_row.pack(fill="x")

        tk.Frame(hdr_row, bg=C_ORANGE, width=3).pack(side="left", fill="y")
        tk.Label(hdr_row, text=f"  {title}",
                 font=("Segoe UI", 8, "bold"),
                 bg=T["surface"], fg=C_ORANGE,
                 anchor="w", pady=6).pack(side="left", fill="x", expand=True)

        body = self._r(
            tk.Frame(outer, bg=T["surface"],
                     highlightbackground=T["border"],
                     highlightthickness=1), "surface")
        body.pack(fill="x")
        pad = self._r(tk.Frame(body, bg=T["surface"]), "surface")
        pad.pack(fill="x", padx=12, pady=8)
        return pad

    def _file_row(self, parent, var, cmd):
        T = self._T
        row = self._r(tk.Frame(parent, bg=T["surface"]), "surface")
        row.pack(fill="x")

        e = tk.Entry(row, textvariable=var, font=("Segoe UI", 9),
                     state="readonly", relief="flat", bd=0,
                     bg=T["entry_bg"], fg=T["entry_fg"],
                     readonlybackground=T["entry_bg"],
                     highlightbackground=T["border"],
                     highlightthickness=1)
        self._r(e, "entry")
        e.pack(side="left", fill="x", expand=True, padx=(0, 10), ipady=4)

        btn = AnimButton(row, C_ORANGE, C_DARK,
                         text="Обзор", font=("Segoe UI", 9, "bold"),
                         fg="#FFFFFF", command=cmd, padx=14, pady=4)
        btn.pack(side="left")
        return btn

    def _label(self, parent, text, size=9, color="muted", bold=False, **kw):
        T = self._T
        fg = {"muted": T["muted"], "text": T["text"], "text2": T["text2"],
              "orange": C_ORANGE}.get(color, T["muted"])
        font = ("Segoe UI", size, "bold") if bold else ("Segoe UI", size)
        lbl = tk.Label(parent, text=text, font=font,
                       bg=T["bg"], fg=fg, **kw)
        self._r(lbl, "muted_bg" if color == "muted" else "text_bg")
        return lbl

    def _section_lbl(self, parent, text):
        T = self._T
        f = self._r(tk.Frame(parent, bg=T["bg"]), "bg")
        f.pack(fill="x", padx=16, pady=(10, 2))
        tk.Frame(f, bg=C_ORANGE, width=2, height=14).pack(side="left", padx=(0, 8))
        lbl = tk.Label(f, text=text, font=("Segoe UI", 8, "bold"),
                       bg=T["bg"], fg=T["text2"])
        self._r(lbl, "text2_bg")
        lbl.pack(side="left")

    def _log_box(self, parent, height=7):
        T = self._T
        box = scrolledtext.ScrolledText(
            parent, font=("Consolas", 9), height=height,
            state="disabled", bg=T["log_bg"], fg=T["log_fg"],
            relief="flat", selectbackground=C_ORANGE,
            insertbackground=T["log_fg"], padx=8, pady=6)
        box.pack(fill="both", expand=True, padx=16, pady=(0, 4))
        box.tag_config("ok",   foreground=T["tag_ok"])
        box.tag_config("err",  foreground=T["tag_err"])
        box.tag_config("warn", foreground=T["tag_warn"])
        return box

    def _progress_row(self, parent):
        T = self._T
        row = self._r(tk.Frame(parent, bg=T["bg"]), "bg")
        row.pack(fill="x", padx=16, pady=(2, 4))

        style = ttk.Style()
        style.configure("F.Horizontal.TProgressbar",
                        troughcolor=T["border"],
                        background=C_ORANGE, thickness=6)
        bar = ttk.Progressbar(row, style="F.Horizontal.TProgressbar",
                               mode="determinate")
        bar.pack(side="left", fill="x", expand=True, padx=(0, 12), ipady=1)

        lbl = tk.Label(row, text="0%", font=("Segoe UI", 9, "bold"),
                       bg=T["bg"], fg=C_ORANGE, width=5, anchor="e")
        self._r(lbl, "bg")
        lbl.pack(side="left")
        return bar, lbl

    def _run_btn(self, parent, text, cmd, color=C_ORANGE, hover=C_DARK):
        btn = AnimButton(parent, color, hover,
                         text=text, font=("Segoe UI", 10, "bold"),
                         fg="#FFFFFF", command=cmd, padx=28, pady=12)
        btn.pack(pady=(6, 14))
        return btn

    def _write_log(self, box, msg):
        T = self._T
        tag = ("ok"   if "✅" in msg or "Готово" in msg else
               "err"  if "❌" in msg or "ОШИБКА" in msg else
               "warn" if "⚠" in msg else None)
        ts = datetime.now().strftime("%H:%M:%S")
        box.configure(state="normal")
        box.insert("end", f"[{ts}] ", "ts")
        box.tag_config("ts", foreground=T["muted"])
        line = f"{msg}\n"
        box.insert("end", line, tag) if tag else box.insert("end", line)
        box.see("end")
        box.configure(state="disabled")
        self._status_var.set(msg[:90])

    # ── Прогресс ─────────────────────────────────────────────
    STEPS = {
        "загрузк": 1, "качеств": 2, "клиент": 3, "менеджер": 4,
        "отрасл": 5, "сезонн": 6, "rfm": 7, "лояльн": 8,
        "план": 9, "номенклатур": 10, "сохран": 11,
    }
    TOTAL = 11

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
    # ВКЛАДКА 1: АНАЛИЗ
    # ============================================================
    def _build_tab_main(self):
        p = self.tab_main

        self._label(p, "Загрузите файл выгрузки из CRM и запустите анализ",
                    size=10, color="muted").pack(pady=(12, 6), padx=16, anchor="w")

        c1 = self._card(p, "ВХОДНОЙ ФАЙЛ")
        self.input_var = tk.StringVar(value=self._cfg.get("last_input", ""))
        self._file_row(c1, self.input_var, self._browse_input)

        c2 = self._card(p, "ПАПКА ДЛЯ ОТЧЁТА")
        self.output_dir_var = tk.StringVar(value=self._cfg.get("last_output_dir", ""))
        self._file_row(c2, self.output_dir_var, self._browse_output)

        # НОВОЕ: настройки отчёта
        c3 = self._card(p, "НАСТРОЙКИ ОТЧЁТА")
        self._build_settings_card(c3)

        self._section_lbl(p, "ПРОГРЕСС ВЫПОЛНЕНИЯ")
        self.main_bar, self.main_pct = self._progress_row(p)

        self._section_lbl(p, "ЖУРНАЛ")
        self.log_box = self._log_box(p, height=5)

        # НОВОЕ: индикатор расхождения с бухгалтерией
        self._build_recon_indicator(p)

        self.run_btn = self._run_btn(p, "▶   Запустить анализ",
                                      self._start_analysis)

    # ── НОВЫЙ блок: настройки отчёта ─────────────────────────
    def _build_settings_card(self, parent):
        T = self._T
        # Первая строка — переключатель даты
        row1 = self._r(tk.Frame(parent, bg=T["surface"]), "surface")
        row1.pack(fill="x", pady=(0, 4))

        lbl1 = tk.Label(row1, text="Основная дата для итогов:",
                        font=("Segoe UI", 9),
                        bg=T["surface"], fg=T["text"])
        self._r(lbl1, "text_sf")
        lbl1.pack(side="left")

        rb1 = tk.Radiobutton(row1, text="Дата заказа",
                             variable=self._date_by, value="order",
                             bg=T["surface"], fg=T["text"],
                             activebackground=T["surface"],
                             selectcolor=T["surface"],
                             font=("Segoe UI", 9),
                             cursor="hand2")
        self._r(rb1, "text_sf")
        rb1.pack(side="left", padx=(10, 6))

        rb2 = tk.Radiobutton(row1, text="Дата оплаты (бьётся с бухгалтерией)",
                             variable=self._date_by, value="payment",
                             bg=T["surface"], fg=T["text"],
                             activebackground=T["surface"],
                             selectcolor=T["surface"],
                             font=("Segoe UI", 9),
                             cursor="hand2")
        self._r(rb2, "text_sf")
        rb2.pack(side="left")

        # Вторая строка — кнопки редакторов
        row2 = self._r(tk.Frame(parent, bg=T["surface"]), "surface")
        row2.pack(fill="x", pady=(6, 0))

        lbl2 = tk.Label(row2, text="Данные бухгалтерии:",
                        font=("Segoe UI", 9),
                        bg=T["surface"], fg=T["text"])
        self._r(lbl2, "text_sf")
        lbl2.pack(side="left")

        if HAS_VF_EDITOR:
            btn1 = AnimButton(row2, T["surface2"], T["border"],
                              text="⚙ Верифицированные цифры",
                              font=("Segoe UI", 9),
                              fg=T["text"],
                              command=self._open_verified_editor,
                              padx=10, pady=3)
            btn1.pack(side="left", padx=(10, 6))

        if HAS_EXT_EDITOR:
            btn2 = AnimButton(row2, T["surface2"], T["border"],
                              text="⚙ Внешние доходы (программатик)",
                              font=("Segoe UI", 9),
                              fg=T["text"],
                              command=self._open_external_editor,
                              padx=10, pady=3)
            btn2.pack(side="left")

    # ── НОВЫЙ блок: индикатор расхождения ────────────────────
    def _build_recon_indicator(self, parent):
        T = self._T
        f = self._r(tk.Frame(parent, bg=T["bg"]), "bg")
        f.pack(fill="x", padx=16, pady=(4, 4))

        self._recon_label = tk.Label(
            f, text="  Сверка с бухгалтерией появится после запуска анализа",
            font=("Segoe UI", 9),
            bg=T["bg"], fg=T["muted"],
            anchor="w", justify="left",
            padx=10, pady=8)
        self._r(self._recon_label, "muted_bg")
        self._recon_label.pack(fill="x")

    # ── Открытие редакторов ──────────────────────────────────
    def _open_external_editor(self):
        path = os.path.join(APP_DIR, "external_income.json")
        if not os.path.exists(path):
            messagebox.showerror(
                "Нет файла",
                f"Файл external_income.json не найден.\n\n"
                f"Ожидается по пути:\n{path}")
            return
        ExternalIncomeEditor(self, path, theme=self._T)

    def _open_verified_editor(self):
        path = os.path.join(APP_DIR, "verified_figures.json")
        if not os.path.exists(path):
            # Создаём дефолтный
            default = {
                "_описание": "Верифицированные годовые итоги из бухгалтерии.",
                "_год": 2025,
                "total_with_prog": 482404000,
                "total_with_barter_no_prog": 363000166,
                "advertising_no_events": 243005820,
                "programmatic_external": 119403897,
                "other_external_income": 41057102
            }
            try:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(default, f, ensure_ascii=False, indent=2)
            except Exception as e:
                messagebox.showerror("Ошибка",
                    f"Не удалось создать verified_figures.json:\n{e}")
                return
        VerifiedFiguresEditor(self, path, theme=self._T)

    # ============================================================
    # ВКЛАДКА 2: СРАВНЕНИЕ
    # ============================================================
    def _build_tab_compare(self):
        p = self.tab_compare

        self._label(p, "Сравните показатели двух произвольных периодов",
                    size=10, color="muted").pack(pady=(12, 6), padx=16, anchor="w")

        def period_block(title, path_var, label_var):
            c = self._card(p, title)
            self._file_row(c, path_var, lambda v=path_var: self._browse_cmp(v))
            row = self._r(tk.Frame(c, bg=self._T["surface"]), "surface")
            row.pack(fill="x", pady=(6, 0))
            self._r(tk.Label(row, text="Метка периода:",
                             font=("Segoe UI", 9),
                             bg=self._T["surface"],
                             fg=self._T["muted"]), "muted_sf").pack(side="left")
            e = tk.Entry(row, textvariable=label_var,
                         font=("Segoe UI", 9), width=22,
                         relief="flat", bd=0,
                         highlightbackground=self._T["border"],
                         highlightthickness=1,
                         bg=self._T["entry_bg"], fg=self._T["entry_fg"])
            self._r(e, "entry")
            e.pack(side="left", padx=10, ipady=3)

        self.cmp_path_a  = tk.StringVar()
        self.cmp_label_a = tk.StringVar(value="Период А")
        self.cmp_path_b  = tk.StringVar()
        self.cmp_label_b = tk.StringVar(value="Период Б")
        period_block("ПЕРИОД А", self.cmp_path_a, self.cmp_label_a)
        period_block("ПЕРИОД Б", self.cmp_path_b, self.cmp_label_b)

        co = self._card(p, "ПАПКА ДЛЯ ОТЧЁТА")
        self.cmp_output_dir = tk.StringVar(value=self._cfg.get("last_output_dir", ""))
        self._file_row(co, self.cmp_output_dir, self._browse_cmp_output)

        self._section_lbl(p, "ЖУРНАЛ")
        self.cmp_log_box = self._log_box(p, height=5)
        self.cmp_run_btn = self._run_btn(p, "▶   Сравнить периоды",
                                          self._start_comparison)

    # ============================================================
    # ВКЛАДКА 3: ПЛАН
    # ============================================================
    def _build_tab_plan(self):
        p = self.tab_plan
        T = self._T

        self._label(p, "Плановые суммы по менеджерам (тыс. руб.)",
                    size=10, color="muted").pack(pady=(12, 2), padx=16, anchor="w")
        self._label(p, "Планы сохраняются автоматически и применяются при каждом запуске анализа",
                    size=8, color="muted").pack(padx=16, anchor="w")

        outer = self._r(tk.Frame(p, bg=T["bg"]), "bg")
        outer.pack(fill="both", expand=True, padx=16, pady=8)

        canvas = tk.Canvas(outer, bg=T["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self._plan_inner = self._r(tk.Frame(canvas, bg=T["surface"]), "surface")
        self._plan_inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._plan_inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        hdr = tk.Frame(self._plan_inner, bg=C_ORANGE)
        hdr.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Label(hdr, text="  Менеджер",
                 font=("Segoe UI", 9, "bold"),
                 bg=C_ORANGE, fg="#FFFFFF",
                 width=30, anchor="w", pady=7).pack(side="left")
        tk.Label(hdr, text="План, тыс. руб.",
                 font=("Segoe UI", 9, "bold"),
                 bg=C_ORANGE, fg="#FFFFFF",
                 width=16, pady=7).pack(side="left")

        saved = self._cfg.get("manager_plan", {})
        managers = list(dict.fromkeys([
            "Петрова Э.", "Егорова И.", "Демахина Е.",
            "Дубровская Т.", "Долгополова Н.", "Демидова Г.",
            "Сосновик А.", "Киселёва Е.", "Лея К.",
            "Борисенко Е.", "Сорокина Т.", "Шевчук Е.", "Карпус О.",
        ] + list(saved.keys())))
        for i, name in enumerate(managers, 1):
            self._add_plan_row(name, str(saved.get(name, "")), i)

        add = self._r(tk.Frame(p, bg=T["bg"]), "bg")
        add.pack(fill="x", padx=16, pady=(4, 8))
        self._r(tk.Label(add, text="Добавить менеджера:",
                         font=("Segoe UI", 9),
                         bg=T["bg"], fg=T["text2"]), "text2_bg").pack(side="left")
        self._new_mgr = tk.StringVar()
        ne = tk.Entry(add, textvariable=self._new_mgr,
                      font=("Segoe UI", 9), width=22,
                      relief="flat", bd=0,
                      highlightbackground=T["border"],
                      highlightthickness=1,
                      bg=T["entry_bg"], fg=T["entry_fg"])
        self._r(ne, "entry")
        ne.pack(side="left", padx=10, ipady=3)
        AnimButton(add, C_ORANGE, C_DARK,
                   text="+ Добавить", font=("Segoe UI", 9, "bold"),
                   fg="#FFFFFF", command=self._add_manager_ui,
                   padx=12, pady=4).pack(side="left")

    def _add_plan_row(self, name, val="", idx=None):
        T = self._T
        if idx is None:
            idx = len(self._plan_vars) + 1
        even = idx % 2 == 0
        bg = T["surface"] if even else T["plan_alt"]

        lbl = tk.Label(self._plan_inner, text=f"  {name}",
                       font=("Segoe UI", 9), bg=bg,
                       fg=T["text"], anchor="w", pady=5)
        self._r(lbl, "plan_le" if even else "plan_lo")
        lbl.grid(row=idx, column=0, sticky="ew")

        var = tk.StringVar(value=val)
        var.trace_add("write", lambda *_, n=name, v=var: self._save_plan(n, v))
        self._plan_vars[name] = var

        ent = tk.Entry(self._plan_inner, textvariable=var,
                       font=("Segoe UI", 9), width=14,
                       relief="flat", bd=0,
                       highlightbackground=T["border"],
                       highlightthickness=1,
                       bg=bg, fg=T["text"],
                       insertbackground=T["text"])
        self._r(ent, "plan_e" if even else "plan_o")
        ent.grid(row=idx, column=1, pady=2, padx=8)

    def _add_manager_ui(self):
        name = self._new_mgr.get().strip()
        if not name:
            return
        if name in self._plan_vars:
            messagebox.showinfo("", f"«{name}» уже есть в списке.")
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
    # ВКЛАДКА 4: НАБЛЮДАТЕЛЬ
    # ============================================================
    def _build_tab_watch(self):
        p = self.tab_watch
        T = self._T

        self._label(p,
                    "Автоматически запускает анализ через 5 минут после появления нового .xlsx",
                    size=10, color="muted").pack(pady=(12, 6), padx=16, anchor="w")

        cw = self._card(p, "ПАПКА ДЛЯ НАБЛЮДЕНИЯ")
        self.watch_dir_var = tk.StringVar(value=self._cfg.get("watch_dir", ""))
        self._file_row(cw, self.watch_dir_var, self._browse_watch_dir)

        co = self._card(p, "ПАПКА ДЛЯ ОТЧЁТОВ")
        self.watch_out_var = tk.StringVar(value=self._cfg.get("watch_output_dir", ""))
        self._file_row(co, self.watch_out_var, self._browse_watch_out)

        sf = self._r(tk.Frame(p, bg=T["bg"]), "bg")
        sf.pack(fill="x", padx=16, pady=(12, 4))

        self._status_dot = tk.Label(sf, text="●", font=("Segoe UI", 14),
                                     bg=T["bg"], fg=T["muted"])
        self._r(self._status_dot, "muted_bg")
        self._status_dot.pack(side="left", padx=(0, 8))

        self.watch_status_var = tk.StringVar(value="Остановлен")
        self.watch_status_lbl = tk.Label(
            sf, textvariable=self.watch_status_var,
            font=("Segoe UI", 11, "bold"),
            bg=T["bg"], fg=T["muted"])
        self._r(self.watch_status_lbl, "muted_bg")
        self.watch_status_lbl.pack(side="left")

        self._section_lbl(p, "ЖУРНАЛ")
        self.watch_log_box = self._log_box(p, height=7)

        bf = self._r(tk.Frame(p, bg=T["bg"]), "bg")
        bf.pack(pady=(4, 14))

        self.watch_start_btn = AnimButton(
            bf, C_GREEN, C_GREEN_D,
            text="▶   Запустить наблюдатель",
            font=("Segoe UI", 10, "bold"),
            fg="#FFFFFF", command=self._start_watcher,
            padx=18, pady=10)
        self.watch_start_btn.pack(side="left", padx=(0, 10))

        self.watch_stop_btn = AnimButton(
            bf, C_RED, C_RED_D,
            text="⏹   Остановить",
            font=("Segoe UI", 10, "bold"),
            fg="#FFFFFF", command=self._stop_watcher,
            padx=18, pady=10, state="disabled")
        self.watch_stop_btn.pack(side="left")

    # ── Browse ───────────────────────────────────────────────
    def _browse_input(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            self.input_var.set(p)
            self._cfg["last_input"] = p
            if not self.output_dir_var.get():
                d = os.path.dirname(p)
                self.output_dir_var.set(d)
                self._cfg["last_output_dir"] = d
            save_config(self._cfg)

    def _browse_output(self):
        p = filedialog.askdirectory()
        if p:
            self.output_dir_var.set(p)
            self._cfg["last_output_dir"] = p
            save_config(self._cfg)

    def _browse_cmp(self, var):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p:
            var.set(p)
            if not self.cmp_output_dir.get():
                self.cmp_output_dir.set(os.path.dirname(p))

    def _browse_cmp_output(self):
        p = filedialog.askdirectory()
        if p:
            self.cmp_output_dir.set(p)

    def _browse_watch_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.watch_dir_var.set(p)
            self._cfg["watch_dir"] = p
            if not self.watch_out_var.get():
                self.watch_out_var.set(p)
                self._cfg["watch_output_dir"] = p
            save_config(self._cfg)

    def _browse_watch_out(self):
        p = filedialog.askdirectory()
        if p:
            self.watch_out_var.set(p)
            self._cfg["watch_output_dir"] = p
            save_config(self._cfg)

    # ── Логи ─────────────────────────────────────────────────
    def _log(self, msg):
        self._write_log(self.log_box, msg)

    def _cmp_log(self, msg):
        self._write_log(self.cmp_log_box, msg)

    def _watch_log(self, msg):
        self.after(0, lambda: self._write_log(self.watch_log_box, msg))

    # ── Анализ ───────────────────────────────────────────────
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
        self.run_btn.configure(state="disabled", text="  Выполняется...")
        self._start_anim()
        # Сбрасываем индикатор
        self._recon_label.config(
            text="  Считаю...",
            fg=self._T["muted"])
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

            # Обработка нового (dict) и старого (str) формата
            if isinstance(result, dict):
                path = result.get("output_path", out_path)
                self.after(0, lambda r=result: self._update_recon(r))
            else:
                path = result or out_path

            self.after(0, lambda: self._on_success(path))
        except Exception as e:
            log_fn(f"❌ ОШИБКА: {e}")
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self._stop_anim()
            self.after(0, lambda: self.run_btn.configure(
                state="normal", text="▶   Запустить анализ"))

    def _update_recon(self, result):
        """Обновляет индикатор расхождения с бухгалтерией."""
        pct = result.get("deviation_pct")
        grand = result.get("grand_total", 0)
        verified = result.get("verified_total", 0)
        date_by = result.get("date_by", "order")

        if pct is None or verified <= 0:
            self._recon_label.config(
                text=f"  CRM+внешние: {grand:,.0f} тыс. руб.".replace(",", " "),
                fg=self._T["text"])
            return

        if abs(pct) < 2:
            color = C_GREEN
            icon = "✅"
            verdict = "отлично"
        elif abs(pct) < 5:
            color = C_AMBER
            icon = "⚠"
            verdict = "в норме"
        else:
            color = C_RED
            icon = "❌"
            verdict = "большое расхождение"

        date_label = "дата оплаты" if date_by == "payment" else "дата заказа"

        text = (f"  {icon}  {date_label.capitalize()}: "
                f"CRM+внешние {grand:,.0f} тыс. │ "
                f"Цель: {verified:,.0f} тыс. │ "
                f"{pct:+.2f}% — {verdict}").replace(",", " ")

        self._recon_label.config(text=text, fg=color)

    # ── Сравнение ────────────────────────────────────────────
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
        self.cmp_run_btn.configure(state="disabled", text="  Выполняется...")
        self._start_anim()
        threading.Thread(target=self._thread_comparison,
                         args=(pa, pb, out_path, la, lb),
                         daemon=True).start()

    def _thread_comparison(self, pa, pb, out_path, la, lb):
        try:
            self._cmp_log(f"Сравниваю «{la}» и «{lb}»")
            run_comparison(pa, pb, out_path, la, lb, log=self._cmp_log)
            self._cmp_log(f"✅ Готово: {out_path}")
            self.after(0, lambda: self._on_success(out_path))
        except Exception as e:
            self._cmp_log(f"❌ ОШИБКА: {e}")
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self._stop_anim()
            self.after(0, lambda: self.cmp_run_btn.configure(
                state="normal", text="▶   Сравнить периоды"))

    # ── Наблюдатель ──────────────────────────────────────────
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
    # ВКЛАДКА 5: ДАШБОРД
    # ============================================================
    def _build_tab_dash(self):
        p = self.tab_dash
        T = self._T

        self._label(p,
                    "Загрузите файл и откройте интерактивный дашборд в браузере",
                    size=10, color="muted").pack(pady=(12, 6), padx=16, anchor="w")
        self._label(p,
                    "Дашборд включает: KPI, выручка по месяцам, топ клиентов и менеджеров, отрасли",
                    size=8, color="muted").pack(padx=16, anchor="w", pady=(0, 8))

        c1 = self._card(p, "ВХОДНОЙ ФАЙЛ")
        self.dash_input_var = tk.StringVar(value=self._cfg.get("last_input", ""))
        self._file_row(c1, self.dash_input_var, self._browse_dash_input)

        preview = self._r(tk.Frame(p, bg=T["bg"]), "bg")
        preview.pack(fill="x", padx=16, pady=(8, 4))

        items = [
            ("📈", "Выручка по месяцам", "Линейный график с градиентом"),
            ("🏆", "Топ-10 клиентов",    "Горизонтальный столбчатый"),
            ("👤", "Топ-10 менеджеров",  "Вертикальный столбчатый"),
            ("🍩", "Отрасли",             "Пончик (doughnut)"),
        ]

        row = self._r(tk.Frame(preview, bg=T["bg"]), "bg")
        row.pack(fill="x")
        for icon, title, desc in items:
            card = self._r(tk.Frame(row, bg=T["surface"],
                                    highlightbackground=T["border"],
                                    highlightthickness=1), "surface")
            card.pack(side="left", fill="x", expand=True, padx=(0, 8), pady=4, ipady=10, ipadx=8)
            tk.Label(card, text=icon, font=("Segoe UI", 22),
                     bg=T["surface"]).pack()
            tk.Label(card, text=title, font=("Segoe UI", 9, "bold"),
                     bg=T["surface"], fg=T["text"]).pack()
            tk.Label(card, text=desc, font=("Segoe UI", 8),
                     bg=T["surface"], fg=T["muted"]).pack()

        self._section_lbl(p, "ЖУРНАЛ")
        self.dash_log_box = self._log_box(p, height=5)

        self.dash_btn = self._run_btn(
            p, "🌐   Открыть дашборд в браузере",
            self._start_dashboard,
            color="#1D4ED8", hover="#1E40AF")

    def _browse_dash_input(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.dash_input_var.set(path)
            self._cfg["last_input"] = path
            self.input_var.set(path)
            save_config(self._cfg)

    def _dash_log(self, msg):
        self._write_log(self.dash_log_box, msg)

    def _start_dashboard(self):
        inp = self.dash_input_var.get().strip()
        if not inp:
            messagebox.showwarning("", "Выберите входной файл.")
            return
        self.dash_btn.configure(state="disabled", text="⏳ Генерирую...")
        self._start_anim()
        threading.Thread(
            target=self._thread_dashboard,
            args=(inp,),
            daemon=True).start()

    def _thread_dashboard(self, inp):
        try:
            generate_dashboard(inp, log=self._dash_log)
        except Exception as e:
            self._dash_log(f"❌ ОШИБКА: {e}")
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self._stop_anim()
            self.after(0, lambda: self.dash_btn.configure(
                state="normal", text="🌐   Открыть дашборд в браузере"))

    def _on_success(self, path):
        self._status_var.set(f"Готово: {os.path.basename(path)}")
        if messagebox.askyesno("Готово!", f"Отчёт создан:\n{path}\n\nОткрыть папку?"):
            os.startfile(os.path.dirname(path))


if __name__ == "__main__":
    app = App()
    app.mainloop()
