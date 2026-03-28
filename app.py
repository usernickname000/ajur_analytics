# ============================================================
# app.py — графический интерфейс утилиты аналитики заказов
# Запуск: python app.py
# ============================================================

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import json
import os
from datetime import datetime

from analytics import run_analytics
from comparison import run_comparison
from watcher import FolderWatcher

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
ICON_PATH   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")

# ── Палитра ──────────────────────────────────────────────────
C_ORANGE      = "#F38120"
C_ORANGE_DARK = "#D06A10"
C_ORANGE_LITE = "#FEF0E3"
C_BG          = "#F5F5F5"
C_WHITE       = "#FFFFFF"
C_BORDER      = "#E0D8D0"
C_TEXT        = "#1A1A1A"
C_MUTED       = "#777777"
C_LOG_BG      = "#1C1C1E"
C_LOG_FG      = "#F0EDE8"
C_BAR1        = "#F38120"
C_BAR2        = "#4A90D9"
C_BAR3        = "#2ECC71"
C_SUCCESS     = "#27AE60"
C_ERROR       = "#E74C3C"


# ============================================================
# КОНФИГ
# ============================================================

def load_config() -> dict:
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}

def save_config(cfg: dict):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ============================================================
# КАСТОМНЫЕ ВИДЖЕТЫ
# ============================================================

class HoverButton(tk.Button):
    """Кнопка с hover-эффектом."""
    def __init__(self, master, bg_normal, bg_hover, **kwargs):
        super().__init__(master, bg=bg_normal, activebackground=bg_hover,
                         relief="flat", cursor="hand2", bd=0, **kwargs)
        self._bg_normal = bg_normal
        self._bg_hover  = bg_hover
        self.bind("<Enter>", lambda e: self.config(bg=bg_hover))
        self.bind("<Leave>", lambda e: self.config(bg=bg_normal))


class MiniBarChart(tk.Canvas):
    """Мини-столбчатый график для шапки."""
    def __init__(self, master, values, colors, labels, width=120, height=48, **kw):
        super().__init__(master, width=width, height=height,
                         bg=C_ORANGE, highlightthickness=0, **kw)
        self._values = values
        self._colors = colors
        self._labels = labels
        self._w = width
        self._h = height
        self._draw()

    def _draw(self):
        self.delete("all")
        if not self._values:
            return
        max_v = max(self._values) or 1
        n = len(self._values)
        gap = 4
        bar_w = (self._w - gap * (n + 1)) // n
        for i, (v, c) in enumerate(zip(self._values, self._colors)):
            bh = int((v / max_v) * (self._h - 14))
            x0 = gap + i * (bar_w + gap)
            y0 = self._h - 12 - bh
            x1 = x0 + bar_w
            y1 = self._h - 12
            self.create_rectangle(x0, y0, x1, y1, fill=c, outline="", width=0)
        # Подписи
        for i, lbl in enumerate(self._labels):
            x = gap + i * (bar_w + gap) + bar_w // 2
            self.create_text(x, self._h - 6, text=lbl,
                             fill="white", font=("Segoe UI", 6), anchor="center")

    def update_values(self, values):
        self._values = values
        self._draw()


# ============================================================
# ГЛАВНОЕ ОКНО
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Аналитика заказов — Фонтанка.ру")
        self.geometry("740x640")
        self.resizable(False, False)
        self.configure(bg=C_BG)

        # Иконка
        if os.path.exists(ICON_PATH):
            try:
                self.iconbitmap(ICON_PATH)
            except Exception:
                pass

        self._cfg = load_config()
        self._watcher: FolderWatcher | None = None
        self._plan_vars: dict[str, tk.StringVar] = {}
        self._anim_running = False
        self._anim_step = 0

        self._build_header()
        self._build_tabs()
        self._build_statusbar()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        if self._watcher:
            self._watcher.stop()
        save_config(self._cfg)
        self.destroy()

    # ============================================================
    # ШАПКА В СТИЛЕ ФОНТАНКИ
    # ============================================================

    def _build_header(self):
        header = tk.Frame(self, bg=C_ORANGE, height=72)
        header.pack(fill="x")
        header.pack_propagate(False)

        # Левая часть — логотип + название
        left = tk.Frame(header, bg=C_ORANGE)
        left.pack(side="left", padx=16, pady=0, fill="y")

        # Птица-символ (canvas)
        bird = tk.Canvas(left, width=36, height=36, bg=C_ORANGE,
                         highlightthickness=0)
        bird.pack(side="left", padx=(0, 8), pady=18)
        self._draw_bird(bird)

        title_frame = tk.Frame(left, bg=C_ORANGE)
        title_frame.pack(side="left", fill="y", pady=12)

        tk.Label(title_frame, text="фонтанка.ру",
                 font=("Georgia", 18, "bold italic"),
                 bg=C_ORANGE, fg=C_WHITE).pack(anchor="w")
        tk.Label(title_frame, text="АНАЛИТИКА КОММЕРЧЕСКОГО ОТДЕЛА",
                 font=("Segoe UI", 7, "bold"),
                 bg=C_ORANGE, fg="#FFD8B0",
                 letterSpacing=2).pack(anchor="w")

        # Разделитель
        tk.Frame(header, bg="#D06A10", width=1).pack(side="left", fill="y",
                                                      padx=12, pady=8)

        # Правая часть — мини-графики
        charts_frame = tk.Frame(header, bg=C_ORANGE)
        charts_frame.pack(side="left", fill="y", pady=10)

        tk.Label(charts_frame, text="ПОСЛЕДНИЙ ОТЧЁТ",
                 font=("Segoe UI", 7, "bold"),
                 bg=C_ORANGE, fg="#FFD8B0").pack(anchor="w")

        self._header_charts = tk.Frame(charts_frame, bg=C_ORANGE)
        self._header_charts.pack(fill="x", pady=(2, 0))

        # Три мини-графика
        self._chart_mgr  = self._make_mini_chart(self._header_charts,
                                                   "Менеджеры", [40,70,55,80,60], C_WHITE)
        self._chart_cli  = self._make_mini_chart(self._header_charts,
                                                   "Клиенты", [30,50,90,45,70], "#FFD8B0")
        self._chart_rev  = self._make_mini_chart(self._header_charts,
                                                   "Выручка", [60,80,40,95,50], "#FFC080")

        # Дата справа
        right = tk.Frame(header, bg=C_ORANGE)
        right.pack(side="right", padx=16, fill="y")
        now = datetime.now()
        tk.Label(right, text=now.strftime("%d.%m.%Y"),
                 font=("Segoe UI", 11, "bold"),
                 bg=C_ORANGE, fg=C_WHITE).pack(anchor="e", pady=(18, 0))
        tk.Label(right, text=now.strftime("%H:%M"),
                 font=("Segoe UI", 9),
                 bg=C_ORANGE, fg="#FFD8B0").pack(anchor="e")

    def _draw_bird(self, canvas):
        """Рисуем упрощённую птицу-символ Фонтанки."""
        # Дуга-основа
        canvas.create_arc(3, 5, 33, 33, start=40, extent=260,
                          outline="white", width=2, style="arc")
        # Тело птицы
        canvas.create_oval(12, 8, 22, 16, fill="white", outline="")
        # Крыло
        canvas.create_line(17, 12, 10, 22, fill="white", width=2)
        canvas.create_line(17, 12, 26, 20, fill="white", width=2)
        # Хвост
        canvas.create_line(12, 15, 5, 24, fill="white", width=2)
        # Точка
        canvas.create_oval(15, 3, 20, 8, fill="white", outline="")

    def _make_mini_chart(self, parent, label, values, color):
        frame = tk.Frame(parent, bg=C_ORANGE)
        frame.pack(side="left", padx=6)
        tk.Label(frame, text=label, font=("Segoe UI", 7),
                 bg=C_ORANGE, fg="#FFD8B0").pack()
        chart = MiniBarChart(frame, values,
                              [color] * len(values), [""] * len(values),
                              width=80, height=32)
        chart.pack()
        return chart

    # ============================================================
    # ВКЛАДКИ
    # ============================================================

    def _build_tabs(self):
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Fontanka.TNotebook",
                         background=C_BG, borderwidth=0,
                         tabmargins=[0, 0, 0, 0])
        style.configure("Fontanka.TNotebook.Tab",
                         font=("Segoe UI", 10),
                         padding=[16, 7],
                         background=C_BORDER,
                         foreground=C_MUTED)
        style.map("Fontanka.TNotebook.Tab",
                  background=[("selected", C_WHITE)],
                  foreground=[("selected", C_ORANGE)],
                  font=[("selected", ("Segoe UI", 10, "bold"))])

        self.notebook = ttk.Notebook(self, style="Fontanka.TNotebook")
        self.notebook.pack(fill="both", expand=True, padx=0, pady=0)

        self.tab_main    = tk.Frame(self.notebook, bg=C_BG)
        self.tab_compare = tk.Frame(self.notebook, bg=C_BG)
        self.tab_plan    = tk.Frame(self.notebook, bg=C_BG)
        self.tab_watch   = tk.Frame(self.notebook, bg=C_BG)

        self.notebook.add(self.tab_main,    text="  Анализ  ")
        self.notebook.add(self.tab_compare, text="  Сравнение  ")
        self.notebook.add(self.tab_plan,    text="  План  ")
        self.notebook.add(self.tab_watch,   text="  Наблюдатель  ")

        self._build_tab_main()
        self._build_tab_compare()
        self._build_tab_plan()
        self._build_tab_watch()

    # ============================================================
    # СТАТУС-БАР
    # ============================================================

    def _build_statusbar(self):
        sb = tk.Frame(self, bg=C_BORDER, height=26)
        sb.pack(fill="x", side="bottom")
        sb.pack_propagate(False)

        self._status_var = tk.StringVar(value="Готов к работе")
        tk.Label(sb, textvariable=self._status_var,
                 font=("Segoe UI", 9), bg=C_BORDER,
                 fg=C_MUTED, anchor="w").pack(side="left", padx=12, pady=4)

        self._anim_label = tk.Label(sb, text="",
                                     font=("Segoe UI", 9),
                                     bg=C_BORDER, fg=C_ORANGE)
        self._anim_label.pack(side="right", padx=12, pady=4)

        tk.Label(sb, text=f"v1.5  |  Фонтанка.ру",
                 font=("Segoe UI", 8), bg=C_BORDER,
                 fg=C_MUTED).pack(side="right", padx=12)

    def _set_status(self, text, color=None):
        self._status_var.set(text)
        # цвет статуса меняем через label поиском
        for w in self.winfo_children():
            pass  # статус-бар сам обновится

    def _start_anim(self):
        self._anim_running = True
        self._anim_step = 0
        self._tick_anim()

    def _tick_anim(self):
        if not self._anim_running:
            self._anim_label.config(text="")
            return
        frames = ["▪▫▫", "▫▪▫", "▫▫▪", "▫▪▫"]
        self._anim_label.config(text=frames[self._anim_step % len(frames)],
                                 fg=C_ORANGE)
        self._anim_step += 1
        self.after(300, self._tick_anim)

    def _stop_anim(self):
        self._anim_running = False

    # ============================================================
    # ВСПОМОГАТЕЛЬНЫЕ БИЛДЕРЫ
    # ============================================================

    def _section(self, parent, title):
        """LabelFrame в стиле Фонтанки."""
        f = tk.Frame(parent, bg=C_WHITE, relief="flat",
                     highlightbackground=C_BORDER, highlightthickness=1)
        f.pack(fill="x", padx=14, pady=4)
        tk.Label(f, text=f"  {title}",
                 font=("Segoe UI", 8, "bold"),
                 bg=C_ORANGE, fg=C_WHITE,
                 anchor="w", height=2).pack(fill="x")
        inner = tk.Frame(f, bg=C_WHITE)
        inner.pack(fill="x", padx=10, pady=6)
        return inner

    def _file_row(self, parent, var, browse_cmd):
        """Строка с полем пути и кнопкой Обзор."""
        row = tk.Frame(parent, bg=C_WHITE)
        row.pack(fill="x")
        tk.Entry(row, textvariable=var,
                 font=("Segoe UI", 9), state="readonly",
                 relief="solid", bd=1,
                 bg="#FAFAFA", fg=C_TEXT,
                 highlightbackground=C_BORDER).pack(side="left", fill="x",
                                                     expand=True, padx=(0, 8))
        HoverButton(row, bg_normal=C_ORANGE, bg_hover=C_ORANGE_DARK,
                    text="Обзор...", font=("Segoe UI", 9),
                    fg=C_WHITE, command=browse_cmd,
                    padx=12, pady=4).pack(side="left")

    def _log_widget(self, parent, height=7):
        box = scrolledtext.ScrolledText(
            parent, font=("Consolas", 9), height=height,
            state="disabled", bg=C_LOG_BG, fg=C_LOG_FG,
            insertbackground="white", relief="flat",
            selectbackground=C_ORANGE)
        box.pack(fill="both", expand=True, padx=14, pady=(0, 4))
        return box

    def _run_button(self, parent, text, command, color=None):
        c = color or C_ORANGE
        dark = C_ORANGE_DARK if c == C_ORANGE else "#1A7A40"
        btn = HoverButton(parent, bg_normal=c, bg_hover=dark,
                          text=text,
                          font=("Segoe UI", 11, "bold"),
                          fg=C_WHITE, command=command,
                          padx=24, pady=11)
        btn.pack(pady=(6, 12))
        return btn

    def _progress_row(self, parent):
        """Прогресс-бар + процент."""
        row = tk.Frame(parent, bg=C_BG)
        row.pack(fill="x", padx=14, pady=(2, 0))

        style = ttk.Style()
        style.configure("Fontanka.Horizontal.TProgressbar",
                         troughcolor=C_BORDER,
                         background=C_ORANGE,
                         thickness=8)
        bar = ttk.Progressbar(row, style="Fontanka.Horizontal.TProgressbar",
                               mode="determinate", length=580)
        bar.pack(side="left", fill="x", expand=True, padx=(0, 10))
        pct_lbl = tk.Label(row, text="0%", font=("Segoe UI", 9, "bold"),
                            bg=C_BG, fg=C_ORANGE, width=4)
        pct_lbl.pack(side="left")
        return bar, pct_lbl

    # ============================================================
    # ВКЛАДКА 1: АНАЛИЗ
    # ============================================================

    def _build_tab_main(self):
        p = self.tab_main

        tk.Label(p, text="Выберите файл с заказами и нажмите «Запустить»",
                 font=("Segoe UI", 10), bg=C_BG, fg=C_MUTED).pack(pady=(10, 4))

        sec_in = self._section(p, "ВХОДНОЙ ФАЙЛ")
        self.input_var = tk.StringVar(value=self._cfg.get("last_input", ""))
        self._file_row(sec_in, self.input_var, self._browse_input)

        sec_out = self._section(p, "ПАПКА ДЛЯ ОТЧЁТА")
        self.output_dir_var = tk.StringVar(value=self._cfg.get("last_output_dir", ""))
        self._file_row(sec_out, self.output_dir_var, self._browse_output)

        # Прогресс
        tk.Label(p, text="ПРОГРЕСС",
                 font=("Segoe UI", 8, "bold"),
                 bg=C_BG, fg=C_MUTED).pack(anchor="w", padx=14, pady=(8, 2))
        self.main_bar, self.main_pct = self._progress_row(p)

        # Лог
        tk.Label(p, text="ЖУРНАЛ ВЫПОЛНЕНИЯ",
                 font=("Segoe UI", 8, "bold"),
                 bg=C_BG, fg=C_MUTED).pack(anchor="w", padx=14, pady=(6, 2))
        self.log_box = self._log_widget(p, height=7)

        self.run_btn = self._run_button(p, "▶  Запустить анализ",
                                         self._start_analysis)

    # ============================================================
    # ВКЛАДКА 2: СРАВНЕНИЕ
    # ============================================================

    def _build_tab_compare(self):
        p = self.tab_compare

        tk.Label(p, text="Выберите два файла и задайте метки периодов",
                 font=("Segoe UI", 10), bg=C_BG, fg=C_MUTED).pack(pady=(10, 4))

        sec_a = self._section(p, "ПЕРИОД А")
        self.cmp_path_a = tk.StringVar()
        self._file_row(sec_a, self.cmp_path_a,
                       lambda: self._browse_cmp(self.cmp_path_a))
        lf_a = tk.Frame(sec_a, bg=C_WHITE)
        lf_a.pack(fill="x", pady=(4, 0))
        tk.Label(lf_a, text="Название:", font=("Segoe UI", 9),
                 bg=C_WHITE, fg=C_MUTED).pack(side="left")
        self.cmp_label_a = tk.StringVar(value="Период А")
        tk.Entry(lf_a, textvariable=self.cmp_label_a,
                 font=("Segoe UI", 9), width=22,
                 relief="solid", bd=1).pack(side="left", padx=8)

        sec_b = self._section(p, "ПЕРИОД Б")
        self.cmp_path_b = tk.StringVar()
        self._file_row(sec_b, self.cmp_path_b,
                       lambda: self._browse_cmp(self.cmp_path_b))
        lf_b = tk.Frame(sec_b, bg=C_WHITE)
        lf_b.pack(fill="x", pady=(4, 0))
        tk.Label(lf_b, text="Название:", font=("Segoe UI", 9),
                 bg=C_WHITE, fg=C_MUTED).pack(side="left")
        self.cmp_label_b = tk.StringVar(value="Период Б")
        tk.Entry(lf_b, textvariable=self.cmp_label_b,
                 font=("Segoe UI", 9), width=22,
                 relief="solid", bd=1).pack(side="left", padx=8)

        sec_out = self._section(p, "ПАПКА ДЛЯ ОТЧЁТА")
        self.cmp_output_dir = tk.StringVar(value=self._cfg.get("last_output_dir", ""))
        self._file_row(sec_out, self.cmp_output_dir, self._browse_cmp_output)

        tk.Label(p, text="ЖУРНАЛ ВЫПОЛНЕНИЯ",
                 font=("Segoe UI", 8, "bold"),
                 bg=C_BG, fg=C_MUTED).pack(anchor="w", padx=14, pady=(8, 2))
        self.cmp_log_box = self._log_widget(p, height=6)
        self.cmp_run_btn = self._run_button(p, "▶  Сравнить периоды",
                                             self._start_comparison)

    # ============================================================
    # ВКЛАДКА 3: ПЛАН
    # ============================================================

    def _build_tab_plan(self):
        p = self.tab_plan

        tk.Label(p, text="Плановые суммы по менеджерам (тыс. руб.)",
                 font=("Segoe UI", 10), bg=C_BG, fg=C_MUTED).pack(pady=(10, 2))
        tk.Label(p, text="Сохраняется автоматически. Применяется при каждом запуске анализа.",
                 font=("Segoe UI", 8), bg=C_BG, fg=C_MUTED).pack(pady=(0, 6))

        outer = tk.Frame(p, bg=C_BG)
        outer.pack(fill="both", expand=True, padx=14)

        canvas = tk.Canvas(outer, bg=C_BG, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self._plan_inner = tk.Frame(canvas, bg=C_WHITE,
                                     highlightbackground=C_BORDER,
                                     highlightthickness=1)
        self._plan_inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._plan_inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # Заголовок таблицы
        hdr = tk.Frame(self._plan_inner, bg=C_ORANGE)
        hdr.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Label(hdr, text="  Менеджер", font=("Segoe UI", 9, "bold"),
                 bg=C_ORANGE, fg=C_WHITE, width=28, anchor="w",
                 pady=6).pack(side="left")
        tk.Label(hdr, text="План, тыс. руб.", font=("Segoe UI", 9, "bold"),
                 bg=C_ORANGE, fg=C_WHITE, width=16,
                 pady=6).pack(side="left")

        saved_plans = self._cfg.get("manager_plan", {})
        default_managers = [
            "Петрова Э.", "Егорова И.", "Демахина Е.",
            "Дубровская Т.", "Долгополова Н.", "Демидова Г.",
            "Сосновик А.", "Киселёва Е.", "Лея К.",
            "Борисенко Е.", "Сорокина Т.", "Шевчук Е.", "Карпус О.",
        ]
        all_managers = list(dict.fromkeys(
            default_managers + list(saved_plans.keys())))

        for i, name in enumerate(all_managers, start=1):
            self._add_plan_row(name, str(saved_plans.get(name, "")), i)

        add_frame = tk.Frame(p, bg=C_BG)
        add_frame.pack(fill="x", padx=14, pady=(6, 0))
        tk.Label(add_frame, text="Добавить менеджера:",
                 font=("Segoe UI", 9), bg=C_BG, fg=C_TEXT).pack(side="left")
        self._new_manager_var = tk.StringVar()
        tk.Entry(add_frame, textvariable=self._new_manager_var,
                 font=("Segoe UI", 9), width=22,
                 relief="solid", bd=1).pack(side="left", padx=8)
        HoverButton(add_frame, bg_normal=C_ORANGE, bg_hover=C_ORANGE_DARK,
                    text="Добавить", font=("Segoe UI", 9),
                    fg=C_WHITE, command=self._add_manager_row_ui,
                    padx=10, pady=4).pack(side="left")

    def _add_plan_row(self, name: str, saved_val: str = "", row_idx: int = None):
        if row_idx is None:
            row_idx = len(self._plan_vars) + 1
        bg = C_WHITE if row_idx % 2 == 0 else C_ORANGE_LITE
        tk.Label(self._plan_inner, text=f"  {name}",
                 font=("Segoe UI", 9), bg=bg, anchor="w",
                 pady=5).grid(row=row_idx, column=0, sticky="ew", padx=(0, 1))
        var = tk.StringVar(value=saved_val)
        var.trace_add("write", lambda *_, n=name, v=var: self._save_plan_entry(n, v))
        self._plan_vars[name] = var
        e = tk.Entry(self._plan_inner, textvariable=var,
                     font=("Segoe UI", 9), width=16,
                     relief="solid", bd=1, bg=bg)
        e.grid(row=row_idx, column=1, pady=1, padx=4)

    def _add_manager_row_ui(self):
        name = self._new_manager_var.get().strip()
        if not name:
            return
        if name in self._plan_vars:
            messagebox.showinfo("", f"Менеджер «{name}» уже есть в списке.")
            return
        self._add_plan_row(name)
        self._new_manager_var.set("")

    def _save_plan_entry(self, manager_name: str, var: tk.StringVar):
        plans = self._cfg.setdefault("manager_plan", {})
        val = var.get().strip()
        if val:
            plans[manager_name] = val
        else:
            plans.pop(manager_name, None)
        save_config(self._cfg)

    def _get_manager_plan_rub(self) -> dict | None:
        result = {}
        for name, var in self._plan_vars.items():
            val = var.get().strip().replace(",", ".")
            if val:
                try:
                    result[name] = float(val) * 1000
                except ValueError:
                    pass
        return result if result else None

    # ============================================================
    # ВКЛАДКА 4: НАБЛЮДАТЕЛЬ
    # ============================================================

    def _build_tab_watch(self):
        p = self.tab_watch

        tk.Label(p,
                 text="Автоматически запускает анализ через 5 минут\n"
                      "после появления нового .xlsx файла в папке",
                 font=("Segoe UI", 10), bg=C_BG, fg=C_MUTED,
                 justify="center").pack(pady=(10, 8))

        sec_w = self._section(p, "ПАПКА ДЛЯ НАБЛЮДЕНИЯ")
        self.watch_dir_var = tk.StringVar(value=self._cfg.get("watch_dir", ""))
        self._file_row(sec_w, self.watch_dir_var, self._browse_watch_dir)

        sec_wo = self._section(p, "ПАПКА ДЛЯ ОТЧЁТОВ")
        self.watch_out_var = tk.StringVar(value=self._cfg.get("watch_output_dir", ""))
        self._file_row(sec_wo, self.watch_out_var, self._browse_watch_out)

        # Статус-индикатор
        status_frame = tk.Frame(p, bg=C_BG)
        status_frame.pack(pady=(8, 4))
        self.watch_status_var = tk.StringVar(value="⏹  Остановлен")
        self.watch_status_lbl = tk.Label(
            status_frame, textvariable=self.watch_status_var,
            font=("Segoe UI", 11, "bold"),
            bg=C_BG, fg=C_MUTED)
        self.watch_status_lbl.pack()

        tk.Label(p, text="ЖУРНАЛ НАБЛЮДАТЕЛЯ",
                 font=("Segoe UI", 8, "bold"),
                 bg=C_BG, fg=C_MUTED).pack(anchor="w", padx=14, pady=(4, 2))
        self.watch_log_box = self._log_widget(p, height=8)

        btn_frame = tk.Frame(p, bg=C_BG)
        btn_frame.pack(pady=(4, 12))
        self.watch_start_btn = HoverButton(
            btn_frame, bg_normal=C_SUCCESS, bg_hover="#1E5E2A",
            text="▶  Запустить наблюдатель",
            font=("Segoe UI", 10, "bold"),
            fg=C_WHITE, command=self._start_watcher,
            padx=16, pady=9)
        self.watch_start_btn.pack(side="left", padx=(0, 10))

        self.watch_stop_btn = HoverButton(
            btn_frame, bg_normal=C_ERROR, bg_hover="#A93226",
            text="⏹  Остановить",
            font=("Segoe UI", 10, "bold"),
            fg=C_WHITE, command=self._stop_watcher,
            padx=16, pady=9, state="disabled")
        self.watch_stop_btn.pack(side="left")

    # ============================================================
    # ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ (browse / log)
    # ============================================================

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Выберите файл с заказами",
            filetypes=[("Excel файлы", "*.xlsx *.xls")])
        if path:
            self.input_var.set(path)
            self._cfg["last_input"] = path
            if not self.output_dir_var.get():
                d = os.path.dirname(path)
                self.output_dir_var.set(d)
                self._cfg["last_output_dir"] = d
            save_config(self._cfg)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Папка для отчёта")
        if path:
            self.output_dir_var.set(path)
            self._cfg["last_output_dir"] = path
            save_config(self._cfg)

    def _browse_cmp(self, var):
        path = filedialog.askopenfilename(
            title="Выберите файл с заказами",
            filetypes=[("Excel файлы", "*.xlsx *.xls")])
        if path:
            var.set(path)
            if not self.cmp_output_dir.get():
                self.cmp_output_dir.set(os.path.dirname(path))

    def _browse_cmp_output(self):
        path = filedialog.askdirectory(title="Папка для отчёта")
        if path:
            self.cmp_output_dir.set(path)

    def _browse_watch_dir(self):
        path = filedialog.askdirectory(title="Папка для наблюдения")
        if path:
            self.watch_dir_var.set(path)
            self._cfg["watch_dir"] = path
            if not self.watch_out_var.get():
                self.watch_out_var.set(path)
                self._cfg["watch_output_dir"] = path
            save_config(self._cfg)

    def _browse_watch_out(self):
        path = filedialog.askdirectory(title="Папка для отчётов")
        if path:
            self.watch_out_var.set(path)
            self._cfg["watch_output_dir"] = path
            save_config(self._cfg)

    def _log(self, message: str):
        self._write_log(self.log_box, message)

    def _cmp_log(self, message: str):
        self._write_log(self.cmp_log_box, message)

    def _watch_log(self, message: str):
        self.after(0, lambda: self._write_log(self.watch_log_box, message))

    def _write_log(self, box, message: str):
        ts = datetime.now().strftime("%H:%M:%S")
        # Цвет строки по содержимому
        tag = None
        if "✅" in message or "Готово" in message:
            tag = "ok"
        elif "❌" in message or "ОШИБКА" in message:
            tag = "err"
        elif "⚠" in message:
            tag = "warn"

        box.configure(state="normal")
        box.tag_config("ok",   foreground="#6ECB6E")
        box.tag_config("err",  foreground="#FF6B6B")
        box.tag_config("warn", foreground="#FFD080")

        line = f"[{ts}] {message}\n"
        if tag:
            box.insert("end", line, tag)
        else:
            box.insert("end", line)
        box.see("end")
        box.configure(state="disabled")
        self._set_status(message[:80])

    # ============================================================
    # ПРОГРЕСС-БАР: шаги
    # ============================================================

    ANALYSIS_STEPS = [
        "Загрузка файла",
        "Проверка качества",
        "Клиенты",
        "Менеджеры",
        "Отрасли",
        "Сезонность",
        "RFM-анализ",
        "Лояльность",
        "План",
        "Номенклатура",
        "Сохранение",
    ]

    def _make_progress_log(self, bar, pct_lbl, total_steps):
        """Возвращает log-функцию, которая двигает прогресс-бар."""
        step_keywords = {
            "загрузк": 1, "качеств": 2, "клиент": 3, "менеджер": 4,
            "отрасл": 5, "сезонн": 6, "rfm": 7, "лояльн": 8,
            "план": 9, "номенклатур": 10, "сохран": 11,
        }
        counter = [0]

        def log_fn(message: str):
            lower = message.lower()
            for kw, step in step_keywords.items():
                if kw in lower and step > counter[0]:
                    counter[0] = step
                    pct = int(step / total_steps * 100)
                    self.after(0, lambda p=pct: (
                        bar.config(value=p),
                        pct_lbl.config(text=f"{p}%")
                    ))
                    break
            self._log(message)
        return log_fn

    # ============================================================
    # ЗАПУСК: АНАЛИЗ
    # ============================================================

    def _start_analysis(self):
        input_path = self.input_var.get().strip()
        output_dir = self.output_dir_var.get().strip()
        if not input_path:
            messagebox.showwarning("Внимание", "Выберите входной файл.")
            return
        if not output_dir:
            messagebox.showwarning("Внимание", "Укажите папку для отчёта.")
            return

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_dir, f"аналитика_заказов_{ts}.xlsx")
        manager_plan = self._get_manager_plan_rub()

        self.main_bar.config(value=0)
        self.main_pct.config(text="0%")
        self.run_btn.configure(state="disabled", text="⏳ Выполняется...")
        self._start_anim()

        log_fn = self._make_progress_log(self.main_bar, self.main_pct,
                                          len(self.ANALYSIS_STEPS))
        threading.Thread(
            target=self._run_analysis_thread,
            args=(input_path, output_path, manager_plan, log_fn),
            daemon=True).start()

    def _run_analysis_thread(self, input_path, output_path, manager_plan, log_fn):
        try:
            log_fn(f"Начинаю анализ: {os.path.basename(input_path)}")
            if manager_plan:
                log_fn(f"Планы: {len(manager_plan)} менеджеров")
            run_analytics(input_path, output_path,
                          log=log_fn, manager_plan=manager_plan)
            log_fn(f"✅ Отчёт: {output_path}")
            self.after(0, lambda: (
                self.main_bar.config(value=100),
                self.main_pct.config(text="100%")
            ))
            self.after(0, lambda: self._on_success(output_path))
        except Exception as e:
            log_fn(f"❌ ОШИБКА: {e}")
            self.after(0, lambda: messagebox.showerror(
                "Ошибка", f"Анализ не выполнен:\n\n{e}"))
        finally:
            self._stop_anim()
            self.after(0, lambda: self.run_btn.configure(
                state="normal", text="▶  Запустить анализ"))

    # ============================================================
    # ЗАПУСК: СРАВНЕНИЕ
    # ============================================================

    def _start_comparison(self):
        path_a     = self.cmp_path_a.get().strip()
        path_b     = self.cmp_path_b.get().strip()
        output_dir = self.cmp_output_dir.get().strip()
        label_a    = self.cmp_label_a.get().strip() or "Период А"
        label_b    = self.cmp_label_b.get().strip() or "Период Б"
        if not path_a:
            messagebox.showwarning("Внимание", "Выберите файл для периода А.")
            return
        if not path_b:
            messagebox.showwarning("Внимание", "Выберите файл для периода Б.")
            return
        if not output_dir:
            messagebox.showwarning("Внимание", "Укажите папку для отчёта.")
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_dir, f"сравнение_{ts}.xlsx")
        self.cmp_run_btn.configure(state="disabled", text="⏳ Выполняется...")
        self._start_anim()
        threading.Thread(
            target=self._run_comparison_thread,
            args=(path_a, path_b, output_path, label_a, label_b),
            daemon=True).start()

    def _run_comparison_thread(self, path_a, path_b, output_path, label_a, label_b):
        try:
            self._cmp_log(f"Сравниваю: «{label_a}» vs «{label_b}»")
            run_comparison(path_a, path_b, output_path,
                           label_a, label_b, log=self._cmp_log)
            self._cmp_log(f"✅ Отчёт: {output_path}")
            self.after(0, lambda: self._on_success(output_path))
        except Exception as e:
            self._cmp_log(f"❌ ОШИБКА: {e}")
            self.after(0, lambda: messagebox.showerror(
                "Ошибка", f"Сравнение не выполнено:\n\n{e}"))
        finally:
            self._stop_anim()
            self.after(0, lambda: self.cmp_run_btn.configure(
                state="normal", text="▶  Сравнить периоды"))

    # ============================================================
    # НАБЛЮДАТЕЛЬ
    # ============================================================

    def _start_watcher(self):
        watch_dir = self.watch_dir_var.get().strip()
        out_dir   = self.watch_out_var.get().strip()
        if not watch_dir:
            messagebox.showwarning("Внимание", "Укажите папку для наблюдения.")
            return
        if not out_dir:
            messagebox.showwarning("Внимание", "Укажите папку для отчётов.")
            return
        if not os.path.isdir(watch_dir):
            messagebox.showerror("Ошибка", f"Папка не существует:\n{watch_dir}")
            return
        self._watcher = FolderWatcher(watch_dir, out_dir, self._watch_log)
        self._watcher.start()
        self.watch_status_var.set("👁  Наблюдатель активен")
        self.watch_status_lbl.config(fg=C_SUCCESS)
        self.watch_start_btn.configure(state="disabled")
        self.watch_stop_btn.configure(state="normal")

    def _stop_watcher(self):
        if self._watcher:
            self._watcher.stop()
            self._watcher = None
        self.watch_status_var.set("⏹  Остановлен")
        self.watch_status_lbl.config(fg=C_MUTED)
        self.watch_start_btn.configure(state="normal")
        self.watch_stop_btn.configure(state="disabled")

    # ── Диалог успеха ────────────────────────────────────────

    def _on_success(self, output_path):
        self._set_status(f"Готово: {os.path.basename(output_path)}")
        answer = messagebox.askyesno(
            "Готово!",
            f"Отчёт создан:\n{output_path}\n\nОткрыть папку?")
        if answer:
            os.startfile(os.path.dirname(output_path))


# ============================================================
# ТОЧКА ВХОДА
# ============================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()
