# ============================================================
# app.py — Аналитика заказов, Фонтанка.ру
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

C_ORANGE  = "#F38120"
C_DARK    = "#D06A10"
C_BG      = "#F5F5F5"
C_WHITE   = "#FFFFFF"
C_BORDER  = "#DEDEDE"
C_TEXT    = "#1A1A1A"
C_MUTED   = "#777777"
C_LOG_BG  = "#1C1C1E"
C_LOG_FG  = "#F0EDE8"
C_GREEN   = "#27AE60"
C_RED     = "#E74C3C"


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


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Аналитика заказов — Фонтанка.ру")
        self.geometry("740x660")
        self.resizable(False, False)
        self.configure(bg=C_BG)

        if os.path.exists(ICON_PATH):
            try:
                self.iconbitmap(ICON_PATH)
            except Exception:
                pass

        self._cfg = load_config()
        self._watcher = None
        self._plan_vars = {}
        self._anim_running = False
        self._anim_step = 0

        self._build_header()
        self._build_tabs()
        self._build_statusbar()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _update_clock(self):
        try:
            from datetime import timezone, timedelta
            msk = timezone(timedelta(hours=3))
            now = datetime.now(tz=msk)
        except Exception:
            now = datetime.now()
        self._clock_var.set(now.strftime("%d.%m.%Y  %H:%M  МСК"))
        self.after(60000, self._update_clock)

    def _on_close(self):
        if self._watcher:
            self._watcher.stop()
        save_config(self._cfg)
        self.destroy()

    # ── Шапка ────────────────────────────────────────────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=C_ORANGE, height=64)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        # Логотип текстом
        tk.Label(hdr, text="фонтанка.ру",
                 font=("Georgia", 20, "bold italic"),
                 bg=C_ORANGE, fg=C_WHITE).pack(side="left", padx=20, pady=14)

        # Разделитель
        tk.Frame(hdr, bg="#D06A10", width=1).pack(side="left", fill="y", pady=10, padx=4)

        # Подзаголовок
        tk.Label(hdr, text="АНАЛИТИКА КОММЕРЧЕСКОГО ОТДЕЛА",
                 font=("Segoe UI", 9),
                 bg=C_ORANGE, fg="#FFD8B0").pack(side="left", padx=16)

        # Дата + время (обновляется каждую минуту)
        self._clock_var = tk.StringVar()
        tk.Label(hdr, textvariable=self._clock_var,
                 font=("Segoe UI", 11, "bold"),
                 bg=C_ORANGE, fg=C_WHITE).pack(side="right", padx=20)
        self._update_clock()

    # ── Вкладки ──────────────────────────────────────────────
    def _build_tabs(self):
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("F.TNotebook", background=C_BG, borderwidth=0)
        style.configure("F.TNotebook.Tab",
                        font=("Segoe UI", 10), padding=[14, 7],
                        background=C_BORDER, foreground=C_MUTED)
        style.map("F.TNotebook.Tab",
                  background=[("selected", C_WHITE)],
                  foreground=[("selected", C_ORANGE)],
                  font=[("selected", ("Segoe UI", 10, "bold"))])

        nb = ttk.Notebook(self, style="F.TNotebook")
        nb.pack(fill="both", expand=True)
        self.notebook = nb

        self.tab_main    = tk.Frame(nb, bg=C_BG)
        self.tab_compare = tk.Frame(nb, bg=C_BG)
        self.tab_plan    = tk.Frame(nb, bg=C_BG)
        self.tab_watch   = tk.Frame(nb, bg=C_BG)

        nb.add(self.tab_main,    text="  Анализ  ")
        nb.add(self.tab_compare, text="  Сравнение  ")
        nb.add(self.tab_plan,    text="  План  ")
        nb.add(self.tab_watch,   text="  Наблюдатель  ")

        self._build_tab_main()
        self._build_tab_compare()
        self._build_tab_plan()
        self._build_tab_watch()

    # ── Статус-бар ───────────────────────────────────────────
    def _build_statusbar(self):
        sb = tk.Frame(self, bg=C_BORDER, height=24)
        sb.pack(fill="x", side="bottom")
        sb.pack_propagate(False)
        self._status_var = tk.StringVar(value="Готов к работе")
        tk.Label(sb, textvariable=self._status_var,
                 font=("Segoe UI", 8), bg=C_BORDER,
                 fg=C_MUTED, anchor="w").pack(side="left", padx=10)
        self._anim_lbl = tk.Label(sb, text="", font=("Segoe UI", 8),
                                   bg=C_BORDER, fg=C_ORANGE)
        self._anim_lbl.pack(side="right", padx=10)
        tk.Label(sb, text="v1.5", font=("Segoe UI", 8),
                 bg=C_BORDER, fg=C_MUTED).pack(side="right", padx=10)

    def _start_anim(self):
        self._anim_running = True
        self._anim_step = 0
        self._tick_anim()

    def _stop_anim(self):
        self._anim_running = False
        self._anim_lbl.config(text="")

    def _tick_anim(self):
        if not self._anim_running:
            return
        f = ["●○○", "○●○", "○○●", "○●○"]
        self._anim_lbl.config(text=f[self._anim_step % 4])
        self._anim_step += 1
        self.after(250, self._tick_anim)

    # ── Общие строители ──────────────────────────────────────
    def _card(self, parent, title):
        """Белая карточка с оранжевой полоской заголовка."""
        wrap = tk.Frame(parent, bg=C_BG)
        wrap.pack(fill="x", padx=14, pady=4)
        tk.Label(wrap, text=f"  {title}",
                 font=("Segoe UI", 8, "bold"),
                 bg=C_ORANGE, fg=C_WHITE,
                 anchor="w", height=2).pack(fill="x")
        inner = tk.Frame(wrap, bg=C_WHITE,
                         highlightbackground=C_BORDER,
                         highlightthickness=1)
        inner.pack(fill="x")
        pad = tk.Frame(inner, bg=C_WHITE)
        pad.pack(fill="x", padx=10, pady=6)
        return pad

    def _file_row(self, parent, var, cmd):
        row = tk.Frame(parent, bg=C_WHITE)
        row.pack(fill="x")
        tk.Entry(row, textvariable=var, font=("Segoe UI", 9),
                 state="readonly", relief="solid", bd=1,
                 bg="#FAFAFA").pack(side="left", fill="x", expand=True, padx=(0, 8))
        btn = tk.Button(row, text="Обзор...", font=("Segoe UI", 9),
                        command=cmd, bg=C_ORANGE, fg=C_WHITE,
                        activebackground=C_DARK, relief="flat",
                        padx=10, pady=3, cursor="hand2")
        btn.pack(side="left")
        return btn

    def _log_box(self, parent, height=7):
        box = scrolledtext.ScrolledText(
            parent, font=("Consolas", 9), height=height,
            state="disabled", bg=C_LOG_BG, fg=C_LOG_FG,
            relief="flat", selectbackground=C_ORANGE)
        box.pack(fill="both", expand=True, padx=14, pady=(0, 4))
        box.tag_config("ok",   foreground="#6ECB6E")
        box.tag_config("err",  foreground="#FF6B6B")
        box.tag_config("warn", foreground="#FFD080")
        return box

    def _progress_row(self, parent):
        row = tk.Frame(parent, bg=C_BG)
        row.pack(fill="x", padx=14, pady=(2, 4))
        style = ttk.Style()
        style.configure("F.Horizontal.TProgressbar",
                        troughcolor=C_BORDER,
                        background=C_ORANGE, thickness=8)
        bar = ttk.Progressbar(row, style="F.Horizontal.TProgressbar",
                               mode="determinate", length=560)
        bar.pack(side="left", fill="x", expand=True, padx=(0, 8))
        lbl = tk.Label(row, text="0%", font=("Segoe UI", 9, "bold"),
                       bg=C_BG, fg=C_ORANGE, width=4)
        lbl.pack(side="left")
        return bar, lbl

    def _run_btn(self, parent, text, cmd, color=None):
        c = color or C_ORANGE
        d = C_DARK if c == C_ORANGE else "#1E5E2A"
        b = tk.Button(parent, text=text,
                      font=("Segoe UI", 11, "bold"),
                      command=cmd, bg=c, fg=C_WHITE,
                      activebackground=d, relief="flat",
                      padx=22, pady=10, cursor="hand2")
        b.pack(pady=(4, 12))
        b.bind("<Enter>", lambda e: b.config(bg=d))
        b.bind("<Leave>", lambda e: b.config(bg=c))
        return b

    def _write_log(self, box, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        tag = ("ok"   if "✅" in msg or "Готово" in msg else
               "err"  if "❌" in msg or "ОШИБКА" in msg else
               "warn" if "⚠" in msg else None)
        box.configure(state="normal")
        line = f"[{ts}] {msg}\n"
        if tag:
            box.insert("end", line, tag)
        else:
            box.insert("end", line)
        box.see("end")
        box.configure(state="disabled")
        self._status_var.set(msg[:90])

    # ── Вкладка 1: Анализ ────────────────────────────────────
    def _build_tab_main(self):
        p = self.tab_main
        tk.Label(p, text="Выберите файл и запустите анализ",
                 font=("Segoe UI", 10), bg=C_BG,
                 fg=C_MUTED).pack(pady=(10, 4))

        c1 = self._card(p, "ВХОДНОЙ ФАЙЛ")
        self.input_var = tk.StringVar(value=self._cfg.get("last_input", ""))
        self._file_row(c1, self.input_var, self._browse_input)

        c2 = self._card(p, "ПАПКА ДЛЯ ОТЧЁТА")
        self.output_dir_var = tk.StringVar(value=self._cfg.get("last_output_dir", ""))
        self._file_row(c2, self.output_dir_var, self._browse_output)

        tk.Label(p, text="ПРОГРЕСС", font=("Segoe UI", 8, "bold"),
                 bg=C_BG, fg=C_MUTED).pack(anchor="w", padx=14, pady=(8, 0))
        self.main_bar, self.main_pct = self._progress_row(p)

        tk.Label(p, text="ЖУРНАЛ", font=("Segoe UI", 8, "bold"),
                 bg=C_BG, fg=C_MUTED).pack(anchor="w", padx=14, pady=(4, 2))
        self.log_box = self._log_box(p, height=7)
        self.run_btn = self._run_btn(p, "▶  Запустить анализ",
                                      self._start_analysis)

    # ── Вкладка 2: Сравнение ─────────────────────────────────
    def _build_tab_compare(self):
        p = self.tab_compare
        tk.Label(p, text="Сравните два периода",
                 font=("Segoe UI", 10), bg=C_BG,
                 fg=C_MUTED).pack(pady=(10, 4))

        ca = self._card(p, "ПЕРИОД А")
        self.cmp_path_a = tk.StringVar()
        self._file_row(ca, self.cmp_path_a,
                       lambda: self._browse_cmp(self.cmp_path_a))
        lra = tk.Frame(ca, bg=C_WHITE)
        lra.pack(fill="x", pady=(4, 0))
        tk.Label(lra, text="Название:", font=("Segoe UI", 9),
                 bg=C_WHITE, fg=C_MUTED).pack(side="left")
        self.cmp_label_a = tk.StringVar(value="Период А")
        tk.Entry(lra, textvariable=self.cmp_label_a,
                 font=("Segoe UI", 9), width=22,
                 relief="solid", bd=1).pack(side="left", padx=8)

        cb = self._card(p, "ПЕРИОД Б")
        self.cmp_path_b = tk.StringVar()
        self._file_row(cb, self.cmp_path_b,
                       lambda: self._browse_cmp(self.cmp_path_b))
        lrb = tk.Frame(cb, bg=C_WHITE)
        lrb.pack(fill="x", pady=(4, 0))
        tk.Label(lrb, text="Название:", font=("Segoe UI", 9),
                 bg=C_WHITE, fg=C_MUTED).pack(side="left")
        self.cmp_label_b = tk.StringVar(value="Период Б")
        tk.Entry(lrb, textvariable=self.cmp_label_b,
                 font=("Segoe UI", 9), width=22,
                 relief="solid", bd=1).pack(side="left", padx=8)

        co = self._card(p, "ПАПКА ДЛЯ ОТЧЁТА")
        self.cmp_output_dir = tk.StringVar(value=self._cfg.get("last_output_dir", ""))
        self._file_row(co, self.cmp_output_dir, self._browse_cmp_output)

        tk.Label(p, text="ЖУРНАЛ", font=("Segoe UI", 8, "bold"),
                 bg=C_BG, fg=C_MUTED).pack(anchor="w", padx=14, pady=(8, 2))
        self.cmp_log_box = self._log_box(p, height=6)
        self.cmp_run_btn = self._run_btn(p, "▶  Сравнить периоды",
                                          self._start_comparison)

    # ── Вкладка 3: План ──────────────────────────────────────
    def _build_tab_plan(self):
        p = self.tab_plan
        tk.Label(p, text="Плановые суммы по менеджерам (тыс. руб.)",
                 font=("Segoe UI", 10), bg=C_BG, fg=C_MUTED).pack(pady=(10, 2))
        tk.Label(p, text="Сохраняется автоматически. Применяется при каждом запуске анализа.",
                 font=("Segoe UI", 8), bg=C_BG, fg=C_MUTED).pack()

        outer = tk.Frame(p, bg=C_BG)
        outer.pack(fill="both", expand=True, padx=14, pady=8)
        canvas = tk.Canvas(outer, bg=C_BG, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self._plan_inner = tk.Frame(canvas, bg=C_WHITE)
        self._plan_inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._plan_inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # Заголовок
        hdr = tk.Frame(self._plan_inner, bg=C_ORANGE)
        hdr.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Label(hdr, text="  Менеджер", font=("Segoe UI", 9, "bold"),
                 bg=C_ORANGE, fg=C_WHITE, width=28,
                 anchor="w", pady=5).pack(side="left")
        tk.Label(hdr, text="План, тыс. руб.", font=("Segoe UI", 9, "bold"),
                 bg=C_ORANGE, fg=C_WHITE, width=14,
                 pady=5).pack(side="left")

        saved = self._cfg.get("manager_plan", {})
        managers = list(dict.fromkeys([
            "Петрова Э.", "Егорова И.", "Демахина Е.",
            "Дубровская Т.", "Долгополова Н.", "Демидова Г.",
            "Сосновик А.", "Киселёва Е.", "Лея К.",
            "Борисенко Е.", "Сорокина Т.", "Шевчук Е.", "Карпус О.",
        ] + list(saved.keys())))
        for i, name in enumerate(managers, 1):
            self._add_plan_row(name, str(saved.get(name, "")), i)

        add = tk.Frame(p, bg=C_BG)
        add.pack(fill="x", padx=14, pady=(4, 8))
        tk.Label(add, text="Добавить:", font=("Segoe UI", 9),
                 bg=C_BG).pack(side="left")
        self._new_mgr = tk.StringVar()
        tk.Entry(add, textvariable=self._new_mgr,
                 font=("Segoe UI", 9), width=22,
                 relief="solid", bd=1).pack(side="left", padx=8)
        tk.Button(add, text="Добавить", font=("Segoe UI", 9),
                  command=self._add_manager_ui,
                  bg=C_ORANGE, fg=C_WHITE,
                  activebackground=C_DARK,
                  relief="flat", padx=10, pady=3,
                  cursor="hand2").pack(side="left")

    def _add_plan_row(self, name, val="", idx=None):
        if idx is None:
            idx = len(self._plan_vars) + 1
        bg = C_WHITE if idx % 2 == 0 else "#FFF7F0"
        tk.Label(self._plan_inner, text=f"  {name}",
                 font=("Segoe UI", 9), bg=bg, anchor="w",
                 pady=4).grid(row=idx, column=0, sticky="ew")
        var = tk.StringVar(value=val)
        var.trace_add("write",
            lambda *_, n=name, v=var: self._save_plan(n, v))
        self._plan_vars[name] = var
        tk.Entry(self._plan_inner, textvariable=var,
                 font=("Segoe UI", 9), width=14,
                 relief="solid", bd=1, bg=bg).grid(
                     row=idx, column=1, pady=1, padx=6)

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

    # ── Вкладка 4: Наблюдатель ───────────────────────────────
    def _build_tab_watch(self):
        p = self.tab_watch
        tk.Label(p,
                 text="Автоматически запускает анализ через 5 минут\n"
                      "после появления нового .xlsx в папке",
                 font=("Segoe UI", 10), bg=C_BG,
                 fg=C_MUTED, justify="center").pack(pady=(10, 8))

        cw = self._card(p, "ПАПКА ДЛЯ НАБЛЮДЕНИЯ")
        self.watch_dir_var = tk.StringVar(value=self._cfg.get("watch_dir", ""))
        self._file_row(cw, self.watch_dir_var, self._browse_watch_dir)

        co = self._card(p, "ПАПКА ДЛЯ ОТЧЁТОВ")
        self.watch_out_var = tk.StringVar(value=self._cfg.get("watch_output_dir", ""))
        self._file_row(co, self.watch_out_var, self._browse_watch_out)

        self.watch_status_var = tk.StringVar(value="⏹  Остановлен")
        self.watch_status_lbl = tk.Label(
            p, textvariable=self.watch_status_var,
            font=("Segoe UI", 11, "bold"),
            bg=C_BG, fg=C_MUTED)
        self.watch_status_lbl.pack(pady=(10, 4))

        tk.Label(p, text="ЖУРНАЛ", font=("Segoe UI", 8, "bold"),
                 bg=C_BG, fg=C_MUTED).pack(anchor="w", padx=14, pady=(4, 2))
        self.watch_log_box = self._log_box(p, height=8)

        bf = tk.Frame(p, bg=C_BG)
        bf.pack(pady=(4, 12))
        self.watch_start_btn = tk.Button(
            bf, text="▶  Запустить наблюдатель",
            font=("Segoe UI", 10, "bold"),
            command=self._start_watcher,
            bg=C_GREEN, fg=C_WHITE,
            activebackground="#1E5E2A",
            relief="flat", padx=16, pady=9, cursor="hand2")
        self.watch_start_btn.pack(side="left", padx=(0, 10))
        self.watch_stop_btn = tk.Button(
            bf, text="⏹  Остановить",
            font=("Segoe UI", 10, "bold"),
            command=self._stop_watcher,
            bg=C_RED, fg=C_WHITE,
            activebackground="#A93226",
            relief="flat", padx=16, pady=9,
            cursor="hand2", state="disabled")
        self.watch_stop_btn.pack(side="left")

    # ── Browse ───────────────────────────────────────────────
    def _browse_input(self):
        p = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls")])
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
        p = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls")])
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

    # ── Запуск анализа ───────────────────────────────────────
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
        self.run_btn.configure(state="disabled", text="⏳ Выполняется...")
        self._start_anim()
        log_fn = self._progress_log(self.main_bar, self.main_pct)
        threading.Thread(
            target=self._thread_analysis,
            args=(inp, out_path, plan, log_fn),
            daemon=True).start()

    def _thread_analysis(self, inp, out_path, plan, log_fn):
        try:
            log_fn(f"Файл: {os.path.basename(inp)}")
            run_analytics(inp, out_path, log=log_fn, manager_plan=plan)
            log_fn(f"✅ Готово: {out_path}")
            self.after(0, lambda: (
                self.main_bar.config(value=100),
                self.main_pct.config(text="100%")))
            self.after(0, lambda: self._on_success(out_path))
        except Exception as e:
            log_fn(f"❌ ОШИБКА: {e}")
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self._stop_anim()
            self.after(0, lambda: self.run_btn.configure(
                state="normal", text="▶  Запустить анализ"))

    # ── Запуск сравнения ─────────────────────────────────────
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
        self.cmp_run_btn.configure(state="disabled", text="⏳ Выполняется...")
        self._start_anim()
        threading.Thread(
            target=self._thread_comparison,
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
                state="normal", text="▶  Сравнить периоды"))

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
        self.watch_status_var.set("👁  Наблюдатель активен")
        self.watch_status_lbl.config(fg=C_GREEN)
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

    def _on_success(self, path):
        self._status_var.set(f"Готово: {os.path.basename(path)}")
        if messagebox.askyesno("Готово!", f"Отчёт создан:\n{path}\n\nОткрыть папку?"):
            os.startfile(os.path.dirname(path))


if __name__ == "__main__":
    app = App()
    app.mainloop()
