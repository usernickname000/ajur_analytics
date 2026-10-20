# ============================================================
# app.py — Аналитика заказов, Фонтанка.ру  v3.1
# Окно в стиле Windows 11: тема — ui_theme.py
# ============================================================

import sys
import os
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import json
import queue
import shutil
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import crm_data
import ui_theme as ui
from ui_theme import ICON, T, px
from analytics import RECON_BAD_PCT, RECON_WARN_PCT, recon_verdict, run_analytics
from comparison import run_comparison
from dashboard import generate_dashboard

try:
    from external_income_editor import ExternalIncomeEditor
except ImportError:
    ExternalIncomeEditor = None
try:
    from verified_figures_editor import VerifiedFiguresEditor
except ImportError:
    VerifiedFiguresEditor = None
try:
    from client_aliases_editor import ClientAliasesEditor
except ImportError:
    ClientAliasesEditor = None
try:
    from dohody_import_dialog import DohodyImportDialog
except ImportError:
    DohodyImportDialog = None

VERSION = "3.1"

# В собранном onefile-exe __file__ указывает во временную папку распаковки
# (sys._MEIPASS), которая удаляется при выходе: правки JSON и config
# пропадали бы. Записываемые файлы держим рядом с exe, ресурсы (иконка,
# дефолтные JSON) берём из папки распаковки.
RES_DIR = _HERE
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
    for _name in ("external_income.json", "verified_figures.json", "client_aliases.json"):
        _src, _dst = os.path.join(RES_DIR, _name), os.path.join(APP_DIR, _name)
        if os.path.exists(_src) and not os.path.exists(_dst):
            try:
                shutil.copy2(_src, _dst)
            except OSError:
                pass
else:
    APP_DIR = RES_DIR

CONFIG_PATH          = os.path.join(APP_DIR, "config.json")
ICON_PATH            = os.path.join(RES_DIR, "icon.ico")
EXTERNAL_INCOME_PATH = os.path.join(APP_DIR, "external_income.json")
VERIFIED_PATH        = os.path.join(APP_DIR, "verified_figures.json")
ALIASES_PATH         = os.path.join(APP_DIR, "client_aliases.json")

# Анализ и редактор алиасов должны работать с одним и тем же файлом
crm_data.set_aliases_path(ALIASES_PATH)

DEFAULT_MANAGERS = [
    "Петрова Э.", "Егорова И.", "Демахина Е.",
    "Дубровская Т.", "Долгополова Н.", "Демидова Г.",
    "Сосновик А.", "Киселёва Е.", "Лея К.",
    "Борисенко Е.", "Сорокина Т.", "Шевчук Е.", "Карпус О.",
]

DEFAULT_VERIFIED = {
    "_описание": "Верифицированные годовые итоги из бухгалтерии.",
    "_год": 2025,
    "total_with_prog": 482404000,
    "total_with_barter_no_prog": 363000166,
    "advertising_no_events": 243005820,
    "programmatic_external": 119403897,
    "other_external_income": 41057102,
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


def _read_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _fmt_k(v):
    return f"{v:,.0f}".replace(",", " ")


# ============================================================
# TOAST — уведомление в правом нижнем углу
# ============================================================
class Toast(tk.Toplevel):
    def __init__(self, parent, message, kind="ok", action_text=None, action_cmd=None, duration=4000):
        super().__init__(parent)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.attributes("-alpha", 0.0)
        self.configure(bg=T["stroke"])

        inner = tk.Frame(self, bg=T["card"])
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        row = tk.Frame(inner, bg=T["card"])
        row.pack(fill="x", padx=px(16), pady=(px(14), px(14) if not action_text else px(6)))
        ui.label(row, ICON.get(kind, ICON["info"]), "icon_sm", fg=kind).pack(side="left", anchor="n", pady=px(2))
        ui.label(row, message, justify="left", wraplength=px(280)).pack(side="left", padx=(px(12), 0))
        close = ui.label(row, ICON["close"], "icon_sm", fg="text2", cursor="hand2")
        close.pack(side="right", anchor="n", padx=(px(12), 0), pady=px(2))
        close.bind("<Button-1>", lambda e: self._dismiss())

        if action_text and action_cmd:
            ui.button(inner, action_text, lambda: (action_cmd(), self._dismiss())).pack(
                anchor="w", padx=(px(44), 0), pady=(0, px(14)))

        self.update_idletasks()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        self._x = self.winfo_screenwidth() - w - px(20)
        self._y = self.winfo_screenheight() - h - px(60)
        self.geometry(f"{w}x{h}+{self._x}+{self._y}")
        ui.style_window(self, caption=False)
        self._duration = duration
        self._alpha = 0.0
        self._fade(+1)

    def _fade(self, direction):
        self._alpha = min(max(self._alpha + 0.12 * direction, 0.0), 1.0)
        try:
            self.attributes("-alpha", self._alpha)
        except tk.TclError:
            return
        if direction > 0 and self._alpha < 1:
            self.after(16, self._fade, +1)
        elif direction > 0:
            self.after(self._duration, self._fade, -1)
        elif self._alpha > 0:
            self.after(16, self._fade, -1)
        else:
            self.destroy()

    def _dismiss(self):
        self._duration = 0
        self._fade(-1)


# ============================================================
# ГЛАВНОЕ ОКНО
# ============================================================
class App(tk.Tk):
    NAV_ITEMS = [
        ("analysis", ICON["analysis"], "Анализ"),
        ("compare",  ICON["compare"],  "Сравнение"),
        ("plan",     ICON["plan"],     "Планы"),
    ]
    NAV_BOTTOM = [("data", ICON["data"], "Данные")]

    # Шаги прогресса анализа: ключевое слово из лога → номер шага
    STEPS = {
        "загрузк": 1, "качеств": 2, "клиент": 3, "менеджер": 4,
        "отрасл": 5, "сезонн": 6, "rfm": 7, "лояльн": 8,
        "план": 9, "номенклатур": 10, "сигнал": 11,
        "бухгалт": 12, "сохран": 13, "граф": 14,
    }

    def __init__(self):
        super().__init__()
        ui.init(self)
        self.title("Аналитика заказов — Фонтанка.ру")
        self.geometry(f"{px(1100)}x{px(740)}")
        self.minsize(px(940), px(640))
        if os.path.exists(ICON_PATH):
            try:
                self.iconbitmap(default=ICON_PATH)
            except Exception:
                pass

        self._cfg = load_config()
        self._date_by = tk.StringVar(value=self._cfg.get("date_by", "order"))
        self._pages, self._nav = {}, {}
        self._active = None
        self._plan_vars = {}
        self._tasks = 0
        self._log_count = 0
        self._last_folder = None
        self._scroll_canvases = set()
        self._ui_queue = queue.Queue()

        self._build()
        self._drain_ui_queue()
        self._show_page("analysis")
        ui.style_window(self)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        self._cfg["date_by"] = self._date_by.get()
        save_config(self._cfg)
        self.destroy()

    # ── Вызовы из рабочих потоков ────────────────────────────
    # Tk можно трогать только из главного потока, а after() из чужого потока
    # ненадёжен. Потоки кладут задание в очередь, главный поток разбирает её.
    def _ui(self, fn, *args):
        if threading.current_thread() is threading.main_thread():
            fn(*args)
        else:
            self._ui_queue.put((fn, args))

    def _drain_ui_queue(self):
        try:
            while True:
                fn, args = self._ui_queue.get_nowait()
                fn(*args)
        except queue.Empty:
            pass
        self.after(50, self._drain_ui_queue)

    def _toast(self, message, kind="ok", action_text=None, action_cmd=None, duration=4000):
        try:
            Toast(self, message, kind, action_text, action_cmd, duration)
        except Exception:
            pass

    # ============================================================
    # КАРКАС: сайдбар | рабочая область (прогресс, страница, журнал)
    # ============================================================
    def _build(self):
        side = tk.Frame(self, bg=T["bg"], width=px(220))
        side.pack(side="left", fill="y")
        side.pack_propagate(False)

        brand = tk.Frame(side, bg=T["bg"])
        brand.pack(fill="x", padx=px(20), pady=(px(22), px(18)))
        ui.label(brand, "фонтанка.ру", "brand", fg="brand", bg="bg").pack(anchor="w")
        ui.label(brand, "Аналитика заказов", "caption", fg="text2", bg="bg").pack(anchor="w")

        for page_id, icon, text in self.NAV_ITEMS:
            self._nav[page_id] = self._nav_item(side, page_id, icon, text)
        ui.label(side, f"Версия {VERSION}", "caption", fg="text3", bg="bg").pack(
            side="bottom", anchor="w", padx=px(20), pady=(0, px(14)))
        for page_id, icon, text in self.NAV_BOTTOM:
            self._nav[page_id] = self._nav_item(side, page_id, icon, text, bottom=True)

        work = tk.Frame(self, bg=T["bg"])
        work.pack(side="left", fill="both", expand=True)

        # Тонкий прогресс над страницей — виден только во время работы
        self._progress_holder = tk.Frame(work, bg=T["bg"], height=px(3))
        self._progress_holder.pack(fill="x")
        self._progress_holder.pack_propagate(False)
        self._progress = ttk.Progressbar(self._progress_holder, style="Thin.Horizontal.TProgressbar",
                                         mode="determinate", maximum=100)

        self._build_log(work)
        self._page_area = tk.Frame(work, bg=T["bg"])
        self._page_area.pack(fill="both", expand=True)

        for page_id, _, _ in self.NAV_ITEMS + self.NAV_BOTTOM:
            self._pages[page_id] = tk.Frame(self._page_area, bg=T["bg"])
        self._build_page_analysis()
        self._build_page_compare()
        self._build_page_plan()
        self._build_page_data()

    def _nav_item(self, parent, page_id, icon, text, bottom=False):
        row = tk.Frame(parent, bg=T["bg"], cursor="hand2", height=px(38))
        row.pack(fill="x", padx=px(8), pady=px(1), side="bottom" if bottom else "top")
        row.pack_propagate(False)
        pill = tk.Frame(row, bg=T["bg"], width=px(3))
        pill.pack(side="left", fill="y", pady=px(10))
        ic = ui.label(row, icon, "icon", fg="text", bg="bg")
        ic.pack(side="left", padx=(px(12), px(12)))
        tx = ui.label(row, text, "body", fg="text", bg="bg")
        tx.pack(side="left")
        item = {"row": row, "pill": pill, "icon": ic, "text": tx}

        def paint(bg):
            for w in (row, ic, tx):
                w.configure(bg=bg)
            pill.configure(bg=T["brand"] if self._active == page_id else bg)

        item["paint"] = paint
        for w in (row, ic, tx):
            w.bind("<Button-1>", lambda e: self._show_page(page_id))
            w.bind("<Enter>", lambda e: self._active != page_id and paint("#EAEAEA"))
            w.bind("<Leave>", lambda e: self._active != page_id and paint(T["bg"]))
        return item

    def _show_page(self, page_id):
        self._active = page_id
        for pid, item in self._nav.items():
            item["paint"](T["subtle_press"] if pid == page_id else T["bg"])
            item["text"].configure(font=ui.FONTS["strong"] if pid == page_id else ui.FONTS["body"])
        for pid, page in self._pages.items():
            if pid == page_id:
                page.pack(fill="both", expand=True)
            else:
                page.pack_forget()
        if page_id == "data":
            self._refresh_data_page()

    # ── Общие элементы страниц ───────────────────────────────
    def _page_header(self, page, title, subtitle):
        hdr = tk.Frame(page, bg=T["bg"])
        hdr.pack(fill="x", padx=px(32), pady=(px(26), px(16)))
        left = tk.Frame(hdr, bg=T["bg"])
        left.pack(side="left", fill="x", expand=True)
        ui.label(left, title, "title", bg="bg").pack(anchor="w")
        ui.label(left, subtitle, "caption", fg="text2", bg="bg").pack(anchor="w", pady=(px(2), 0))
        actions = tk.Frame(hdr, bg=T["bg"])
        actions.pack(side="right", anchor="s")
        return actions

    def _section(self, page, title=None):
        outer, body = ui.card(page, title)
        outer.pack(fill="x", padx=px(32), pady=(0, px(12)))
        return body

    def _path_row(self, parent, row, caption, var, command):
        ui.label(parent, caption, "body", fg="text2").grid(row=row, column=0, sticky="w",
                                                            padx=(0, px(16)), pady=px(4))
        ttk.Entry(parent, textvariable=var, state="readonly").grid(row=row, column=1, sticky="ew", pady=px(4))
        ui.button(parent, "Обзор…", command).grid(row=row, column=2, padx=(px(8), 0), pady=px(4))
        parent.columnconfigure(1, weight=1)

    def _bind_mousewheel(self, canvas):
        # Одна глобальная привязка; прокручиваем тот холст, над которым курсор.
        # Редакторы регистрируют свои холсты здесь же (см. external_income_editor).
        self._scroll_canvases.add(canvas)
        if getattr(self, "_wheel_bound", False):
            return
        self._wheel_bound = True

        def _scroll(e):
            w = self.winfo_containing(e.x_root, e.y_root)
            while w is not None and w not in self._scroll_canvases:
                w = getattr(w, "master", None)
            if w is not None:
                w.yview_scroll(int(-1 * (e.delta / 120)), "units")
        self.bind_all("<MouseWheel>", _scroll)

    # ============================================================
    # ЖУРНАЛ — один на все страницы, по умолчанию свёрнут
    # ============================================================
    def _build_log(self, parent):
        wrap = tk.Frame(parent, bg=T["bg"])
        wrap.pack(fill="x", side="bottom", padx=px(32), pady=(px(4), px(18)))
        self._log_exp = ui.Expander(wrap, "Журнал")
        self._log_exp.pack(fill="x")
        self._status = ui.label(self._log_exp.header, "", "caption", fg="text3", bg="bg")
        self._status.pack(side="left", padx=(px(16), 0))
        self._open_folder_btn = ui.button(self._log_exp.header, "Открыть папку", self._open_last_folder,
                                          kind="Subtle", on="bg", icon=ICON["folder"])

        holder, body = ui.card(self._log_exp.body, pad=8)
        holder.pack(fill="x")
        self._log_box = tk.Text(
            body, font=ui.FONTS["mono"], height=9, state="disabled", relief="flat", bd=0,
            bg=T["card"], fg=T["text"], selectbackground=T["brand"], padx=px(4), pady=px(2), wrap="word")
        sb = ttk.Scrollbar(body, orient="vertical", command=self._log_box.yview,
                           style="Thin.Vertical.TScrollbar")
        self._log_box.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._log_box.pack(side="left", fill="both", expand=True)
        for tag, color in (("ok", T["ok"]), ("err", T["err"]), ("warn", T["warn"]), ("ts", T["text3"])):
            self._log_box.tag_config(tag, foreground=color)

    def _log(self, msg):
        # Логи приходят из рабочих потоков, а Tk-виджеты можно трогать только
        # из главного — перекладываем вызов в очередь событий Tk.
        if threading.current_thread() is not threading.main_thread():
            self._ui(self._log, msg)
            return
        tag = ("ok" if "✅" in msg or "Готово" in msg else
               "err" if "❌" in msg or "ОШИБКА" in msg else
               "warn" if "⚠" in msg else None)
        box = self._log_box
        box.configure(state="normal")
        box.insert("end", datetime.now().strftime("%H:%M:%S  "), "ts")
        box.insert("end", msg + "\n", tag or ())
        box.see("end")
        box.configure(state="disabled")
        self._log_count += 1
        self._log_exp.set_title(f"Журнал ({self._log_count})")
        self._status.configure(text=msg[:90])
        if tag in ("err", "warn"):
            self._log_exp.toggle(True)  # замечания не должны прятаться

    def _open_last_folder(self):
        if self._last_folder and os.path.isdir(self._last_folder):
            os.startfile(self._last_folder)

    def _report_ready(self, path, title):
        self._last_folder = os.path.dirname(path)
        self._open_folder_btn.pack(side="right")
        self._toast(f"{title}: {os.path.basename(path)}", "ok", "Открыть папку",
                    self._open_last_folder, duration=6000)

    # ============================================================
    # ФОНОВЫЕ ЗАДАЧИ: прогресс + блокировка кнопки + обработка ошибок
    # ============================================================
    def _run_task(self, button, busy_text, worker, determinate=False):
        idle_text = button.cget("text")
        button.configure(state="disabled", text=busy_text)
        self._tasks += 1
        if self._tasks == 1:
            self._progress.pack(fill="both", expand=True)
        if determinate:  # анализ — проценты по шагам; остальное — «бегущая» полоса
            self._progress.stop()
            self._progress.configure(mode="determinate", value=0)
        elif self._tasks == 1:
            self._progress.configure(mode="indeterminate")
            self._progress.start(12)

        def finish():
            button.configure(state="normal", text=idle_text)
            self._tasks = max(0, self._tasks - 1)
            if not self._tasks:
                self._progress.stop()
                self._progress.pack_forget()

        def run():
            try:
                worker()
            except Exception as e:
                # Копируем текст: после выхода из except имя `e` удаляется
                err = str(e)
                self._log(f"❌ ОШИБКА: {err}")
                self._ui(lambda: messagebox.showerror("Ошибка", err, parent=self))
            finally:
                self._ui(finish)

        threading.Thread(target=run, daemon=True).start()

    def _set_progress(self, value):
        if self._progress.cget("mode") == "determinate":
            self._progress.configure(value=value)

    # ============================================================
    # СТРАНИЦА: АНАЛИЗ
    # ============================================================
    def _build_page_analysis(self):
        p = self._pages["analysis"]
        actions = self._page_header(p, "Анализ заказов", "Выгрузка CRM → отчёт Excel и HTML-дашборд")
        self.run_btn = ui.button(actions, "Запустить анализ", self._start_analysis,
                                 kind="Accent", on="bg", icon=ICON["play"])
        self.run_btn.pack(side="right")
        self.dash_btn = ui.button(actions, "Открыть дашборд", self._start_dashboard,
                                  on="bg", icon=ICON["globe"])
        self.dash_btn.pack(side="right", padx=(0, px(8)))

        src = self._section(p, "Источник")
        grid = tk.Frame(src, bg=T["card"])
        grid.pack(fill="x")
        self.input_var = tk.StringVar(value=self._cfg.get("last_input", ""))
        self.output_dir_var = tk.StringVar(value=self._cfg.get("last_output_dir", ""))
        self._path_row(grid, 0, "Файл CRM", self.input_var, self._browse_input)
        self._path_row(grid, 1, "Папка отчёта", self.output_dir_var, self._browse_output)
        ui.label(grid, "Группировать по", "body", fg="text2").grid(row=2, column=0, sticky="w",
                                                                   pady=(px(10), 0))
        radios = tk.Frame(grid, bg=T["card"])
        radios.grid(row=2, column=1, columnspan=2, sticky="w", pady=(px(10), 0))
        for val, text in (("order", "дате заказа"), ("payment", "дате оплаты")):
            ttk.Radiobutton(radios, text=text, variable=self._date_by, value=val).pack(
                side="left", padx=(0, px(20)))

        # Результат: сверка и KPI — появляются после первого запуска
        self._result = tk.Frame(p, bg=T["bg"])
        self._recon_bar = ui.InfoBar(self._result)
        self._recon_bar.pack(fill="x", pady=(0, px(12)))
        kpis = tk.Frame(self._result, bg=T["bg"])
        kpis.pack(fill="x")
        self._kpi = {}
        for i, (key, caption) in enumerate((("crm", "CRM, тыс. ₽"), ("external", "Внешние, тыс. ₽"),
                                            ("grand", "Итого, тыс. ₽"), ("dev", "Отклонение от бухгалтерии"))):
            outer, body = ui.card(kpis, pad=14)
            outer.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else px(6), 0))
            kpis.columnconfigure(i, weight=1, uniform="kpi")
            ui.label(body, caption, "caption", fg="text2").pack(anchor="w")
            value = ui.label(body, "—", "kpi")
            value.pack(anchor="w", pady=(px(2), 0))
            self._kpi[key] = value

    def _show_result(self, result):
        pct = result.get("deviation_pct")
        grand = result.get("grand_total", 0)
        verified = result.get("verified_total", 0)
        by = "по дате оплаты" if result.get("date_by") == "payment" else "по дате заказа"

        if pct is None or verified <= 0:
            self._recon_bar.set("info", f"CRM + внешние доходы: {_fmt_k(grand)} тыс. ₽. "
                                        "Сверка с бухгалтерией не применима — подробности в журнале.")
        else:
            kind = "ok" if abs(pct) < RECON_WARN_PCT else "warn" if abs(pct) < RECON_BAD_PCT else "err"
            self._recon_bar.set(kind, f"Сверка {by}: CRM + внешние {_fmt_k(grand)} тыс. ₽ · "
                                      f"бухгалтерия {_fmt_k(verified)} тыс. ₽ · "
                                      f"{pct:+.2f}% — {recon_verdict(pct)}")

        self._kpi["crm"].configure(text=_fmt_k(result.get("crm_total", 0)))
        self._kpi["external"].configure(text=_fmt_k(result.get("full_external_total", 0)))
        self._kpi["grand"].configure(text=_fmt_k(grand))
        if pct is None:
            self._kpi["dev"].configure(text="—", fg=T["text3"])
        else:
            color = T["ok"] if abs(pct) < RECON_WARN_PCT else T["warn"] if abs(pct) < RECON_BAD_PCT else T["err"]
            self._kpi["dev"].configure(text=f"{pct:+.2f}%", fg=color)
        if not self._result.winfo_ismapped():
            self._result.pack(fill="x", padx=px(32))

    def _browse_input(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.input_var.set(path)
            self._cfg["last_input"] = path
            if not self.output_dir_var.get():
                self.output_dir_var.set(os.path.dirname(path))
                self._cfg["last_output_dir"] = os.path.dirname(path)
            save_config(self._cfg)

    def _browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir_var.set(path)
            self._cfg["last_output_dir"] = path
            save_config(self._cfg)

    def _start_analysis(self):
        inp = self.input_var.get().strip()
        out = self.output_dir_var.get().strip()
        if not inp:
            messagebox.showwarning("", "Выберите файл выгрузки CRM.", parent=self)
            return
        if not out:
            messagebox.showwarning("", "Укажите папку для отчёта.", parent=self)
            return
        out_path = os.path.join(out, f"аналитика_заказов_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        plan = self._get_plan_rub()
        date_by = self._date_by.get()  # Tk-переменную читаем в главном потоке
        step = [0]

        def log(msg):
            lo = msg.lower()
            for kw, n in self.STEPS.items():
                if kw in lo and n > step[0]:
                    step[0] = n
                    self._ui(self._set_progress, n / len(self.STEPS) * 100)
                    break
            self._log(msg)

        def work():
            log(f"Файл: {os.path.basename(inp)}")
            result = run_analytics(inp, out_path, log=log, manager_plan=plan, date_by=date_by,
                                   verified_path=VERIFIED_PATH, external_path=EXTERNAL_INCOME_PATH)
            self._ui(self._show_result, result)
            self._ui(self._report_ready, result.get("output_path", out_path), "Отчёт готов")

        self._run_task(self.run_btn, "Выполняется…", work, determinate=True)

    def _start_dashboard(self):
        inp = self.input_var.get().strip()
        if not inp:
            messagebox.showwarning("", "Выберите файл выгрузки CRM.", parent=self)
            return

        def work():
            path = generate_dashboard(inp, log=self._log)
            self._ui(self._report_ready, path, "Дашборд открыт")

        self._run_task(self.dash_btn, "Готовлю…", work)

    # ============================================================
    # СТРАНИЦА: СРАВНЕНИЕ
    # ============================================================
    def _build_page_compare(self):
        p = self._pages["compare"]
        actions = self._page_header(p, "Сравнение периодов", "Две выгрузки CRM: менеджеры, клиенты, движение клиентов")
        self.cmp_run_btn = ui.button(actions, "Сравнить", self._start_comparison,
                                     kind="Accent", on="bg", icon=ICON["compare"])
        self.cmp_run_btn.pack(side="right")

        self.cmp_path_a, self.cmp_path_b = tk.StringVar(), tk.StringVar()
        self.cmp_label_a = tk.StringVar(value="Период А")
        self.cmp_label_b = tk.StringVar(value="Период Б")
        self.cmp_output_dir = tk.StringVar(value=self._cfg.get("last_output_dir", ""))

        for title, path_var, label_var in (("Период А", self.cmp_path_a, self.cmp_label_a),
                                           ("Период Б", self.cmp_path_b, self.cmp_label_b)):
            body = self._section(p, title)
            grid = tk.Frame(body, bg=T["card"])
            grid.pack(fill="x")
            self._path_row(grid, 0, "Файл CRM", path_var, lambda v=path_var: self._browse_cmp(v))
            ui.label(grid, "Метка в отчёте", "body", fg="text2").grid(row=1, column=0, sticky="w", pady=px(4))
            ttk.Entry(grid, textvariable=label_var, width=24).grid(row=1, column=1, sticky="w", pady=px(4))

        body = self._section(p)
        grid = tk.Frame(body, bg=T["card"])
        grid.pack(fill="x")
        self._path_row(grid, 0, "Папка отчёта", self.cmp_output_dir, self._browse_cmp_output)

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

    def _start_comparison(self):
        pa, pb = self.cmp_path_a.get().strip(), self.cmp_path_b.get().strip()
        od = self.cmp_output_dir.get().strip()
        la, lb = self.cmp_label_a.get().strip(), self.cmp_label_b.get().strip()
        for value, msg in ((pa, "Выберите файл периода А."), (pb, "Выберите файл периода Б."),
                           (od, "Укажите папку для отчёта.")):
            if not value:
                messagebox.showwarning("", msg, parent=self)
                return
        out_path = os.path.join(od, f"сравнение_{datetime.now():%Y%m%d_%H%M%S}.xlsx")

        def work():
            self._log(f"Сравниваю «{la or 'Период А'}» и «{lb or 'Период Б'}»")
            run_comparison(pa, pb, out_path, la, lb, log=self._log)
            self._ui(self._report_ready, out_path, "Сравнение готово")

        self._run_task(self.cmp_run_btn, "Выполняется…", work)

    # ============================================================
    # СТРАНИЦА: ПЛАНЫ МЕНЕДЖЕРОВ
    # ============================================================
    def _build_page_plan(self):
        p = self._pages["plan"]
        actions = self._page_header(p, "Планы менеджеров", "Тыс. ₽ — попадают в лист «03а_План_менеджеров»")
        self._new_mgr = tk.StringVar()
        ui.label(actions, "Новый менеджер", "body", fg="text2", bg="bg").pack(side="left", padx=(0, px(8)))
        add = ttk.Entry(actions, textvariable=self._new_mgr, width=22)
        add.pack(side="left", padx=(0, px(8)))
        add.bind("<Return>", lambda e: self._add_manager_ui())
        ui.button(actions, "Добавить", self._add_manager_ui, on="bg", icon=ICON["add"]).pack(side="left")

        outer, body = ui.card(p, pad=8)
        outer.pack(fill="both", expand=True, padx=px(32), pady=(0, px(12)))
        canvas = tk.Canvas(body, bg=T["card"], highlightthickness=0)
        sb = ttk.Scrollbar(body, orient="vertical", command=canvas.yview, style="Thin.Vertical.TScrollbar")
        self._plan_inner = tk.Frame(canvas, bg=T["card"])
        self._plan_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win = canvas.create_window((0, 0), window=self._plan_inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._bind_mousewheel(canvas)

        ui.label(self._plan_inner, "Менеджер", "caption", fg="text2").grid(
            row=0, column=0, sticky="w", padx=px(12), pady=(px(6), px(4)))
        ui.label(self._plan_inner, "План, тыс. ₽", "caption", fg="text2").grid(
            row=0, column=1, sticky="w", padx=px(12), pady=(px(6), px(4)))
        self._plan_inner.columnconfigure(0, weight=1)

        saved = self._cfg.get("manager_plan", {})
        for name in dict.fromkeys(DEFAULT_MANAGERS + list(saved)):
            self._add_plan_row(name, str(saved.get(name, "")))

    def _add_plan_row(self, name, val=""):
        idx = len(self._plan_vars) + 1
        tk.Frame(self._plan_inner, bg=T["divider"], height=1).grid(row=idx * 2 - 1, column=0, columnspan=2,
                                                                   sticky="ew", padx=px(8))
        ui.label(self._plan_inner, name).grid(row=idx * 2, column=0, sticky="w", padx=px(12), pady=px(4))
        var = tk.StringVar(value=val)
        var.trace_add("write", lambda *_: self._schedule_plan_save())
        self._plan_vars[name] = var
        ttk.Entry(self._plan_inner, textvariable=var, width=16, justify="right").grid(
            row=idx * 2, column=1, padx=px(12), pady=px(4))

    def _add_manager_ui(self):
        name = self._new_mgr.get().strip()
        if not name:
            return
        if name in self._plan_vars:
            messagebox.showinfo("", f"«{name}» уже есть.", parent=self)
            return
        self._add_plan_row(name)
        self._new_mgr.set("")

    def _schedule_plan_save(self):
        # Не на каждое нажатие клавиши, а через полсекунды после последнего
        if getattr(self, "_plan_job", None):
            self.after_cancel(self._plan_job)
        self._plan_job = self.after(500, self._save_plans)

    def _save_plans(self):
        self._plan_job = None
        self._cfg["manager_plan"] = {n: v.get().strip() for n, v in self._plan_vars.items() if v.get().strip()}
        save_config(self._cfg)

    def _get_plan_rub(self):
        result = {}
        for name, var in self._plan_vars.items():
            val = var.get().strip().replace(",", ".").replace(" ", "")
            if val:
                try:
                    result[name] = float(val) * 1000
                except ValueError:
                    pass
        return result or None

    # ============================================================
    # СТРАНИЦА: ДАННЫЕ БУХГАЛТЕРИИ (редакторы справочников)
    # ============================================================
    def _build_page_data(self):
        p = self._pages["data"]
        self._page_header(p, "Данные бухгалтерии", "Справочники, которые используются при сверке")
        outer, body = ui.card(p, pad=4)
        outer.pack(fill="x", padx=px(32))
        self._data_rows = {}
        items = [
            ("verified", ICON["ok"], "Верифицированные цифры", self._open_verified_editor, VerifiedFiguresEditor),
            ("external", ICON["money"], "Внешние доходы", self._open_external_editor, ExternalIncomeEditor),
            ("import", ICON["import"], "Импорт «Доходы ГГГГ.xlsx»", self._open_dohody_import, DohodyImportDialog),
            ("aliases", ICON["link"], "Алиасы клиентов", self._open_aliases_editor, ClientAliasesEditor),
        ]
        first = True
        for key, icon, title, command, available in items:
            if available is None:
                continue
            if not first:
                tk.Frame(body, bg=T["divider"], height=1).pack(fill="x", padx=px(12))
            first = False
            self._data_rows[key] = self._settings_row(body, icon, title, command)

    def _settings_row(self, parent, icon, title, command):
        row = tk.Frame(parent, bg=T["card"], cursor="hand2")
        row.pack(fill="x")
        ic = ui.label(row, icon, "icon", fg="text2")
        ic.pack(side="left", padx=(px(16), px(16)), pady=px(14))
        text = tk.Frame(row, bg=T["card"])
        text.pack(side="left", fill="x", expand=True)
        t1 = ui.label(text, title, "body")
        t1.pack(anchor="w")
        t2 = ui.label(text, "", "caption", fg="text2")
        t2.pack(anchor="w")
        ch = ui.label(row, ICON["chev_right"], "icon_sm", fg="text2")
        ch.pack(side="right", padx=px(16))
        widgets = (row, ic, text, t1, t2, ch)

        def paint(bg):
            for w in widgets:
                w.configure(bg=bg)
        for w in widgets:
            w.bind("<Button-1>", lambda e: command())
            w.bind("<Enter>", lambda e: paint(T["subtle"]))
            w.bind("<Leave>", lambda e: paint(T["card"]))
        return t2

    def _refresh_data_page(self):
        """Подписи строк — по текущему содержимому файлов."""
        captions = {}
        vf = _read_json(VERIFIED_PATH)
        if vf:
            captions["verified"] = (f"Годовые итоги · {vf.get('_год', '—')} · всего с программатиком "
                                    f"{_fmt_k(vf.get('total_with_prog', 0))} ₽")
        ext = _read_json(EXTERNAL_INCOME_PATH)
        if ext:
            n = sum(1 for k, v in ext.items() if not k.startswith('_') and isinstance(v, dict))
            captions["external"] = f"Помесячно · {ext.get('_год', '—')} · статей: {n}"
        captions["import"] = "Обновить внешние доходы из мастер-файла бухгалтерии"
        al = _read_json(ALIASES_PATH)
        captions["aliases"] = (f"Групп: {len(al.get('groups', []))} — объединение написаний клиентов"
                               if al else "Объединение написаний клиентов")
        for key, lbl in self._data_rows.items():
            lbl.configure(text=captions.get(key, "Файл не найден"))

    def _after_editor(self, dialog):
        """После закрытия редактора — обновить подписи на странице «Данные»."""
        if dialog is not None and dialog.winfo_exists():
            self.wait_window(dialog)
        self._refresh_data_page()

    def _open_external_editor(self):
        if not os.path.exists(EXTERNAL_INCOME_PATH):
            messagebox.showerror("Нет файла", f"Файл не найден:\n{EXTERNAL_INCOME_PATH}", parent=self)
            return
        self._after_editor(ExternalIncomeEditor(self, EXTERNAL_INCOME_PATH))

    def _open_dohody_import(self):
        self._after_editor(DohodyImportDialog(self, EXTERNAL_INCOME_PATH))

    def _open_aliases_editor(self):
        if not os.path.exists(ALIASES_PATH):
            save_json_default(ALIASES_PATH, {"_описание": "Правила нормализации клиентов.", "groups": []})
        self._after_editor(ClientAliasesEditor(self, ALIASES_PATH))

    def _open_verified_editor(self):
        if not os.path.exists(VERIFIED_PATH):
            save_json_default(VERIFIED_PATH, DEFAULT_VERIFIED)
        self._after_editor(VerifiedFiguresEditor(self, VERIFIED_PATH))


def save_json_default(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showerror("Ошибка", str(e))


ui.enable_dpi_awareness()

if __name__ == "__main__":
    App().mainloop()
