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
# ГЛАВНОЕ ОКНО
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Аналитика заказов — Коммерческий отдел")
        self.geometry("720x620")
        self.resizable(False, False)
        self.configure(bg="#f0f2f5")

        self._cfg = load_config()
        self._watcher: FolderWatcher | None = None
        self._plan_vars: dict[str, tk.StringVar] = {}

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        if self._watcher:
            self._watcher.stop()
        save_config(self._cfg)
        self.destroy()

    def _build_ui(self):
        tk.Label(self, text="📊 Аналитика заказов",
                 font=("Segoe UI", 16, "bold"),
                 bg="#f0f2f5", fg="#1a1a2e").pack(pady=(14, 4))

        style = ttk.Style(self)
        style.configure("TNotebook", background="#f0f2f5", borderwidth=0)
        style.configure("TNotebook.Tab", font=("Segoe UI", 10), padding=[12, 6])

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=14, pady=(4, 0))

        self.tab_main    = tk.Frame(self.notebook, bg="#f0f2f5")
        self.tab_compare = tk.Frame(self.notebook, bg="#f0f2f5")
        self.tab_plan    = tk.Frame(self.notebook, bg="#f0f2f5")
        self.tab_watch   = tk.Frame(self.notebook, bg="#f0f2f5")

        self.notebook.add(self.tab_main,    text="  Анализ  ")
        self.notebook.add(self.tab_compare, text="  Сравнение  ")
        self.notebook.add(self.tab_plan,    text="  План  ")
        self.notebook.add(self.tab_watch,   text="  Наблюдатель  ")

        self._build_tab_main()
        self._build_tab_compare()
        self._build_tab_plan()
        self._build_tab_watch()

    # ============================================================
    # ВКЛАДКА 1: АНАЛИЗ
    # ============================================================

    def _build_tab_main(self):
        p = self.tab_main

        tk.Label(p, text="Выберите файл Excel с заказами и нажмите «Запустить»",
                 font=("Segoe UI", 10), bg="#f0f2f5", fg="#555").pack(pady=(12, 8))

        frame_in = tk.LabelFrame(p, text=" Входной файл (заказы) ",
                                  font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                  padx=10, pady=8)
        frame_in.pack(fill="x", padx=16, pady=4)
        self.input_var = tk.StringVar(value=self._cfg.get("last_input", ""))
        tk.Entry(frame_in, textvariable=self.input_var,
                 font=("Segoe UI", 9), width=56, state="readonly",
                 relief="solid", bd=1).pack(side="left", padx=(0, 8))
        tk.Button(frame_in, text="Обзор...", font=("Segoe UI", 9),
                  command=self._browse_input,
                  bg="#4a7fe5", fg="white", activebackground="#3a6fd5",
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")

        frame_out = tk.LabelFrame(p, text=" Папка для сохранения отчёта ",
                                   font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                   padx=10, pady=8)
        frame_out.pack(fill="x", padx=16, pady=4)
        self.output_dir_var = tk.StringVar(value=self._cfg.get("last_output_dir", ""))
        tk.Entry(frame_out, textvariable=self.output_dir_var,
                 font=("Segoe UI", 9), width=56, state="readonly",
                 relief="solid", bd=1).pack(side="left", padx=(0, 8))
        tk.Button(frame_out, text="Обзор...", font=("Segoe UI", 9),
                  command=self._browse_output,
                  bg="#4a7fe5", fg="white", activebackground="#3a6fd5",
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")

        log_frame = tk.LabelFrame(p, text=" Журнал выполнения ",
                                   font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                   padx=10, pady=8)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(8, 4))
        self.log_box = scrolledtext.ScrolledText(
            log_frame, font=("Consolas", 9), height=8, state="disabled",
            bg="#1e1e2e", fg="#cdd6f4", insertbackground="white", relief="flat")
        self.log_box.pack(fill="both", expand=True)

        self.run_btn = tk.Button(
            p, text="▶  Запустить анализ",
            font=("Segoe UI", 11, "bold"), command=self._start_analysis,
            bg="#2ecc71", fg="white", activebackground="#27ae60",
            relief="flat", padx=20, pady=10, cursor="hand2")
        self.run_btn.pack(pady=(8, 14))

    # ============================================================
    # ВКЛАДКА 2: СРАВНЕНИЕ
    # ============================================================

    def _build_tab_compare(self):
        p = self.tab_compare

        tk.Label(p, text="Выберите два файла с выгрузками и задайте метки периодов",
                 font=("Segoe UI", 10), bg="#f0f2f5", fg="#555").pack(pady=(12, 8))

        def make_period_block(parent, label_text, path_var, label_var):
            frame = tk.LabelFrame(parent, text=f" {label_text} ",
                                   font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                   padx=10, pady=8)
            frame.pack(fill="x", padx=16, pady=4)
            tk.Entry(frame, textvariable=path_var,
                     font=("Segoe UI", 9), width=46, state="readonly",
                     relief="solid", bd=1).pack(side="left", padx=(0, 8))
            tk.Button(frame, text="Обзор...", font=("Segoe UI", 9),
                      command=lambda: self._browse_cmp(path_var),
                      bg="#4a7fe5", fg="white", activebackground="#3a6fd5",
                      relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")
            lf = tk.Frame(parent, bg="#f0f2f5")
            lf.pack(fill="x", padx=16, pady=(0, 4))
            tk.Label(lf, text=f"Название {label_text}:", font=("Segoe UI", 9),
                     bg="#f0f2f5", fg="#333").pack(side="left")
            tk.Entry(lf, textvariable=label_var, font=("Segoe UI", 9), width=24,
                     relief="solid", bd=1).pack(side="left", padx=(8, 0))

        self.cmp_path_a  = tk.StringVar()
        self.cmp_label_a = tk.StringVar(value="Период А")
        self.cmp_path_b  = tk.StringVar()
        self.cmp_label_b = tk.StringVar(value="Период Б")
        make_period_block(p, "Период А", self.cmp_path_a, self.cmp_label_a)
        make_period_block(p, "Период Б", self.cmp_path_b, self.cmp_label_b)

        frame_out = tk.LabelFrame(p, text=" Папка для сохранения отчёта ",
                                   font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                   padx=10, pady=8)
        frame_out.pack(fill="x", padx=16, pady=4)
        self.cmp_output_dir = tk.StringVar(value=self._cfg.get("last_output_dir", ""))
        tk.Entry(frame_out, textvariable=self.cmp_output_dir,
                 font=("Segoe UI", 9), width=46, state="readonly",
                 relief="solid", bd=1).pack(side="left", padx=(0, 8))
        tk.Button(frame_out, text="Обзор...", font=("Segoe UI", 9),
                  command=self._browse_cmp_output,
                  bg="#4a7fe5", fg="white", activebackground="#3a6fd5",
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")

        log_frame = tk.LabelFrame(p, text=" Журнал выполнения ",
                                   font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                   padx=10, pady=8)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(6, 4))
        self.cmp_log_box = scrolledtext.ScrolledText(
            log_frame, font=("Consolas", 9), height=5, state="disabled",
            bg="#1e1e2e", fg="#cdd6f4", insertbackground="white", relief="flat")
        self.cmp_log_box.pack(fill="both", expand=True)

        self.cmp_run_btn = tk.Button(
            p, text="▶  Сравнить периоды",
            font=("Segoe UI", 11, "bold"), command=self._start_comparison,
            bg="#e67e22", fg="white", activebackground="#d35400",
            relief="flat", padx=20, pady=10, cursor="hand2")
        self.cmp_run_btn.pack(pady=(6, 14))

    # ============================================================
    # ВКЛАДКА 3: ПЛАН
    # ============================================================

    def _build_tab_plan(self):
        p = self.tab_plan

        tk.Label(p, text="Введите плановые суммы по менеджерам (тыс. руб.)",
                 font=("Segoe UI", 10), bg="#f0f2f5", fg="#555").pack(pady=(12, 4))
        tk.Label(p, text="Оставьте поле пустым, если план не задан. Данные сохраняются автоматически.",
                 font=("Segoe UI", 9), bg="#f0f2f5", fg="#888").pack(pady=(0, 8))

        outer = tk.Frame(p, bg="#f0f2f5")
        outer.pack(fill="both", expand=True, padx=16)

        canvas = tk.Canvas(outer, bg="#f0f2f5", highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self._plan_inner = tk.Frame(canvas, bg="#f0f2f5")
        self._plan_inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._plan_inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        tk.Label(self._plan_inner, text="Менеджер", font=("Segoe UI", 9, "bold"),
                 bg="#f0f2f5", width=30, anchor="w").grid(row=0, column=0, padx=(0,8), pady=2)
        tk.Label(self._plan_inner, text="План, тыс. руб.", font=("Segoe UI", 9, "bold"),
                 bg="#f0f2f5", width=16, anchor="w").grid(row=0, column=1, pady=2)

        saved_plans = self._cfg.get("manager_plan", {})
        default_managers = [
            "Антонова Е.", "Борисова Н.", "Волкова С.",
            "Громова И.", "Данилова К.", "Ефимова О.",
            "Жукова Т.", "Зайцева Л.", "Иванова М.", "Козлова Р.",
        ]
        # Добавляем менеджеров из сохранённого плана, которых нет в списке по умолчанию
        all_managers = list(dict.fromkeys(default_managers + list(saved_plans.keys())))

        for i, name in enumerate(all_managers, start=1):
            self._add_plan_row(name, str(saved_plans.get(name, "")), i)

        add_frame = tk.Frame(p, bg="#f0f2f5")
        add_frame.pack(fill="x", padx=16, pady=(6, 0))
        tk.Label(add_frame, text="Добавить менеджера:", font=("Segoe UI", 9),
                 bg="#f0f2f5").pack(side="left")
        self._new_manager_var = tk.StringVar()
        tk.Entry(add_frame, textvariable=self._new_manager_var,
                 font=("Segoe UI", 9), width=22, relief="solid", bd=1).pack(side="left", padx=8)
        tk.Button(add_frame, text="Добавить", font=("Segoe UI", 9),
                  command=self._add_manager_row_ui,
                  bg="#4a7fe5", fg="white", relief="flat",
                  padx=10, pady=3, cursor="hand2").pack(side="left")

        tk.Label(p, text="💡 Планы применяются автоматически при запуске анализа на вкладке «Анализ»",
                 font=("Segoe UI", 8), bg="#f0f2f5", fg="#aaa").pack(pady=(8, 6))

    def _add_plan_row(self, name: str, saved_val: str = "", row_idx: int = None):
        if row_idx is None:
            row_idx = len(self._plan_vars) + 1
        tk.Label(self._plan_inner, text=name, font=("Segoe UI", 9),
                 bg="#f0f2f5", anchor="w").grid(row=row_idx, column=0, padx=(0,8), pady=2, sticky="w")
        var = tk.StringVar(value=saved_val)
        var.trace_add("write", lambda *_, n=name, v=var: self._save_plan_entry(n, v))
        self._plan_vars[name] = var
        tk.Entry(self._plan_inner, textvariable=var,
                 font=("Segoe UI", 9), width=16, relief="solid", bd=1).grid(row=row_idx, column=1, pady=2)

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
                 text="Папка-наблюдатель автоматически запускает анализ\n"
                      "через 5 минут после появления нового .xlsx файла",
                 font=("Segoe UI", 10), bg="#f0f2f5", fg="#555",
                 justify="center").pack(pady=(12, 10))

        frame_w = tk.LabelFrame(p, text=" Папка для наблюдения ",
                                 font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                 padx=10, pady=8)
        frame_w.pack(fill="x", padx=16, pady=4)
        self.watch_dir_var = tk.StringVar(value=self._cfg.get("watch_dir", ""))
        tk.Entry(frame_w, textvariable=self.watch_dir_var,
                 font=("Segoe UI", 9), width=52, state="readonly",
                 relief="solid", bd=1).pack(side="left", padx=(0, 8))
        tk.Button(frame_w, text="Обзор...", font=("Segoe UI", 9),
                  command=self._browse_watch_dir,
                  bg="#4a7fe5", fg="white", activebackground="#3a6fd5",
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")

        frame_wo = tk.LabelFrame(p, text=" Папка для сохранения отчётов ",
                                  font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                  padx=10, pady=8)
        frame_wo.pack(fill="x", padx=16, pady=4)
        self.watch_out_var = tk.StringVar(value=self._cfg.get("watch_output_dir", ""))
        tk.Entry(frame_wo, textvariable=self.watch_out_var,
                 font=("Segoe UI", 9), width=52, state="readonly",
                 relief="solid", bd=1).pack(side="left", padx=(0, 8))
        tk.Button(frame_wo, text="Обзор...", font=("Segoe UI", 9),
                  command=self._browse_watch_out,
                  bg="#4a7fe5", fg="white", activebackground="#3a6fd5",
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")

        self.watch_status_var = tk.StringVar(value="⏹ Остановлен")
        tk.Label(p, textvariable=self.watch_status_var,
                 font=("Segoe UI", 10, "bold"),
                 bg="#f0f2f5", fg="#555").pack(pady=(8, 4))

        log_frame = tk.LabelFrame(p, text=" Журнал наблюдателя ",
                                   font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                   padx=10, pady=8)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(4, 4))
        self.watch_log_box = scrolledtext.ScrolledText(
            log_frame, font=("Consolas", 9), height=8, state="disabled",
            bg="#1e1e2e", fg="#cdd6f4", insertbackground="white", relief="flat")
        self.watch_log_box.pack(fill="both", expand=True)

        btn_frame = tk.Frame(p, bg="#f0f2f5")
        btn_frame.pack(pady=(6, 14))
        self.watch_start_btn = tk.Button(
            btn_frame, text="▶  Запустить наблюдатель",
            font=("Segoe UI", 10, "bold"), command=self._start_watcher,
            bg="#2ecc71", fg="white", activebackground="#27ae60",
            relief="flat", padx=16, pady=8, cursor="hand2")
        self.watch_start_btn.pack(side="left", padx=(0, 10))
        self.watch_stop_btn = tk.Button(
            btn_frame, text="⏹  Остановить",
            font=("Segoe UI", 10, "bold"), command=self._stop_watcher,
            bg="#e74c3c", fg="white", activebackground="#c0392b",
            relief="flat", padx=16, pady=8, cursor="hand2", state="disabled")
        self.watch_stop_btn.pack(side="left")

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

    def _start_watcher(self):
        watch_dir = self.watch_dir_var.get().strip()
        out_dir   = self.watch_out_var.get().strip()
        if not watch_dir:
            messagebox.showwarning("Внимание", "Укажите папку для наблюдения.")
            return
        if not out_dir:
            messagebox.showwarning("Внимание", "Укажите папку для сохранения отчётов.")
            return
        if not os.path.isdir(watch_dir):
            messagebox.showerror("Ошибка", f"Папка не существует:\n{watch_dir}")
            return
        self._watcher = FolderWatcher(watch_dir, out_dir, self._watch_log)
        self._watcher.start()
        self.watch_status_var.set("👁 Наблюдатель активен")
        self.watch_start_btn.configure(state="disabled")
        self.watch_stop_btn.configure(state="normal")

    def _stop_watcher(self):
        if self._watcher:
            self._watcher.stop()
            self._watcher = None
        self.watch_status_var.set("⏹ Остановлен")
        self.watch_start_btn.configure(state="normal")
        self.watch_stop_btn.configure(state="disabled")

    # ============================================================
    # ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ
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
        path = filedialog.askdirectory(title="Выберите папку для сохранения отчёта")
        if path:
            self.output_dir_var.set(path)
            self._cfg["last_output_dir"] = path
            save_config(self._cfg)

    def _browse_cmp(self, var: tk.StringVar):
        path = filedialog.askopenfilename(
            title="Выберите файл с заказами",
            filetypes=[("Excel файлы", "*.xlsx *.xls")])
        if path:
            var.set(path)
            if not self.cmp_output_dir.get():
                self.cmp_output_dir.set(os.path.dirname(path))

    def _browse_cmp_output(self):
        path = filedialog.askdirectory(title="Выберите папку для сохранения отчёта")
        if path:
            self.cmp_output_dir.set(path)

    def _log(self, message: str):
        self._write_log(self.log_box, message)

    def _cmp_log(self, message: str):
        self._write_log(self.cmp_log_box, message)

    def _watch_log(self, message: str):
        self.after(0, lambda: self._write_log(self.watch_log_box, message))

    def _write_log(self, box: scrolledtext.ScrolledText, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}\n"
        box.configure(state="normal")
        box.insert("end", line)
        box.see("end")
        box.configure(state="disabled")

    # ============================================================
    # ЗАПУСК: АНАЛИЗ
    # ============================================================

    def _start_analysis(self):
        input_path = self.input_var.get().strip()
        output_dir = self.output_dir_var.get().strip()
        if not input_path:
            messagebox.showwarning("Внимание", "Пожалуйста, выберите входной файл с заказами.")
            return
        if not output_dir:
            messagebox.showwarning("Внимание", "Пожалуйста, укажите папку для сохранения отчёта.")
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_dir, f"аналитика_заказов_{timestamp}.xlsx")
        manager_plan = self._get_manager_plan_rub()
        self.run_btn.configure(state="disabled", text="⏳ Выполняется...")
        threading.Thread(
            target=self._run_analysis_thread,
            args=(input_path, output_path, manager_plan),
            daemon=True).start()

    def _run_analysis_thread(self, input_path, output_path, manager_plan):
        try:
            self._log(f"Начинаю анализ файла: {os.path.basename(input_path)}")
            if manager_plan:
                self._log(f"Планы заданы для {len(manager_plan)} менеджеров")
            run_analytics(input_path, output_path, log=self._log, manager_plan=manager_plan)
            self._log(f"Файл отчёта: {output_path}")
            self.after(0, lambda: self._on_success(output_path))
        except Exception as e:
            self._log(f"❌ ОШИБКА: {e}")
            self.after(0, lambda: messagebox.showerror("Ошибка", f"Анализ не выполнен:\n\n{e}"))
        finally:
            self.after(0, lambda: self.run_btn.configure(state="normal", text="▶  Запустить анализ"))

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
            messagebox.showwarning("Внимание", "Укажите папку для сохранения отчёта.")
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_dir, f"сравнение_{timestamp}.xlsx")
        self.cmp_run_btn.configure(state="disabled", text="⏳ Выполняется...")
        threading.Thread(
            target=self._run_comparison_thread,
            args=(path_a, path_b, output_path, label_a, label_b),
            daemon=True).start()

    def _run_comparison_thread(self, path_a, path_b, output_path, label_a, label_b):
        try:
            self._cmp_log(f"Сравниваю: «{label_a}» vs «{label_b}»")
            run_comparison(path_a, path_b, output_path, label_a, label_b, log=self._cmp_log)
            self._cmp_log(f"Файл отчёта: {output_path}")
            self.after(0, lambda: self._on_success(output_path))
        except Exception as e:
            self._cmp_log(f"❌ ОШИБКА: {e}")
            self.after(0, lambda: messagebox.showerror("Ошибка", f"Сравнение не выполнено:\n\n{e}"))
        finally:
            self.after(0, lambda: self.cmp_run_btn.configure(state="normal", text="▶  Сравнить периоды"))

    def _on_success(self, output_path):
        answer = messagebox.askyesno(
            "Готово!",
            f"Отчёт успешно создан.\n\n{output_path}\n\nОткрыть папку с отчётом?")
        if answer:
            os.startfile(os.path.dirname(output_path))


# ============================================================
# ТОЧКА ВХОДА
# ============================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()
