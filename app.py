# ============================================================
# app.py — графический интерфейс утилиты аналитики заказов
# Запуск: python app.py
# ============================================================

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import os
from datetime import datetime

from analytics import run_analytics
from comparison import run_comparison


# ============================================================
# ГЛАВНОЕ ОКНО
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Аналитика заказов — Коммерческий отдел")
        self.geometry("700x580")
        self.resizable(False, False)
        self.configure(bg="#f0f2f5")

        self._build_ui()

    def _build_ui(self):
        # Заголовок
        tk.Label(
            self, text="📊 Аналитика заказов",
            font=("Segoe UI", 16, "bold"),
            bg="#f0f2f5", fg="#1a1a2e"
        ).pack(pady=(16, 4))

        # Вкладки
        style = ttk.Style(self)
        style.configure("TNotebook", background="#f0f2f5", borderwidth=0)
        style.configure("TNotebook.Tab", font=("Segoe UI", 10), padding=[14, 6])

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=16, pady=(4, 0))

        self.tab_main = tk.Frame(self.notebook, bg="#f0f2f5")
        self.tab_compare = tk.Frame(self.notebook, bg="#f0f2f5")

        self.notebook.add(self.tab_main,    text="  Анализ периода  ")
        self.notebook.add(self.tab_compare, text="  Сравнение периодов  ")

        self._build_tab_main()
        self._build_tab_compare()

    # ============================================================
    # ВКЛАДКА 1: АНАЛИЗ ОДНОГО ПЕРИОДА (прежний интерфейс)
    # ============================================================

    def _build_tab_main(self):
        p = self.tab_main

        tk.Label(
            p, text="Выберите файл Excel с заказами и нажмите «Запустить»",
            font=("Segoe UI", 10), bg="#f0f2f5", fg="#555"
        ).pack(pady=(12, 8))

        # Входной файл
        frame_in = tk.LabelFrame(p, text=" Входной файл (заказы) ",
                                  font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                  padx=10, pady=8)
        frame_in.pack(fill="x", padx=16, pady=4)

        self.input_var = tk.StringVar()
        tk.Entry(frame_in, textvariable=self.input_var,
                 font=("Segoe UI", 9), width=56, state="readonly",
                 relief="solid", bd=1).pack(side="left", padx=(0, 8))
        tk.Button(frame_in, text="Обзор...", font=("Segoe UI", 9),
                  command=self._browse_input,
                  bg="#4a7fe5", fg="white", activebackground="#3a6fd5",
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")

        # Папка для отчёта
        frame_out = tk.LabelFrame(p, text=" Папка для сохранения отчёта ",
                                   font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                   padx=10, pady=8)
        frame_out.pack(fill="x", padx=16, pady=4)

        self.output_dir_var = tk.StringVar()
        tk.Entry(frame_out, textvariable=self.output_dir_var,
                 font=("Segoe UI", 9), width=56, state="readonly",
                 relief="solid", bd=1).pack(side="left", padx=(0, 8))
        tk.Button(frame_out, text="Обзор...", font=("Segoe UI", 9),
                  command=self._browse_output,
                  bg="#4a7fe5", fg="white", activebackground="#3a6fd5",
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")

        # Лог
        log_frame = tk.LabelFrame(p, text=" Журнал выполнения ",
                                   font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                   padx=10, pady=8)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(8, 4))

        self.log_box = scrolledtext.ScrolledText(
            log_frame, font=("Consolas", 9), height=8, state="disabled",
            bg="#1e1e2e", fg="#cdd6f4", insertbackground="white", relief="flat"
        )
        self.log_box.pack(fill="both", expand=True)

        # Кнопка
        self.run_btn = tk.Button(
            p, text="▶  Запустить анализ",
            font=("Segoe UI", 11, "bold"),
            command=self._start_analysis,
            bg="#2ecc71", fg="white", activebackground="#27ae60",
            relief="flat", padx=20, pady=10, cursor="hand2"
        )
        self.run_btn.pack(pady=(8, 16))

    # ============================================================
    # ВКЛАДКА 2: СРАВНЕНИЕ ДВУХ ПЕРИОДОВ
    # ============================================================

    def _build_tab_compare(self):
        p = self.tab_compare

        tk.Label(
            p, text="Выберите два файла с выгрузками и задайте метки периодов",
            font=("Segoe UI", 10), bg="#f0f2f5", fg="#555"
        ).pack(pady=(12, 8))

        # Период А
        frame_a = tk.LabelFrame(p, text=" Период А ",
                                  font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                  padx=10, pady=8)
        frame_a.pack(fill="x", padx=16, pady=4)

        self.cmp_path_a = tk.StringVar()
        tk.Entry(frame_a, textvariable=self.cmp_path_a,
                 font=("Segoe UI", 9), width=46, state="readonly",
                 relief="solid", bd=1).pack(side="left", padx=(0, 8))
        tk.Button(frame_a, text="Обзор...", font=("Segoe UI", 9),
                  command=lambda: self._browse_cmp(self.cmp_path_a),
                  bg="#4a7fe5", fg="white", activebackground="#3a6fd5",
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")

        # Метка периода А
        lbl_a_frame = tk.Frame(p, bg="#f0f2f5")
        lbl_a_frame.pack(fill="x", padx=16, pady=(0, 4))
        tk.Label(lbl_a_frame, text="Название периода А:", font=("Segoe UI", 9),
                 bg="#f0f2f5", fg="#333").pack(side="left")
        self.cmp_label_a = tk.StringVar(value="Период А")
        tk.Entry(lbl_a_frame, textvariable=self.cmp_label_a,
                 font=("Segoe UI", 9), width=24,
                 relief="solid", bd=1).pack(side="left", padx=(8, 0))

        # Период Б
        frame_b = tk.LabelFrame(p, text=" Период Б ",
                                  font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                  padx=10, pady=8)
        frame_b.pack(fill="x", padx=16, pady=4)

        self.cmp_path_b = tk.StringVar()
        tk.Entry(frame_b, textvariable=self.cmp_path_b,
                 font=("Segoe UI", 9), width=46, state="readonly",
                 relief="solid", bd=1).pack(side="left", padx=(0, 8))
        tk.Button(frame_b, text="Обзор...", font=("Segoe UI", 9),
                  command=lambda: self._browse_cmp(self.cmp_path_b),
                  bg="#4a7fe5", fg="white", activebackground="#3a6fd5",
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")

        # Метка периода Б
        lbl_b_frame = tk.Frame(p, bg="#f0f2f5")
        lbl_b_frame.pack(fill="x", padx=16, pady=(0, 4))
        tk.Label(lbl_b_frame, text="Название периода Б:", font=("Segoe UI", 9),
                 bg="#f0f2f5", fg="#333").pack(side="left")
        self.cmp_label_b = tk.StringVar(value="Период Б")
        tk.Entry(lbl_b_frame, textvariable=self.cmp_label_b,
                 font=("Segoe UI", 9), width=24,
                 relief="solid", bd=1).pack(side="left", padx=(8, 0))

        # Папка для отчёта
        frame_out = tk.LabelFrame(p, text=" Папка для сохранения отчёта ",
                                   font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                   padx=10, pady=8)
        frame_out.pack(fill="x", padx=16, pady=4)

        self.cmp_output_dir = tk.StringVar()
        tk.Entry(frame_out, textvariable=self.cmp_output_dir,
                 font=("Segoe UI", 9), width=46, state="readonly",
                 relief="solid", bd=1).pack(side="left", padx=(0, 8))
        tk.Button(frame_out, text="Обзор...", font=("Segoe UI", 9),
                  command=self._browse_cmp_output,
                  bg="#4a7fe5", fg="white", activebackground="#3a6fd5",
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="left")

        # Лог
        log_frame = tk.LabelFrame(p, text=" Журнал выполнения ",
                                   font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
                                   padx=10, pady=8)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(6, 4))

        self.cmp_log_box = scrolledtext.ScrolledText(
            log_frame, font=("Consolas", 9), height=5, state="disabled",
            bg="#1e1e2e", fg="#cdd6f4", insertbackground="white", relief="flat"
        )
        self.cmp_log_box.pack(fill="both", expand=True)

        # Кнопка
        self.cmp_run_btn = tk.Button(
            p, text="▶  Сравнить периоды",
            font=("Segoe UI", 11, "bold"),
            command=self._start_comparison,
            bg="#e67e22", fg="white", activebackground="#d35400",
            relief="flat", padx=20, pady=10, cursor="hand2"
        )
        self.cmp_run_btn.pack(pady=(6, 14))

    # ============================================================
    # ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ
    # ============================================================

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Выберите файл с заказами",
            filetypes=[("Excel файлы", "*.xlsx *.xls")]
        )
        if path:
            self.input_var.set(path)
            if not self.output_dir_var.get():
                self.output_dir_var.set(os.path.dirname(path))

    def _browse_output(self):
        path = filedialog.askdirectory(title="Выберите папку для сохранения отчёта")
        if path:
            self.output_dir_var.set(path)

    def _browse_cmp(self, var: tk.StringVar):
        path = filedialog.askopenfilename(
            title="Выберите файл с заказами",
            filetypes=[("Excel файлы", "*.xlsx *.xls")]
        )
        if path:
            var.set(path)
            if not self.cmp_output_dir.get():
                self.cmp_output_dir.set(os.path.dirname(path))

    def _browse_cmp_output(self):
        path = filedialog.askdirectory(title="Выберите папку для сохранения отчёта")
        if path:
            self.cmp_output_dir.set(path)

    # ── Логирование ──────────────────────────────────────────

    def _log(self, message: str):
        self._write_log(self.log_box, message)

    def _cmp_log(self, message: str):
        self._write_log(self.cmp_log_box, message)

    def _write_log(self, box: scrolledtext.ScrolledText, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}\n"
        box.configure(state="normal")
        box.insert("end", line)
        box.see("end")
        box.configure(state="disabled")

    # ============================================================
    # ЗАПУСК: АНАЛИЗ ОДНОГО ПЕРИОДА
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

        self.run_btn.configure(state="disabled", text="⏳ Выполняется...")
        threading.Thread(
            target=self._run_analysis_thread,
            args=(input_path, output_path),
            daemon=True
        ).start()

    def _run_analysis_thread(self, input_path, output_path):
        try:
            self._log(f"Начинаю анализ файла: {os.path.basename(input_path)}")
            run_analytics(input_path, output_path, log=self._log)
            self._log(f"Файл отчёта: {output_path}")
            self.after(0, lambda: self._on_success(output_path, self.run_btn, "▶  Запустить анализ"))
        except Exception as e:
            self._log(f"❌ ОШИБКА: {e}")
            self.after(0, lambda: messagebox.showerror("Ошибка", f"Анализ не выполнен:\n\n{e}"))
        finally:
            self.after(0, lambda: self.run_btn.configure(state="normal", text="▶  Запустить анализ"))

    # ============================================================
    # ЗАПУСК: СРАВНЕНИЕ ДВУХ ПЕРИОДОВ
    # ============================================================

    def _start_comparison(self):
        path_a    = self.cmp_path_a.get().strip()
        path_b    = self.cmp_path_b.get().strip()
        output_dir = self.cmp_output_dir.get().strip()
        label_a   = self.cmp_label_a.get().strip() or "Период А"
        label_b   = self.cmp_label_b.get().strip() or "Период Б"

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
            daemon=True
        ).start()

    def _run_comparison_thread(self, path_a, path_b, output_path, label_a, label_b):
        try:
            self._cmp_log(f"Сравниваю: «{label_a}» vs «{label_b}»")
            run_comparison(path_a, path_b, output_path, label_a, label_b, log=self._cmp_log)
            self._cmp_log(f"Файл отчёта: {output_path}")
            self.after(0, lambda: self._on_success(output_path, self.cmp_run_btn, "▶  Сравнить периоды"))
        except Exception as e:
            self._cmp_log(f"❌ ОШИБКА: {e}")
            self.after(0, lambda: messagebox.showerror("Ошибка", f"Сравнение не выполнено:\n\n{e}"))
        finally:
            self.after(0, lambda: self.cmp_run_btn.configure(state="normal", text="▶  Сравнить периоды"))

    # ── Диалог после успеха ──────────────────────────────────

    def _on_success(self, output_path, btn, btn_text):
        answer = messagebox.askyesno(
            "Готово!",
            f"Отчёт успешно создан.\n\n{output_path}\n\nОткрыть папку с отчётом?"
        )
        if answer:
            os.startfile(os.path.dirname(output_path))


# ============================================================
# ТОЧКА ВХОДА
# ============================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()
