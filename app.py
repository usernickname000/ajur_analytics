# ============================================================
# app.py — графический интерфейс утилиты аналитики заказов
# Запуск: python app.py
# ============================================================

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import os
from datetime import datetime

# Импортируем нашу аналитику
from analytics import run_analytics


# ============================================================
# ГЛАВНОЕ ОКНО
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Аналитика заказов — Коммерческий отдел")
        self.geometry("680x500")
        self.resizable(False, False)
        self.configure(bg="#f0f2f5")

        self._build_ui()

    # ── Построение интерфейса ────────────────────────────────
    def _build_ui(self):

        # Заголовок
        header = tk.Label(
            self, text="📊 Аналитика заказов",
            font=("Segoe UI", 16, "bold"),
            bg="#f0f2f5", fg="#1a1a2e"
        )
        header.pack(pady=(20, 4))

        subtitle = tk.Label(
            self, text="Выберите файл Excel с заказами и нажмите «Запустить»",
            font=("Segoe UI", 10),
            bg="#f0f2f5", fg="#555"
        )
        subtitle.pack(pady=(0, 16))

        # ── Блок выбора входного файла ───────────────────────
        frame_in = tk.LabelFrame(
            self, text=" Входной файл (заказы) ",
            font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
            padx=10, pady=8
        )
        frame_in.pack(fill="x", padx=24, pady=4)

        self.input_var = tk.StringVar()

        tk.Entry(
            frame_in, textvariable=self.input_var,
            font=("Segoe UI", 9), width=58, state="readonly",
            relief="solid", bd=1
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            frame_in, text="Обзор...",
            font=("Segoe UI", 9),
            command=self._browse_input,
            bg="#4a7fe5", fg="white",
            activebackground="#3a6fd5",
            relief="flat", padx=10, pady=4, cursor="hand2"
        ).pack(side="left")

        # ── Блок выбора папки для отчёта ─────────────────────
        frame_out = tk.LabelFrame(
            self, text=" Папка для сохранения отчёта ",
            font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
            padx=10, pady=8
        )
        frame_out.pack(fill="x", padx=24, pady=4)

        self.output_dir_var = tk.StringVar()

        tk.Entry(
            frame_out, textvariable=self.output_dir_var,
            font=("Segoe UI", 9), width=58, state="readonly",
            relief="solid", bd=1
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            frame_out, text="Обзор...",
            font=("Segoe UI", 9),
            command=self._browse_output,
            bg="#4a7fe5", fg="white",
            activebackground="#3a6fd5",
            relief="flat", padx=10, pady=4, cursor="hand2"
        ).pack(side="left")

        # ── Лог выполнения ───────────────────────────────────
        log_frame = tk.LabelFrame(
            self, text=" Журнал выполнения ",
            font=("Segoe UI", 9), bg="#f0f2f5", fg="#333",
            padx=10, pady=8
        )
        log_frame.pack(fill="both", expand=True, padx=24, pady=(8, 4))

        self.log_box = scrolledtext.ScrolledText(
            log_frame, font=("Consolas", 9),
            height=10, state="disabled",
            bg="#1e1e2e", fg="#cdd6f4",
            insertbackground="white",
            relief="flat"
        )
        self.log_box.pack(fill="both", expand=True)

        # ── Кнопка запуска ───────────────────────────────────
        self.run_btn = tk.Button(
            self, text="▶  Запустить анализ",
            font=("Segoe UI", 11, "bold"),
            command=self._start_analysis,
            bg="#2ecc71", fg="white",
            activebackground="#27ae60",
            relief="flat", padx=20, pady=10,
            cursor="hand2"
        )
        self.run_btn.pack(pady=(8, 20))

    # ── Выбор файлов ─────────────────────────────────────────
    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Выберите файл с заказами",
            filetypes=[("Excel файлы", "*.xlsx *.xls")]
        )
        if path:
            self.input_var.set(path)
            # Автоматически ставим ту же папку для отчёта
            if not self.output_dir_var.get():
                self.output_dir_var.set(os.path.dirname(path))

    def _browse_output(self):
        path = filedialog.askdirectory(title="Выберите папку для сохранения отчёта")
        if path:
            self.output_dir_var.set(path)

    # ── Логирование ──────────────────────────────────────────
    def _log(self, message: str):
        """Добавляет строку в журнал (потокобезопасно)."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}\n"

        self.log_box.configure(state="normal")
        self.log_box.insert("end", line)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    # ── Запуск анализа ───────────────────────────────────────
    def _start_analysis(self):
        input_path = self.input_var.get().strip()
        output_dir = self.output_dir_var.get().strip()

        if not input_path:
            messagebox.showwarning("Внимание", "Пожалуйста, выберите входной файл с заказами.")
            return

        if not output_dir:
            messagebox.showwarning("Внимание", "Пожалуйста, укажите папку для сохранения отчёта.")
            return

        # Формируем имя выходного файла с датой и временем
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_dir, f"аналитика_заказов_{timestamp}.xlsx")

        # Блокируем кнопку на время работы
        self.run_btn.configure(state="disabled", text="⏳ Выполняется...")

        # Запускаем в отдельном потоке, чтобы GUI не зависал
        thread = threading.Thread(
            target=self._run_in_thread,
            args=(input_path, output_path),
            daemon=True
        )
        thread.start()

    def _run_in_thread(self, input_path, output_path):
        try:
            self._log(f"Начинаю анализ файла: {os.path.basename(input_path)}")
            run_analytics(input_path, output_path, log=self._log)
            self._log(f"Файл отчёта: {output_path}")

            # Спрашиваем — открыть папку?
            self.after(0, lambda: self._on_success(output_path))

        except Exception as e:
            self._log(f"❌ ОШИБКА: {e}")
            self.after(0, lambda: messagebox.showerror(
                "Ошибка", f"Анализ не выполнен:\n\n{e}"
            ))
        finally:
            self.after(0, lambda: self.run_btn.configure(
                state="normal", text="▶  Запустить анализ"
            ))

    def _on_success(self, output_path):
        answer = messagebox.askyesno(
            "Готово!",
            f"Отчёт успешно создан.\n\n{output_path}\n\nОткрыть папку с отчётом?"
        )
        if answer:
            folder = os.path.dirname(output_path)
            os.startfile(folder)  # Работает на Windows


# ============================================================
# ТОЧКА ВХОДА
# ============================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()
