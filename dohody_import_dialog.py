# ============================================================
# dohody_import_dialog.py — GUI для импорта "Доходы ГГГГ.xlsx"
# Открывается как модальное окно из основного app.py.
# Логика разбора/записи — в dohody_import.py, здесь только GUI.
# ============================================================

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext

import dohody_import as di

C_ORANGE = "#F38120"
C_GREEN  = "#22C55E"
C_RED    = "#EF4444"


class DohodyImportDialog(tk.Toplevel):
    """
    Диалог импорта мастер-файла бухгалтерии "Доходы ГГГГ.xlsx":
    выбрать файл → посмотреть отчёт (что распознано, что нет, где
    расхождения/выбросы) → применить (запись в external_income.json).
    Ничего не пишется без явного нажатия "Применить".
    """

    def __init__(self, parent, external_json_path, theme=None):
        super().__init__(parent)
        self.title("Импорт Доходы.xlsx")
        self.geometry("880x640")
        self.external_json_path = external_json_path
        self.theme = theme or {
            'bg': '#F5F5F5', 'surface': '#FFFFFF', 'text': '#1A1A1A',
            'muted': '#777777', 'entry_bg': '#FAFAFA', 'border': '#DEDEDE'
        }
        self.parsed = None
        self.xlsx_path = None

        self._build_ui()

        self.transient(parent)
        self.grab_set()
        self.focus_set()

    def _build_ui(self):
        T = self.theme
        self.configure(bg=T['bg'])

        header = tk.Frame(self, bg=C_ORANGE, height=50)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)
        tk.Label(header, text="Импорт Доходы ГГГГ.xlsx",
                 bg=C_ORANGE, fg='white',
                 font=('Segoe UI', 13, 'bold')).pack(side='left', padx=20, pady=12)

        notice = tk.Frame(self, bg=T['bg'])
        notice.pack(fill='x', padx=15, pady=(10, 4))
        tk.Label(notice,
                 text="Читает мастер-файл бухгалтерии и готовит обновление внешних доходов.\n"
                      "Строки, которых нет в CRM, будут записаны в external_income.json.\n"
                      "Строки, которые есть в CRM (реклама, IC-доходы и т.п.), только показываются "
                      "для сверки — не записываются, их источник — выгрузка CRM.",
                 bg=T['bg'], fg=T['muted'], font=('Segoe UI', 9),
                 justify='left').pack(anchor='w')

        pick_row = tk.Frame(self, bg=T['bg'])
        pick_row.pack(fill='x', padx=15, pady=(8, 4))
        self.path_var = tk.StringVar(value="Файл не выбран")
        tk.Label(pick_row, textvariable=self.path_var,
                 bg=T['surface'], fg=T['text'], anchor='w',
                 font=('Segoe UI', 9), padx=8, pady=6).pack(side='left', fill='x', expand=True)
        tk.Button(pick_row, text="Выбрать файл…", command=self._pick_file,
                  bg=T['surface'], fg=T['text'], relief='flat', borderwidth=1,
                  font=('Segoe UI', 9), padx=12, pady=6, cursor='hand2').pack(side='left', padx=(8, 0))

        report_frame = tk.Frame(self, bg=T['bg'])
        report_frame.pack(fill='both', expand=True, padx=15, pady=(8, 4))
        self.report_box = scrolledtext.ScrolledText(
            report_frame, wrap='word', font=('Consolas', 9),
            bg=T['entry_bg'], fg=T['text'], relief='flat', borderwidth=1)
        self.report_box.pack(fill='both', expand=True)
        self.report_box.insert('1.0', "Выбери файл «Доходы ГГГГ.xlsx», чтобы увидеть отчёт.")
        self.report_box.configure(state='disabled')

        btn_frame = tk.Frame(self, bg=T['bg'])
        btn_frame.pack(fill='x', padx=15, pady=(4, 15))

        tk.Button(btn_frame, text="Закрыть", command=self.destroy,
                  bg=T['surface'], fg=T['text'], relief='flat', borderwidth=1,
                  font=('Segoe UI', 10), padx=20, pady=8, cursor='hand2').pack(side='right', padx=(8, 0))

        self.apply_btn = tk.Button(
            btn_frame, text="Применить (записать в external_income.json)",
            command=self._apply, state='disabled',
            bg=C_ORANGE, fg='white', activebackground='#D06A10', activeforeground='white',
            relief='flat', borderwidth=0, font=('Segoe UI', 10, 'bold'),
            padx=20, pady=8, cursor='hand2')
        self.apply_btn.pack(side='right')

    def _set_report(self, text):
        self.report_box.configure(state='normal')
        self.report_box.delete('1.0', 'end')
        self.report_box.insert('1.0', text)
        self.report_box.configure(state='disabled')

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="Выбери файл Доходы ГГГГ.xlsx",
            filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")])
        if not path:
            return
        self.xlsx_path = path
        self.path_var.set(os.path.basename(path))
        try:
            self.parsed = di.parse_workbook(path)
        except Exception as e:
            self.parsed = None
            self.apply_btn.configure(state='disabled')
            messagebox.showerror("Ошибка разбора файла", str(e))
            self._set_report(f"Не удалось разобрать файл:\n{e}")
            return

        self._set_report(di.generate_report(self.parsed))
        self.apply_btn.configure(state='normal')

    def _apply(self):
        if self.parsed is None:
            return
        year = self.parsed.year
        confirm = messagebox.askyesno(
            "Подтверждение",
            f"Записать статьи за {year} год в external_income.json?\n\n"
            "Если сейчас в файле другой отчётный год — он будет сохранён "
            "как резервная копия, а активным станет "
            f"{year} год (сработает при следующем запуске анализа).",
        )
        if not confirm:
            return
        try:
            report = di.apply_to_external_income(self.external_json_path, self.parsed)
        except Exception as e:
            messagebox.showerror("Ошибка записи", str(e))
            return

        messagebox.showinfo("Готово", "external_income.json обновлён:\n\n" + report)
        self._set_report(di.generate_report(self.parsed) + "\n\n── Применено ──\n" + report)
        self.apply_btn.configure(state='disabled')


if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'external_income.json')
    DohodyImportDialog(root, p)
    root.mainloop()
