# ============================================================
# dohody_import_dialog.py — GUI для импорта "Доходы ГГГГ.xlsx"
# Открывается как модальное окно из основного app.py.
# Логика разбора/записи — в dohody_import.py, здесь только GUI.
# ============================================================

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import dohody_import as di
import ui_theme as ui
from ui_theme import ICON, T, px


class DohodyImportDialog(ui.FluentDialog):
    """
    Выбрать файл → посмотреть отчёт (что распознано, что нет, где
    расхождения/выбросы) → применить (запись в external_income.json).
    Ничего не пишется без явного нажатия «Применить».
    """

    def __init__(self, parent, external_json_path, **_):
        self.external_json_path = external_json_path
        self.parsed = None
        self.xlsx_path = None
        super().__init__(parent, "Импорт «Доходы ГГГГ.xlsx»",
                         "Статьи, которых нет в CRM, будут записаны в external_income.json. "
                         "Статьи из CRM (реклама, IC-доходы) только показываются для сверки.",
                         size=(900, 660), minsize=(700, 480))
        self._build_ui()
        self.show_modal(parent)

    def _build_ui(self):
        pick = tk.Frame(self.body, bg=T['bg'])
        pick.pack(fill='x', pady=(0, px(12)))
        self.path_var = tk.StringVar(value="Файл не выбран")
        ttk.Entry(pick, textvariable=self.path_var, state='readonly').pack(side='left', fill='x', expand=True)
        ui.button(pick, "Выбрать файл…", self._pick_file, on='bg', icon=ICON['file']).pack(
            side='left', padx=(px(8), 0))

        outer, body = ui.card(self.body, pad=8)
        outer.pack(fill='both', expand=True)
        self.report_box = tk.Text(body, wrap='word', font=ui.FONTS['mono'], relief='flat', bd=0,
                                  bg=T['card'], fg=T['text'], padx=px(6), pady=px(4))
        sb = ttk.Scrollbar(body, orient='vertical', command=self.report_box.yview,
                           style='Thin.Vertical.TScrollbar')
        self.report_box.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.report_box.pack(side='left', fill='both', expand=True)
        self._set_report("Выберите файл «Доходы ГГГГ.xlsx», чтобы увидеть отчёт.")

        self.hint.configure(text="Сначала отчёт, запись — только по «Применить»")
        self.add_button("Закрыть", self.destroy)
        self.apply_btn = self.add_button("Применить", self._apply, kind="Accent")
        self.apply_btn.state(['disabled'])

    def _set_report(self, text):
        self.report_box.configure(state='normal')
        self.report_box.delete('1.0', 'end')
        self.report_box.insert('1.0', text)
        self.report_box.configure(state='disabled')

    def _pick_file(self):
        path = filedialog.askopenfilename(
            parent=self, title="Выберите файл Доходы ГГГГ.xlsx",
            filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")])
        if not path:
            return
        self.xlsx_path = path
        self.path_var.set(path)
        try:
            self.parsed = di.parse_workbook(path)
        except Exception as e:
            self.parsed = None
            self.apply_btn.state(['disabled'])
            messagebox.showerror("Ошибка разбора файла", str(e), parent=self)
            self._set_report(f"Не удалось разобрать файл:\n{e}")
            return
        self._set_report(di.generate_report(self.parsed))
        self.apply_btn.state(['!disabled'])

    def _apply(self):
        if self.parsed is None:
            return
        year = self.parsed.year
        if not messagebox.askyesno(
            "Подтверждение",
            f"Записать статьи за {year} год в external_income.json?\n\n"
            "Если сейчас в файле другой отчётный год — он будет сохранён "
            f"как резервная копия, а активным станет {year} год "
            "(сработает при следующем запуске анализа).", parent=self):
            return
        try:
            report = di.apply_to_external_income(self.external_json_path, self.parsed)
        except Exception as e:
            messagebox.showerror("Ошибка записи", str(e), parent=self)
            return
        messagebox.showinfo("Готово", "external_income.json обновлён:\n\n" + report, parent=self)
        self._set_report(di.generate_report(self.parsed) + "\n\n── Применено ──\n" + report)
        self.apply_btn.state(['disabled'])


if __name__ == '__main__':
    root = tk.Tk()
    ui.init(root)
    root.withdraw()
    DohodyImportDialog(root, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'external_income.json'))
    root.mainloop()
