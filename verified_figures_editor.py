# ============================================================
# verified_figures_editor.py — Редактор verified_figures.json
# Для редактирования верифицированных годовых итогов
# ============================================================

import tkinter as tk
from tkinter import ttk, messagebox
import json
import os

C_ORANGE = "#F38120"


class VerifiedFiguresEditor(tk.Toplevel):
    """
    Диалог для редактирования верифицированных цифр бухгалтерии.
    5 полей: общие итоги + программатик + прочие.
    """

    FIELDS = [
        ('total_with_prog',
         'Всего с программатиком и бартером',
         'Верхнеуровневый итог — всё, что заработано. 2025: 482 404 000'),
        ('total_with_barter_no_prog',
         'С бартером, без программатика',
         'Выручка CRM + бартер, но без внешнего программатика. 2025: 363 000 166'),
        ('advertising_no_events',
         'Рекламная без мероприятий',
         'Чистая рекламная выручка. 2025: 243 005 820'),
        ('programmatic_external',
         'Программатик внешний (всего)',
         'Программатик, ведущийся вне CRM. 2025: 119 403 897'),
        ('other_external_income',
         'Прочие внешние доходы',
         'ИРИ, гранты, 47News закупка и т.п. 2025: 41 057 102'),
    ]

    def __init__(self, parent, json_path, theme=None):
        super().__init__(parent)
        self.title("Верифицированные цифры (из бухгалтерии)")
        self.geometry("640x480")
        self.json_path = json_path
        self.theme = theme or {
            'bg': '#F5F5F5', 'surface': '#FFFFFF', 'text': '#1A1A1A',
            'muted': '#777777', 'entry_bg': '#FAFAFA', 'border': '#DEDEDE'
        }
        T = self.theme
        self.configure(bg=T['bg'])

        self.data = {}
        self.entries = {}
        self.year_var = tk.StringVar()

        self._load_data()
        self._build_ui()

        self.transient(parent)
        self.grab_set()
        self.focus_set()

    def _load_data(self):
        if os.path.exists(self.json_path):
            try:
                with open(self.json_path, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
            except Exception:
                self.data = {}

    def _build_ui(self):
        T = self.theme

        # Шапка
        header = tk.Frame(self, bg=C_ORANGE, height=50)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)
        tk.Label(header, text="Верифицированные цифры",
                 bg=C_ORANGE, fg='white',
                 font=('Segoe UI', 13, 'bold')).pack(side='left', padx=20, pady=12)

        # Нотис
        notice_frame = tk.Frame(self, bg=T['bg'])
        notice_frame.pack(fill='x', padx=15, pady=(10, 8))
        tk.Label(notice_frame,
                 text="Годовые итоги из отчёта бухгалтерии. Суммы в РУБЛЯХ.\n"
                      "Обновляются раз в год или при значимых корректировках.",
                 bg=T['bg'], fg=T['muted'],
                 font=('Segoe UI', 9),
                 justify='left').pack(side='left')

        # Год
        year_frame = tk.Frame(self, bg=T['surface'])
        year_frame.pack(fill='x', padx=15, pady=5)
        tk.Label(year_frame, text="Отчётный год:",
                 bg=T['surface'], fg=T['text'],
                 font=('Segoe UI', 10, 'bold'),
                 padx=10, pady=8).pack(side='left')
        self.year_var.set(str(self.data.get('_год', 2025)))
        tk.Entry(year_frame, textvariable=self.year_var,
                 width=8, justify='center',
                 bg=T['entry_bg'], fg=T['text'],
                 relief='flat', borderwidth=1,
                 font=('Segoe UI', 10)).pack(side='left', padx=5)

        # Поля
        fields_frame = tk.Frame(self, bg=T['bg'])
        fields_frame.pack(fill='both', expand=True, padx=15, pady=10)

        for i, (key, label, hint) in enumerate(self.FIELDS):
            card = tk.Frame(fields_frame, bg=T['surface'])
            card.pack(fill='x', pady=4)

            # Левая колонка: название и подсказка
            left = tk.Frame(card, bg=T['surface'])
            left.pack(side='left', fill='x', expand=True, padx=10, pady=8)
            tk.Label(left, text=label,
                     bg=T['surface'], fg=T['text'],
                     font=('Segoe UI', 10, 'bold'),
                     anchor='w').pack(fill='x')
            tk.Label(left, text=hint,
                     bg=T['surface'], fg=T['muted'],
                     font=('Segoe UI', 8),
                     anchor='w').pack(fill='x')

            # Правая колонка: поле ввода
            value = self.data.get(key, 0)
            var = tk.StringVar(value=self._fmt(value))
            entry = tk.Entry(card, textvariable=var,
                             width=18, justify='right',
                             bg=T['entry_bg'], fg=T['text'],
                             relief='flat', borderwidth=1,
                             font=('Segoe UI', 10))
            entry.pack(side='right', padx=10, pady=8)
            self.entries[key] = var

        # Кнопки
        btn_frame = tk.Frame(self, bg=T['bg'])
        btn_frame.pack(fill='x', padx=15, pady=(5, 15))

        tk.Button(btn_frame, text="Отмена",
                  command=self.destroy,
                  bg=T['surface'], fg=T['text'],
                  relief='flat', borderwidth=1,
                  font=('Segoe UI', 10),
                  padx=20, pady=8,
                  cursor='hand2').pack(side='right', padx=(8, 0))

        tk.Button(btn_frame, text="Сохранить",
                  command=self._save,
                  bg=C_ORANGE, fg='white',
                  activebackground='#D06A10', activeforeground='white',
                  relief='flat', borderwidth=0,
                  font=('Segoe UI', 10, 'bold'),
                  padx=20, pady=8,
                  cursor='hand2').pack(side='right')

        self.bind('<Return>', lambda e: self._save())
        self.bind('<Escape>', lambda e: self.destroy())

    def _parse(self, s):
        if not s:
            return 0
        s = str(s).strip().replace(' ', '').replace(',', '.')
        try:
            return int(float(s))
        except Exception:
            return 0

    def _fmt(self, v):
        try:
            v = float(v)
        except Exception:
            return '0'
        return f"{int(v):,}".replace(',', ' ')

    def _save(self):
        new_data = dict(self.data)
        try:
            new_data['_год'] = int(self.year_var.get())
        except Exception:
            new_data['_год'] = 2025
        new_data.setdefault('_описание',
            "Верифицированные годовые итоги из бухгалтерии.")

        for key, var in self.entries.items():
            new_data[key] = self._parse(var.get())

        try:
            with open(self.json_path, 'w', encoding='utf-8') as f:
                json.dump(new_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))
            return

        messagebox.showinfo("Сохранено",
            "Верифицированные цифры обновлены.\n"
            "Запусти аналитику заново для пересчёта.")
        self.destroy()


if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     'verified_figures.json')
    VerifiedFiguresEditor(root, p)
    root.mainloop()
