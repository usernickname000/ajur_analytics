# ============================================================
# external_income_editor.py — Редактор external_income.json
# Открывается как модальное окно из основного app.py
# ============================================================

import tkinter as tk
from tkinter import ttk, messagebox
import json
import os

# Цвета — те же что в app.py (светлая/тёмная тема)
C_ORANGE  = "#F38120"
C_BORDER  = "#DEDEDE"

MONTHS = [f"{m:02d}" for m in range(1, 13)]
MONTH_LABELS = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн',
                'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек']


class ExternalIncomeEditor(tk.Toplevel):
    """
    Модальное окно для редактирования external_income.json.
    Показывает таблицу: строки = статьи доходов, колонки = 12 месяцев.
    Позволяет редактировать суммы, переключать "включать в аналитику",
    сохраняет обратно в JSON.
    """

    def __init__(self, parent, json_path, theme=None):
        super().__init__(parent)
        self.title("Редактор внешних доходов")
        self.geometry("1150x650")
        self.json_path = json_path
        self.theme = theme or {
            'bg': '#F5F5F5', 'surface': '#FFFFFF', 'text': '#1A1A1A',
            'muted': '#777777', 'entry_bg': '#FAFAFA', 'border': C_BORDER
        }
        T = self.theme
        self.configure(bg=T['bg'])

        # Хранение данных
        self.data = {}
        self.entries = {}          # {(row_name, month): Entry widget}
        self.include_vars = {}     # {row_name: BooleanVar}
        self.plan_entries = {}     # {month: Entry widget}

        self._load_data()
        self._build_ui()

        # Модальность
        self.transient(parent)
        self.grab_set()
        self.focus_set()

    # ── Загрузка JSON ────────────────────────────────────────
    def _load_data(self):
        if not os.path.exists(self.json_path):
            messagebox.showerror("Ошибка",
                f"Файл не найден:\n{self.json_path}\n\n"
                "Создай его вручную или убедись что он рядом с analytics.py")
            self.destroy()
            return
        try:
            with open(self.json_path, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
        except Exception as e:
            messagebox.showerror("Ошибка чтения JSON",
                f"Не удалось прочитать файл:\n{e}")
            self.destroy()
            return

    # ── Построение интерфейса ────────────────────────────────
    def _build_ui(self):
        T = self.theme

        # Шапка
        header = tk.Frame(self, bg=C_ORANGE, height=50)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)
        tk.Label(header, text="Редактор внешних доходов",
                 bg=C_ORANGE, fg='white',
                 font=('Segoe UI', 13, 'bold')).pack(side='left', padx=20, pady=12)

        year = self.data.get('_год', 2025)
        tk.Label(header, text=f"Год: {year}",
                 bg=C_ORANGE, fg='white',
                 font=('Segoe UI', 10)).pack(side='right', padx=20)

        # Нотис
        notice = tk.Frame(self, bg=T['bg'])
        notice.pack(fill='x', padx=15, pady=(8, 4))
        tk.Label(notice,
                 text="Суммы в рублях. Столбец «В аналитике» — если галочка, "
                      "строка суммируется с CRM-выручкой. Обычно это только программатик.",
                 bg=T['bg'], fg=T['muted'],
                 font=('Segoe UI', 9)).pack(side='left')

        # Контейнер с прокруткой
        main = tk.Frame(self, bg=T['bg'])
        main.pack(fill='both', expand=True, padx=15, pady=(0, 10))

        canvas = tk.Canvas(main, bg=T['bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(main, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg=T['bg'])

        scrollable.bind('<Configure>',
                        lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Скролл колесом мыши
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all('<MouseWheel>', _on_mousewheel)

        # Заголовок таблицы
        self._build_table_header(scrollable)

        # Строки
        include_list = set(self.data.get('_включать_в_аналитику', []))
        row_names = [k for k, v in self.data.items()
                     if not k.startswith('_') and isinstance(v, dict)]

        for i, name in enumerate(row_names):
            self._build_row(scrollable, i + 1, name, self.data[name],
                            in_analytics=name in include_list)

        # Строка плана группы
        plan_data = self.data.get('_план_группы_по_месяцам_руб', {})
        self._build_plan_row(scrollable, len(row_names) + 2, plan_data)

        # Кнопки внизу
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

        # Подсказка снизу
        tk.Label(btn_frame,
                 text="Tab — следующая ячейка · Shift+Tab — предыдущая · Enter — сохранить",
                 bg=T['bg'], fg=T['muted'],
                 font=('Segoe UI', 8)).pack(side='left')

        self.bind('<Return>', lambda e: self._save())
        self.bind('<Escape>', lambda e: self.destroy())

    # ── Заголовок таблицы ────────────────────────────────────
    def _build_table_header(self, parent):
        T = self.theme
        hdr = tk.Frame(parent, bg=T['surface'])
        hdr.grid(row=0, column=0, columnspan=15, sticky='ew', pady=(0, 2))

        tk.Label(hdr, text="Статья дохода", width=36, anchor='w',
                 bg=T['surface'], fg=T['text'],
                 font=('Segoe UI', 9, 'bold'),
                 padx=6, pady=6).grid(row=0, column=0)

        tk.Label(hdr, text="В аналитике", width=11, anchor='center',
                 bg=T['surface'], fg=T['text'],
                 font=('Segoe UI', 9, 'bold'),
                 padx=4, pady=6).grid(row=0, column=1)

        for i, m_label in enumerate(MONTH_LABELS):
            tk.Label(hdr, text=m_label, width=10, anchor='center',
                     bg=T['surface'], fg=T['text'],
                     font=('Segoe UI', 9, 'bold'),
                     padx=2, pady=6).grid(row=0, column=i + 2)

        tk.Label(hdr, text="Итого", width=12, anchor='e',
                 bg=T['surface'], fg=T['muted'],
                 font=('Segoe UI', 9, 'bold'),
                 padx=6, pady=6).grid(row=0, column=14)

    # ── Одна строка таблицы ──────────────────────────────────
    def _build_row(self, parent, row_idx, name, monthly_data, in_analytics=False):
        T = self.theme
        bg_color = T['surface'] if row_idx % 2 == 0 else T['bg']

        frame = tk.Frame(parent, bg=bg_color)
        frame.grid(row=row_idx, column=0, columnspan=15, sticky='ew')

        # Название
        tk.Label(frame, text=name, width=36, anchor='w',
                 bg=bg_color, fg=T['text'],
                 font=('Segoe UI', 9),
                 padx=6, pady=3).grid(row=0, column=0)

        # Чекбокс "в аналитике"
        var = tk.BooleanVar(value=in_analytics)
        self.include_vars[name] = var
        cb = tk.Checkbutton(frame, variable=var,
                            bg=bg_color, activebackground=bg_color,
                            cursor='hand2')
        cb.grid(row=0, column=1)

        # Поля месяцев
        total_var = tk.StringVar(value='0')
        for i, m in enumerate(MONTHS):
            value = monthly_data.get(m, 0) or 0
            entry_var = tk.StringVar(value=self._fmt(value))
            entry = tk.Entry(frame, width=11,
                             textvariable=entry_var,
                             justify='right',
                             bg=T['entry_bg'], fg=T['text'],
                             relief='flat', borderwidth=1,
                             font=('Segoe UI', 9))
            entry.grid(row=0, column=i + 2, padx=1, pady=1)
            entry_var.trace_add('write',
                lambda *a, n=name: self._recalc_total(n))
            self.entries[(name, m)] = entry_var

        # Итого
        total_label = tk.Label(frame, textvariable=total_var, width=12, anchor='e',
                               bg=bg_color, fg=T['muted'],
                               font=('Segoe UI', 9),
                               padx=6, pady=3)
        total_label.grid(row=0, column=14)
        self.entries[(name, '_total_var')] = total_var
        self._recalc_total(name)

    # ── Строка плана группы ──────────────────────────────────
    def _build_plan_row(self, parent, row_idx, plan_data):
        T = self.theme

        # Разделитель
        sep = tk.Frame(parent, bg=T['border'], height=1)
        sep.grid(row=row_idx - 1, column=0, columnspan=15, sticky='ew', pady=6)

        frame = tk.Frame(parent, bg=T['surface'])
        frame.grid(row=row_idx, column=0, columnspan=15, sticky='ew')

        tk.Label(frame, text="План группы (руб.)", width=36, anchor='w',
                 bg=T['surface'], fg=T['text'],
                 font=('Segoe UI', 9, 'bold'),
                 padx=6, pady=4).grid(row=0, column=0)

        tk.Label(frame, text="—", width=11, anchor='center',
                 bg=T['surface'], fg=T['muted'],
                 font=('Segoe UI', 9),
                 padx=4).grid(row=0, column=1)

        total_var = tk.StringVar(value='0')
        for i, m in enumerate(MONTHS):
            value = plan_data.get(m, 0) or 0
            var = tk.StringVar(value=self._fmt(value))
            entry = tk.Entry(frame, width=11,
                             textvariable=var,
                             justify='right',
                             bg=T['entry_bg'], fg=T['text'],
                             relief='flat', borderwidth=1,
                             font=('Segoe UI', 9))
            entry.grid(row=0, column=i + 2, padx=1, pady=2)
            var.trace_add('write', lambda *a: self._recalc_plan_total())
            self.plan_entries[m] = var

        self._plan_total_var = total_var
        tk.Label(frame, textvariable=total_var, width=12, anchor='e',
                 bg=T['surface'], fg=C_ORANGE,
                 font=('Segoe UI', 9, 'bold'),
                 padx=6, pady=4).grid(row=0, column=14)
        self._recalc_plan_total()

    # ── Расчёт итогов ────────────────────────────────────────
    def _recalc_total(self, row_name):
        total = 0.0
        for m in MONTHS:
            var = self.entries.get((row_name, m))
            if var:
                total += self._parse(var.get())
        total_var = self.entries.get((row_name, '_total_var'))
        if total_var:
            total_var.set(self._fmt(total))

    def _recalc_plan_total(self):
        total = 0.0
        for m in MONTHS:
            var = self.plan_entries.get(m)
            if var:
                total += self._parse(var.get())
        if hasattr(self, '_plan_total_var'):
            self._plan_total_var.set(self._fmt(total))

    # ── Парсинг/форматирование ──────────────────────────────
    def _parse(self, s):
        if not s:
            return 0.0
        s = str(s).strip().replace(' ', '').replace(',', '.')
        try:
            return float(s)
        except Exception:
            return 0.0

    def _fmt(self, v):
        try:
            v = float(v)
        except Exception:
            return '0'
        if v == 0:
            return '0'
        if abs(v) >= 1000:
            return f"{v:,.0f}".replace(',', ' ')
        return f"{v:,.2f}".replace(',', ' ')

    # ── Сохранение ───────────────────────────────────────────
    def _save(self):
        # Пересобираем данные
        new_data = dict(self.data)  # копия с _описание, _год, служебными

        # Обновляем месячные суммы
        row_names = [k for k, v in self.data.items()
                     if not k.startswith('_') and isinstance(v, dict)]
        for name in row_names:
            new_data[name] = {}
            for m in MONTHS:
                var = self.entries.get((name, m))
                if var:
                    new_data[name][m] = self._parse(var.get())

        # Обновляем список "включать в аналитику"
        new_data['_включать_в_аналитику'] = [
            name for name, var in self.include_vars.items() if var.get()
        ]

        # Обновляем план группы
        plan = {}
        for m in MONTHS:
            var = self.plan_entries.get(m)
            if var:
                plan[m] = self._parse(var.get())
        new_data['_план_группы_по_месяцам_руб'] = plan

        # Сохраняем
        try:
            with open(self.json_path, 'w', encoding='utf-8') as f:
                json.dump(new_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))
            return

        messagebox.showinfo("Сохранено",
            "Данные сохранены.\n\nЗапусти аналитику заново чтобы\n"
            "новые цифры учитывались в отчёте.")
        self.destroy()


# ── Запуск для теста ─────────────────────────────────────────
if __name__ == '__main__':
    # Тестовый запуск без основного app
    root = tk.Tk()
    root.withdraw()
    json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             'external_income.json')
    ExternalIncomeEditor(root, json_path)
    root.mainloop()
