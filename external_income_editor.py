# ============================================================
# external_income_editor.py — Редактор external_income.json
# Открывается как модальное окно из основного app.py
# ============================================================

import json
import os
import tkinter as tk
from tkinter import messagebox, ttk

import ui_theme as ui
from crm_data import parse_money
from jsonio import save_json_atomic
from ui_theme import T, px

MONTHS = [f"{m:02d}" for m in range(1, 13)]
MONTH_LABELS = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн',
                'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек']


def _cell(parent, var, bold=False):
    """Компактная ячейка таблицы: плоское поле с обводкой, акцент при фокусе."""
    return tk.Entry(parent, textvariable=var, width=11, justify='right', relief='flat', bd=0,
                    font=ui.FONTS['strong' if bold else 'body'], bg=T['card'], fg=T['text'],
                    insertbackground=T['text'], highlightthickness=1,
                    highlightbackground=T['stroke'], highlightcolor=T['accent'])


class ExternalIncomeEditor(ui.FluentDialog):
    """
    Таблица: строки = статьи доходов, колонки = 12 месяцев.
    Позволяет редактировать суммы, переключать «в аналитике»,
    сохраняет обратно в JSON.
    """

    def __init__(self, parent, json_path, **_):
        self.json_path = json_path
        self.data = {}
        self.entries = {}          # {(row_name, month): StringVar}
        self.include_vars = {}     # {row_name: BooleanVar}
        self.plan_entries = {}     # {month: StringVar}
        self._totals = {}          # {row_name: StringVar}
        super().__init__(parent, "Внешние доходы",
                         "Суммы в рублях. «В аналитике» — строка добавляется к CRM-выручке "
                         "в месячной статистике (обычно только программатик).",
                         size=(1500, 700), minsize=(900, 480))
        # При ошибке чтения окно уже уничтожено — UI не строим
        if not self._load_data():
            return
        self._build_ui()
        self.show_modal(parent)

    # ── Загрузка JSON ────────────────────────────────────────
    def _load_data(self):
        try:
            with open(self.json_path, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
            return True
        except Exception as e:
            messagebox.showerror("Ошибка чтения JSON", f"Не удалось прочитать файл:\n{e}", parent=self.master)
            self.destroy()
            return False

    # ── Построение интерфейса ────────────────────────────────
    def _build_ui(self):
        outer, body = ui.card(self.body, pad=8)
        outer.pack(fill='both', expand=True)

        canvas = tk.Canvas(body, bg=T['card'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(body, orient='vertical', command=canvas.yview,
                                  style='Thin.Vertical.TScrollbar')
        # Горизонтальная прокрутка — на узком экране 12 месяцев + итог не влезают
        xscroll = ttk.Scrollbar(body, orient='horizontal', command=canvas.xview,
                                style='Thin.Horizontal.TScrollbar')
        table = tk.Frame(canvas, bg=T['card'])
        table.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=table, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set, xscrollcommand=xscroll.set)
        xscroll.pack(side='bottom', fill='x')
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        # Через общий обработчик приложения: собственный bind_all перехватывал
        # колесо глобально, и после закрытия окна главное окно не прокручивалось
        if hasattr(self.master, '_bind_mousewheel'):
            self.master._bind_mousewheel(canvas)

        # Заголовок
        heads = ["Статья дохода", "В аналитике"] + MONTH_LABELS + ["Итого"]
        for c, h in enumerate(heads):
            ui.label(table, h, 'caption', fg='text2',
                     anchor='w' if c == 0 else 'e' if c == 14 else 'center').grid(
                row=0, column=c, sticky='ew', padx=px(3), pady=(px(4), px(6)))

        include_list = set(self.data.get('_включать_в_аналитику', []))
        row_names = [k for k, v in self.data.items() if not k.startswith('_') and isinstance(v, dict)]
        for i, name in enumerate(row_names, start=1):
            self._build_row(table, i, name, self.data[name], name in include_list)

        r = len(row_names) + 1
        tk.Frame(table, bg=T['divider'], height=1).grid(row=r, column=0, columnspan=15, sticky='ew', pady=px(8))
        self._build_plan_row(table, r + 1, self.data.get('_план_группы_по_месяцам_руб', {}))

        self.hint.configure(text="Tab — следующая ячейка · Esc — закрыть без сохранения")
        self.add_button("Отмена", self.destroy)
        self.add_button("Сохранить", self._save, kind="Accent")

    def _build_row(self, table, r, name, monthly_data, in_analytics):
        ui.label(table, name, anchor='w').grid(row=r, column=0, sticky='w', padx=px(3), pady=px(2))
        var = tk.BooleanVar(value=in_analytics)
        self.include_vars[name] = var
        ttk.Checkbutton(table, variable=var).grid(row=r, column=1)
        for i, m in enumerate(MONTHS):
            entry_var = tk.StringVar(value=self._fmt(monthly_data.get(m, 0) or 0))
            _cell(table, entry_var).grid(row=r, column=i + 2, padx=px(2), pady=px(2), ipady=px(3))
            entry_var.trace_add('write', lambda *a, n=name: self._recalc_total(n))
            self.entries[(name, m)] = entry_var
        self._totals[name] = tk.StringVar()
        ui.label(table, textvariable=self._totals[name], fg='text2', anchor='e', width=12).grid(
            row=r, column=14, sticky='e', padx=px(3))
        self._recalc_total(name)

    def _build_plan_row(self, table, r, plan_data):
        ui.label(table, "План группы", 'strong', anchor='w').grid(row=r, column=0, sticky='w', padx=px(3))
        for i, m in enumerate(MONTHS):
            var = tk.StringVar(value=self._fmt(plan_data.get(m, 0) or 0))
            _cell(table, var).grid(row=r, column=i + 2, padx=px(2), pady=px(2), ipady=px(3))
            var.trace_add('write', lambda *a: self._recalc_plan_total())
            self.plan_entries[m] = var
        self._plan_total_var = tk.StringVar()
        ui.label(table, textvariable=self._plan_total_var, role='strong', fg='accent', anchor='e',
                 width=12).grid(row=r, column=14, sticky='e', padx=px(3))
        self._recalc_plan_total()

    # ── Расчёт итогов ────────────────────────────────────────
    def _sum(self, variables):
        return sum(self._parse(v.get()) or 0.0 for v in variables)

    def _recalc_total(self, row_name):
        if row_name in self._totals:
            self._totals[row_name].set(self._fmt(self._sum(self.entries[(row_name, m)] for m in MONTHS
                                                           if (row_name, m) in self.entries)))

    def _recalc_plan_total(self):
        if hasattr(self, '_plan_total_var'):
            self._plan_total_var.set(self._fmt(self._sum(self.plan_entries.values())))

    # ── Парсинг/форматирование ──────────────────────────────
    def _parse(self, s):
        """Число из ячейки; None — если не распозналось (пустая ячейка = 0)."""
        if not str(s).strip():
            return 0.0
        # parse_money понимает неразрывные пробелы и разные разделители,
        # которые приходят при копировании из Excel
        return parse_money(s)

    def _fmt(self, v):
        try:
            v = float(v)
        except Exception:
            return '0'
        if v == 0:
            return '0'
        # Копейки показываем всегда, когда они есть: _save читает значение
        # обратно из поля, и округление при показе тихо портило данные.
        if v == round(v):
            return f"{v:,.0f}".replace(',', ' ')
        return f"{v:,.2f}".replace(',', ' ')

    # ── Сохранение ───────────────────────────────────────────
    def _save(self):
        new_data = dict(self.data)  # копия с _описание, _год, служебными
        bad_cells = []

        def read(var, where):
            v = self._parse(var.get())
            if v is None:
                bad_cells.append(f"{where}: «{var.get()}»")
            return v

        for name in [k for k, v in self.data.items() if not k.startswith('_') and isinstance(v, dict)]:
            new_data[name] = {m: read(self.entries[(name, m)], f"{name}, {MONTH_LABELS[int(m) - 1]}")
                              for m in MONTHS if (name, m) in self.entries}
        new_data['_включать_в_аналитику'] = [n for n, var in self.include_vars.items() if var.get()]
        new_data['_план_группы_по_месяцам_руб'] = {
            m: read(var, f"План группы, {MONTH_LABELS[int(m) - 1]}") for m, var in self.plan_entries.items()}

        # Нераспознанные числа не превращаем молча в 0 — просим исправить
        if bad_cells:
            messagebox.showerror(
                "Не сохранено",
                "Не удалось распознать числа:\n\n" + "\n".join(bad_cells[:15]) +
                ("\n…" if len(bad_cells) > 15 else ""), parent=self)
            return

        try:
            save_json_atomic(self.json_path, new_data)
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e), parent=self)
            return

        messagebox.showinfo("Сохранено", "Данные сохранены. Запустите анализ заново, "
                                         "чтобы новые цифры попали в отчёт.", parent=self)
        self.destroy()


if __name__ == '__main__':
    root = tk.Tk()
    ui.init(root)
    root.withdraw()
    ExternalIncomeEditor(root, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'external_income.json'))
    root.mainloop()
