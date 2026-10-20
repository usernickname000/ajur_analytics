# ============================================================
# verified_figures_editor.py — Редактор verified_figures.json
# Для редактирования верифицированных годовых итогов
# ============================================================

import json
import os
import tkinter as tk
from tkinter import messagebox, ttk

import ui_theme as ui
from crm_data import parse_money
from jsonio import save_json_atomic
from ui_theme import T, px


class VerifiedFiguresEditor(ui.FluentDialog):
    """Верифицированные цифры бухгалтерии: общие итоги + программатик + прочие."""

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

    def __init__(self, parent, json_path, **_):
        self.json_path = json_path
        self.data = {}
        self.entries = {}
        super().__init__(parent, "Верифицированные цифры",
                         "Годовые итоги из отчёта бухгалтерии, в рублях. "
                         "Обновляются раз в год или при значимых корректировках.",
                         size=(680, 560))
        self.year_var = tk.StringVar()
        if not self._load_data():
            return
        self._build_ui()
        self.show_modal(parent)

    def _load_data(self):
        # Битый файл не открываем «пустым»: сохранение затёрло бы его нулями
        if os.path.exists(self.json_path):
            try:
                with open(self.json_path, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
            except Exception as e:
                messagebox.showerror("Ошибка чтения JSON",
                    f"Не удалось прочитать {os.path.basename(self.json_path)}:\n{e}\n\n"
                    "Исправьте файл вручную, иначе сохранение затрёт его.", parent=self.master)
                self.destroy()
                return False
        return True

    def _build_ui(self):
        outer, body = ui.card(self.body)
        outer.pack(fill='both', expand=True)
        body.columnconfigure(0, weight=1)

        ui.label(body, "Отчётный год", 'strong').grid(row=0, column=0, sticky='w', pady=(0, px(12)))
        self.year_var.set(str(self.data.get('_год', 2025)))
        ttk.Entry(body, textvariable=self.year_var, width=8, justify='center').grid(
            row=0, column=1, sticky='e', pady=(0, px(12)))

        for i, (key, title, hint) in enumerate(self.FIELDS, start=1):
            tk.Frame(body, bg=T['divider'], height=1).grid(row=i * 2 - 1, column=0, columnspan=2,
                                                           sticky='ew', pady=px(4))
            text = tk.Frame(body, bg=T['card'])
            text.grid(row=i * 2, column=0, sticky='w', pady=px(4))
            ui.label(text, title).pack(anchor='w')
            ui.label(text, hint, 'caption', fg='text2').pack(anchor='w')
            var = tk.StringVar(value=self._fmt(self.data.get(key, 0)))
            ttk.Entry(body, textvariable=var, width=18, justify='right').grid(row=i * 2, column=1, sticky='e')
            self.entries[key] = var

        self.add_button("Отмена", self.destroy)
        self.add_button("Сохранить", self._save, kind="Accent")

    def _parse(self, s):
        """Целое число рублей; None — если не распозналось (пусто = 0)."""
        if not str(s).strip():
            return 0
        v = parse_money(s)
        return None if v is None else int(round(v))

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
            messagebox.showerror("Не сохранено",
                f"Отчётный год «{self.year_var.get()}» — не число.", parent=self)
            return
        new_data.setdefault('_описание', "Верифицированные годовые итоги из бухгалтерии.")

        # Нераспознанное число не превращаем в 0: analytics считает 0
        # «не задано» и молча пропускает сверку
        labels = {key: label for key, label, _ in self.FIELDS}
        bad = []
        for key, var in self.entries.items():
            v = self._parse(var.get())
            if v is None:
                bad.append(f"{labels[key]}: «{var.get()}»")
            new_data[key] = v
        if bad:
            messagebox.showerror("Не сохранено",
                "Не удалось распознать числа:\n\n" + "\n".join(bad), parent=self)
            return

        try:
            save_json_atomic(self.json_path, new_data)
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e), parent=self)
            return

        messagebox.showinfo("Сохранено", "Верифицированные цифры обновлены. "
                                         "Запустите анализ заново для пересчёта.", parent=self)
        self.destroy()


if __name__ == '__main__':
    root = tk.Tk()
    ui.init(root)
    root.withdraw()
    VerifiedFiguresEditor(root, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'verified_figures.json'))
    root.mainloop()
