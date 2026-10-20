# ============================================================
# client_aliases_editor.py — Редактор client_aliases.json
# ============================================================

import json
import os
import tkinter as tk
from tkinter import messagebox, ttk

import ui_theme as ui
from jsonio import save_json_atomic
from ui_theme import ICON, T, px


class ClientAliasesEditor(ui.FluentDialog):
    """Правила нормализации клиентов: каноничное имя + список начал слов для поиска."""

    def __init__(self, parent, json_path, **_):
        self.json_path = json_path
        self.groups = []  # list of {'canonical': str, 'patterns': list[str]}
        self._extra = {}  # прочие ключи верхнего уровня — сохраняем как есть
        self._rows = []
        super().__init__(parent, "Алиасы клиентов",
                         "Если имя клиента содержит слово, начинающееся с любого из образцов "
                         "(через запятую), оно заменяется на каноничное. Регистр не важен. "
                         "Порядок групп важен: побеждает первое совпадение.",
                         size=(760, 580), minsize=(600, 400))
        if not self._load_data():
            return
        self._build_ui()
        self.show_modal(parent)

    # ── I/O ────────────────────────────────────────────────

    def _load_data(self):
        if os.path.exists(self.json_path):
            try:
                with open(self.json_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.groups = [
                    {'canonical': g.get('canonical', ''), 'patterns': list(g.get('patterns', []))}
                    for g in data.get('groups', [])
                ]
                self._extra = {k: v for k, v in data.items() if k not in ('groups', '_описание')}
            except Exception as e:
                # Битый файл не открываем «пустым»: сохранение стёрло бы все правила
                messagebox.showerror("Ошибка чтения JSON",
                    f"Не удалось прочитать {os.path.basename(self.json_path)}:\n{e}\n\n"
                    "Исправьте файл вручную, иначе сохранение затрёт его.", parent=self.master)
                self.destroy()
                return False
        return True

    def _save(self):
        groups = []
        for row in self._rows:
            canonical = row['canonical_var'].get().strip()
            # Поиск идёт по имени клиента в верхнем регистре
            patterns = [p.strip().upper() for p in row['patterns_var'].get().split(',') if p.strip()]
            if canonical:
                groups.append({'canonical': canonical, 'patterns': patterns})

        data = {
            '_описание': (
                'Правила нормализации клиентов. Каждая группа: '
                "'canonical' — итоговое имя, 'patterns' — начало слова для поиска "
                '(в верхнем регистре). Редактируется через GUI (Данные → Алиасы клиентов).'
            ),
            **self._extra,
            'groups': groups
        }
        try:
            save_json_atomic(self.json_path, data)
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e), parent=self)
            return

        from crm_data import reload_aliases
        reload_aliases()
        messagebox.showinfo("Сохранено", "Алиасы клиентов обновлены. "
                                         "Запустите анализ заново для пересчёта.", parent=self)
        self.destroy()

    # ── UI ─────────────────────────────────────────────────

    def _build_ui(self):
        outer, body = ui.card(self.body, pad=8)
        outer.pack(fill='both', expand=True)

        head = tk.Frame(body, bg=T['card'])
        head.pack(fill='x', padx=px(8), pady=(px(4), px(4)))
        ui.label(head, "Каноничное имя", 'caption', fg='text2', width=26, anchor='w').pack(side='left')
        ui.label(head, "Образцы (через запятую)", 'caption', fg='text2', anchor='w').pack(side='left')

        canvas = tk.Canvas(body, bg=T['card'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(body, orient='vertical', command=canvas.yview,
                                  style='Thin.Vertical.TScrollbar')
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        self._scroll_frame = tk.Frame(canvas, bg=T['card'])
        win = canvas.create_window((0, 0), window=self._scroll_frame, anchor='nw')
        self._scroll_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(win, width=e.width))
        if hasattr(self.master, '_bind_mousewheel'):
            self.master._bind_mousewheel(canvas)
        self._canvas = canvas

        for g in self.groups:
            self._add_row(g['canonical'], g['patterns'])

        b = ui.button(self._footer_inner, "Добавить группу", self._add_row, icon=ICON['add'])
        b.pack(side='left')
        self.hint.pack_forget()
        self.add_button("Отмена", self.destroy)
        self.add_button("Сохранить", self._save, kind="Accent")

    def _add_row(self, canonical='', patterns=None):
        row_frame = tk.Frame(self._scroll_frame, bg=T['card'])
        row_frame.pack(fill='x', padx=px(8), pady=px(3))
        canonical_var = tk.StringVar(value=canonical)
        patterns_var = tk.StringVar(value=', '.join(patterns or []))
        ttk.Entry(row_frame, textvariable=canonical_var, width=24).pack(side='left', padx=(0, px(8)))
        row_data = {'canonical_var': canonical_var, 'patterns_var': patterns_var, 'frame': row_frame}

        def remove(rd=row_data):
            rd['frame'].destroy()
            self._rows.remove(rd)

        ui.button(row_frame, "", remove, kind="Subtle", icon=ICON['close']).pack(side='right')
        ttk.Entry(row_frame, textvariable=patterns_var).pack(side='left', fill='x', expand=True, padx=(0, px(4)))
        self._rows.append(row_data)
        # прокрутить вниз при добавлении
        self._scroll_frame.update_idletasks()
        self._canvas.yview_moveto(1.0)


if __name__ == '__main__':
    root = tk.Tk()
    ui.init(root)
    root.withdraw()
    ClientAliasesEditor(root, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'client_aliases.json'))
    root.mainloop()
