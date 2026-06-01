# ============================================================
# client_aliases_editor.py — Редактор client_aliases.json
# ============================================================

import tkinter as tk
from tkinter import messagebox
import json
import os

C_ORANGE = "#F38120"


class ClientAliasesEditor(tk.Toplevel):
    """
    Диалог для редактирования правил нормализации клиентов.
    Каждая группа: каноничное имя + список подстрок для поиска.
    """

    def __init__(self, parent, json_path, theme=None):
        super().__init__(parent)
        self.title("Алиасы клиентов")
        self.geometry("700x560")
        self.minsize(600, 400)
        self.json_path = json_path
        self.theme = theme or {
            'bg': '#F5F5F5', 'surface': '#FFFFFF', 'text': '#1A1A1A',
            'muted': '#777777', 'entry_bg': '#FAFAFA', 'border': '#DEDEDE',
            'surface2': '#EFEFEF'
        }
        T = self.theme
        self.configure(bg=T['bg'])

        self.groups = []  # list of {'canonical': str, 'patterns': list[str]}
        self._load_data()
        self._build_ui()

        self.transient(parent)
        self.grab_set()
        self.focus_set()

    # ── I/O ────────────────────────────────────────────────

    def _load_data(self):
        if os.path.exists(self.json_path):
            try:
                with open(self.json_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.groups = [
                    {'canonical': g.get('canonical', ''),
                     'patterns': list(g.get('patterns', []))}
                    for g in data.get('groups', [])
                ]
            except Exception:
                self.groups = []

    def _save(self):
        groups = []
        for row in self._rows:
            canonical = row['canonical_var'].get().strip()
            raw = row['patterns_var'].get().strip()
            patterns = [p.strip() for p in raw.split(',') if p.strip()]
            if not canonical:
                continue
            groups.append({'canonical': canonical, 'patterns': patterns})

        data = {
            '_описание': (
                'Правила нормализации клиентов. Каждая группа: '
                "'canonical' — итоговое имя, 'patterns' — подстроки для поиска "
                '(в верхнем регистре). Редактируется через GUI (кнопка \'Алиасы клиентов\').'
            ),
            'groups': groups
        }
        try:
            with open(self.json_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))
            return

        # сбрасываем кэш в analytics
        try:
            from analytics import reload_aliases
            reload_aliases()
        except Exception:
            pass

        messagebox.showinfo(
            "Сохранено",
            "Алиасы клиентов обновлены.\nЗапусти аналитику заново для пересчёта."
        )
        self.destroy()

    # ── UI ─────────────────────────────────────────────────

    def _build_ui(self):
        T = self.theme

        # Шапка
        header = tk.Frame(self, bg=C_ORANGE, height=50)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)
        tk.Label(header, text="Алиасы клиентов",
                 bg=C_ORANGE, fg='white',
                 font=('Segoe UI', 13, 'bold')).pack(side='left', padx=20, pady=12)

        # Пояснение
        notice = tk.Frame(self, bg=T['bg'])
        notice.pack(fill='x', padx=15, pady=(10, 4))
        tk.Label(
            notice,
            text=(
                "Если имя клиента содержит любую из подстрок (через запятую) — "
                "оно заменяется на каноничное.\n"
                "Поиск ведётся в верхнем регистре. Порядок групп важен: "
                "побеждает первое совпадение."
            ),
            bg=T['bg'], fg=T['muted'],
            font=('Segoe UI', 9),
            justify='left'
        ).pack(side='left')

        # Заголовки колонок
        cols = tk.Frame(self, bg=T['bg'])
        cols.pack(fill='x', padx=15, pady=(6, 2))
        tk.Label(cols, text="Каноничное имя", bg=T['bg'], fg=T['muted'],
                 font=('Segoe UI', 8, 'bold'), width=22, anchor='w').pack(side='left', padx=(0, 6))
        tk.Label(cols, text="Подстроки (через запятую)", bg=T['bg'], fg=T['muted'],
                 font=('Segoe UI', 8, 'bold'), anchor='w').pack(side='left')

        # Прокручиваемая область
        container = tk.Frame(self, bg=T['bg'])
        container.pack(fill='both', expand=True, padx=15, pady=4)

        canvas = tk.Canvas(container, bg=T['bg'], highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        self._scroll_frame = tk.Frame(canvas, bg=T['bg'])
        self._scroll_win = canvas.create_window((0, 0), window=self._scroll_frame, anchor='nw')

        self._scroll_frame.bind('<Configure>', lambda e: canvas.configure(
            scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(
            self._scroll_win, width=e.width))
        canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(
            int(-1 * (e.delta / 120)), 'units'))

        self._canvas = canvas
        self._rows = []

        for g in self.groups:
            self._add_row(g['canonical'], g['patterns'])

        # Кнопки внизу
        btn_frame = tk.Frame(self, bg=T['bg'])
        btn_frame.pack(fill='x', padx=15, pady=(4, 12))

        tk.Button(btn_frame, text="+ Добавить группу",
                  command=self._add_row,
                  bg=T.get('surface2', T['surface']), fg=T['text'],
                  relief='flat', borderwidth=1,
                  font=('Segoe UI', 9),
                  padx=12, pady=6,
                  cursor='hand2').pack(side='left')

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

        self.bind('<Escape>', lambda e: self.destroy())

    def _add_row(self, canonical='', patterns=None):
        T = self.theme
        if patterns is None:
            patterns = []

        row_frame = tk.Frame(self._scroll_frame, bg=T['surface'])
        row_frame.pack(fill='x', pady=3)

        canonical_var = tk.StringVar(value=canonical)
        patterns_var = tk.StringVar(value=', '.join(patterns))

        tk.Entry(row_frame, textvariable=canonical_var,
                 width=24,
                 bg=T['entry_bg'], fg=T['text'],
                 relief='flat', borderwidth=1,
                 font=('Segoe UI', 10)).pack(side='left', padx=(8, 6), pady=6)

        tk.Entry(row_frame, textvariable=patterns_var,
                 bg=T['entry_bg'], fg=T['text'],
                 relief='flat', borderwidth=1,
                 font=('Segoe UI', 10)).pack(side='left', fill='x', expand=True, padx=(0, 6), pady=6)

        row_data = {'canonical_var': canonical_var, 'patterns_var': patterns_var,
                    'frame': row_frame}

        def remove(rd=row_data):
            rd['frame'].destroy()
            self._rows.remove(rd)

        tk.Button(row_frame, text="✕",
                  command=remove,
                  bg=T['surface'], fg=T['muted'],
                  relief='flat', borderwidth=0,
                  font=('Segoe UI', 10),
                  padx=6, pady=4,
                  cursor='hand2').pack(side='right', padx=(0, 6))

        self._rows.append(row_data)
        # прокрутить вниз при добавлении
        self._scroll_frame.update_idletasks()
        self._canvas.yview_moveto(1.0)


if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'client_aliases.json')
    ClientAliasesEditor(root, p)
    root.mainloop()
