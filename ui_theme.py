# ============================================================
# ui_theme.py — светлая тема в стиле Windows 11 (Fluent) для Tkinter
# Без внешних зависимостей: скругления рисуются в PhotoImage при
# запуске и подключаются как 9-slice элементы ttk.
# ============================================================

import ctypes
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

# ── Палитра ──────────────────────────────────────────────────
T = {
    # Поверхности
    "bg":           "#F3F3F3",   # фон окна
    "card":         "#FFFFFF",   # карточка
    "stroke":       "#E5E5E5",   # обводка карточки / контрола
    "stroke_strong": "#D1D1D1",  # нижняя граница поля ввода
    "subtle":       "#F5F5F5",   # hover вторичных элементов
    "subtle_press": "#EDEDED",
    "divider":      "#EBEBEB",
    # Текст
    "text":         "#1B1B1B",
    "text2":        "#5F5F5F",
    "text3":        "#8A8A8A",
    "disabled":     "#A0A0A0",
    # Акцент — фирменный оранжевый, затемнённый для контраста белого текста
    "accent":       "#C2570C",
    "accent_hover": "#D06A10",
    "accent_press": "#A84B0A",
    "accent_soft":  "#FCEBDD",   # выделенный пункт навигации
    "brand":        "#F38120",   # индикатор навигации, фокус, прогресс
    # Статусы (InfoBar)
    "ok":   "#0F7B0F", "ok_bg":   "#DFF6DD",
    "warn": "#9D5D00", "warn_bg": "#FFF4CE",
    "err":  "#C42B1C", "err_bg":  "#FDE7E9",
    "info": "#005FB8", "info_bg": "#EFF4F9",
}

# Иконки Segoe Fluent Icons (фоллбэк — Segoe MDL2 Assets, коды те же)
ICON = {
    "analysis": "", "compare": "", "plan": "", "data": "",
    "play": "", "globe": "", "folder": "", "file": "",
    "ok": "", "warn": "", "err": "", "info": "",
    "close": "", "chev_right": "", "chev_down": "",
    "money": "", "import": "", "link": "", "add": "",
}

_SCALE = 1.0
FONTS = {}


def px(n):
    """Пиксели макета (96 dpi) → экранные с учётом масштаба Windows."""
    return int(round(n * _SCALE))


def enable_dpi_awareness():
    """Чёткий текст при масштабе 125–150 %. Вызывать до создания Tk()."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def _hwnd(win):
    try:
        return ctypes.windll.user32.GetParent(win.winfo_id()) or win.winfo_id()
    except Exception:
        return None


def style_window(win, caption=True):
    """Скруглённые углы окна (Win11) и заголовок цвета фона. На Win10 — без эффекта."""
    try:
        win.update_idletasks()
        hwnd = _hwnd(win)
        if not hwnd:
            return
        dwm = ctypes.windll.dwmapi
        dwm.DwmSetWindowAttribute(hwnd, 33, ctypes.byref(ctypes.c_int(2)), 4)  # round corners
        if caption:
            r, g, b = (int(T["bg"][i:i + 2], 16) for i in (1, 3, 5))
            dwm.DwmSetWindowAttribute(hwnd, 35, ctypes.byref(ctypes.c_int(r | g << 8 | b << 16)), 4)
    except Exception:
        pass


# ── Рисование скруглённых прямоугольников ───────────────────

def _rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (1, 3, 5))


def _inside(x, y, x0, y0, x1, y1, r):
    if x < x0 or x > x1 or y < y0 or y > y1:
        return False
    dx = max(x0 + r - x, 0, x - (x1 - r))
    dy = max(y0 + r - y, 0, y - (y1 - r))
    return dx * dx + dy * dy <= r * r


def _rounded(root, radius, bg, fill, stroke=None, bottom=None, bottom_w=1, ss=4):
    """
    Квадратная картинка для 9-slice: скруглённый прямоугольник fill с обводкой
    stroke (1 px) и необязательной нижней линией bottom (bottom_w px) на фоне bg.
    Сглаживание — суперсэмплинг ss×ss.
    """
    r = max(radius, 1)
    size = 2 * r + 4
    sw = 1 if stroke else 0
    cbg, cfill = _rgb(bg), _rgb(fill)
    cstroke = _rgb(stroke) if stroke else cfill
    cbottom = _rgb(bottom) if bottom else None
    img = tk.PhotoImage(master=root, width=size, height=size)
    rows = []
    for py in range(size):
        row = []
        for pxl in range(size):
            outer = inner = line = 0
            for sy in range(ss):
                for sx in range(ss):
                    x = pxl + (sx + 0.5) / ss
                    y = py + (sy + 0.5) / ss
                    if _inside(x, y, 0, 0, size, size, r):
                        outer += 1
                        if cbottom and y >= size - bottom_w:
                            line += 1
                        elif _inside(x, y, sw, sw, size - sw, size - sw, max(r - sw, 0)):
                            inner += 1
            n = ss * ss
            wl, wi = line / n, inner / n
            ws = outer / n - wl - wi
            wb = 1 - outer / n
            c = [cbg[k] * wb + cfill[k] * wi + cstroke[k] * ws + (cbottom[k] * wl if cbottom else 0)
                 for k in range(3)]
            row.append("#%02x%02x%02x" % tuple(int(round(v)) for v in c))
        rows.append("{" + " ".join(row) + "}")
    img.put(" ".join(rows))
    return img


# ── Инициализация темы ──────────────────────────────────────

def _family(*names):
    available = set(tkfont.families())
    return next((n for n in names if n in available), "TkDefaultFont")


def init(root):
    """Шрифты, ttk-стили и картинки элементов. Вызывать один раз после Tk()."""
    global _SCALE
    _SCALE = root.winfo_fpixels('1i') / 96.0

    display = _family("Segoe UI Variable Display", "Segoe UI")
    text = _family("Segoe UI Variable Text", "Segoe UI")
    small = _family("Segoe UI Variable Small", "Segoe UI")
    icons = _family("Segoe Fluent Icons", "Segoe MDL2 Assets")
    mono = _family("Cascadia Mono", "Consolas")
    FONTS.update({
        "title":    (display, 20, "bold"),
        "subtitle": (display, 14, "bold"),
        "kpi":      (display, 15, "bold"),
        "strong":   (text, 10, "bold"),
        "body":     (text, 10),
        "caption":  (small, 9),
        "mono":     (mono, 9),
        "icon":     (icons, 12),
        "icon_sm":  (icons, 10),
        "brand":    ("Georgia", 15, "bold italic"),
    })
    root.option_add("*Font", FONTS["body"])
    root.configure(bg=T["bg"])

    s = ttk.Style(root)
    s.theme_use("clam")
    s.configure(".", background=T["bg"], foreground=T["text"], font=FONTS["body"],
                bordercolor=T["stroke"], focuscolor=T["brand"])

    imgs = root._fluent_imgs = {}  # держим ссылки, иначе картинки соберёт GC

    def element(name, states, border, padding=None):
        base, *rest = states
        kw = {"border": border, "sticky": "nsew"}
        if padding is not None:
            kw["padding"] = padding
        s.element_create(name, "image", base, *rest, **kw)

    r4, r8 = px(4), px(8)

    # Карточка: белая, обводка 1 px, скругление 8
    imgs["card"] = _rounded(root, r8, T["bg"], T["card"], T["stroke"])
    element("Card.bg", [imgs["card"]], r8 + 1)
    s.layout("Card.TFrame", [("Card.bg", {"sticky": "nsew"})])
    s.configure("Card.TFrame", background=T["card"])

    # Кнопки (на белом и на сером фоне — сглаживание углов зависит от фона)
    for surface in ("card", "bg"):
        bg = T[surface]
        acc = [_rounded(root, r4, bg, T["accent"]),
               ("disabled", _rounded(root, r4, bg, T["subtle_press"])),
               ("pressed", _rounded(root, r4, bg, T["accent_press"])),
               ("active", _rounded(root, r4, bg, T["accent_hover"]))]
        std = [_rounded(root, r4, bg, "#FDFDFD", T["stroke"], bottom=T["stroke_strong"]),
               ("disabled", _rounded(root, r4, bg, T["subtle"], T["stroke"])),
               ("pressed", _rounded(root, r4, bg, T["subtle_press"], T["stroke"])),
               ("active", _rounded(root, r4, bg, T["subtle"], T["stroke"], bottom=T["stroke_strong"]))]
        sub = [_rounded(root, r4, bg, bg),
               ("pressed", _rounded(root, r4, bg, T["subtle_press"])),
               ("active", _rounded(root, r4, bg, T["subtle"] if surface == "card" else "#EAEAEA"))]
        imgs[f"btn_{surface}"] = (acc, std, sub)
        suffix = "" if surface == "card" else ".Bg"
        for kind, states in (("Accent", acc), ("Std", std), ("Subtle", sub)):
            el = f"{kind}{suffix}.Button.bg"
            element(el, states, r4 + 1)
            st = f"{kind}{suffix}.TButton"
            s.layout(st, [(el, {"sticky": "nsew", "children": [
                ("Button.padding", {"sticky": "nsew", "children": [
                    ("Button.label", {"sticky": "nsew"})]})]})])
            s.configure(st, padding=(px(14), px(5)), anchor="center", font=FONTS["body"],
                        background=bg,
                        foreground="#FFFFFF" if kind == "Accent" else T["text"])
            s.map(st, foreground=[("disabled", T["disabled"])] +
                  ([("pressed", "#F5DCC8")] if kind == "Accent" else [("pressed", T["text2"])]))
    s.configure("Accent.TButton", font=FONTS["strong"])
    s.configure("Accent.Bg.TButton", font=FONTS["strong"])

    # Поле ввода: скруглённое, нижняя граница; в фокусе — акцентная линия 2 px
    field = [_rounded(root, r4, T["card"], "#FFFFFF", T["stroke"], bottom=T["stroke_strong"]),
             ("focus", _rounded(root, r4, T["card"], "#FFFFFF", T["stroke"], bottom=T["accent"], bottom_w=px(2))),
             ("disabled", _rounded(root, r4, T["card"], T["subtle"], T["stroke"]))]
    imgs["field"] = field
    element("Fluent.Entry.field", field, r4 + 1)
    s.layout("TEntry", [("Fluent.Entry.field", {"sticky": "nsew", "children": [
        ("Entry.padding", {"sticky": "nsew", "children": [
            ("Entry.textarea", {"sticky": "nsew"})]})]})])
    s.configure("TEntry", padding=(px(8), px(5)), foreground=T["text"],
                selectbackground=T["brand"], selectforeground="#FFFFFF", insertcolor=T["text"])
    s.map("TEntry", foreground=[("disabled", T["disabled"])])

    # Прогресс: тонкая акцентная полоса
    s.configure("Thin.Horizontal.TProgressbar", troughcolor=T["divider"], background=T["brand"],
                bordercolor=T["divider"], lightcolor=T["brand"], darkcolor=T["brand"],
                thickness=px(3), borderwidth=0)

    # Скроллбар: только бегунок
    s.layout("Thin.Vertical.TScrollbar", [("Vertical.Scrollbar.trough", {"sticky": "ns", "children": [
        ("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
    s.configure("Thin.Vertical.TScrollbar", troughcolor=T["bg"], background="#C8C8C8",
                bordercolor=T["bg"], lightcolor="#C8C8C8", darkcolor="#C8C8C8",
                arrowsize=px(6), gripcount=0, borderwidth=0)
    s.map("Thin.Vertical.TScrollbar", background=[("active", "#A8A8A8")])
    s.layout("Thin.Horizontal.TScrollbar", [("Horizontal.Scrollbar.trough", {"sticky": "ew", "children": [
        ("Horizontal.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
    s.configure("Thin.Horizontal.TScrollbar", troughcolor=T["card"], background="#C8C8C8",
                bordercolor=T["card"], lightcolor="#C8C8C8", darkcolor="#C8C8C8",
                arrowsize=px(6), gripcount=0, borderwidth=0)
    s.map("Thin.Horizontal.TScrollbar", background=[("active", "#A8A8A8")])

    # Переключатели и флажки на карточке
    for w in ("TRadiobutton", "TCheckbutton"):
        s.configure(w, background=T["card"], foreground=T["text"], font=FONTS["body"],
                    indicatorbackground="#FFFFFF", indicatorforeground="#FFFFFF",
                    indicatormargin=(0, 0, px(6), 0), focuscolor=T["card"])
        s.map(w, indicatorbackground=[("selected", T["accent"]), ("active", T["subtle"])],
              background=[("active", T["card"])])


# ── Виджеты ─────────────────────────────────────────────────

def label(parent, text="", role="body", fg="text", bg="card", **kw):
    return tk.Label(parent, text=text, font=FONTS[role], fg=T.get(fg, fg), bg=T.get(bg, bg), **kw)


def card(parent, title=None, pad=16):
    """Белая скруглённая карточка. Возвращает (внешний ttk.Frame, тело tk.Frame)."""
    outer = ttk.Frame(parent, style="Card.TFrame", padding=px(1))
    body = tk.Frame(outer, bg=T["card"])
    body.pack(fill="both", expand=True, padx=px(pad - 1), pady=px(pad - 2))
    if title:
        label(body, title, "strong").pack(anchor="w", pady=(0, px(8)))
    return outer, body


def button(parent, text, command, kind="Std", on="card", icon=None, **kw):
    """kind: Accent | Std | Subtle; on: 'card' или 'bg' — фон под кнопкой."""
    style = f"{kind}{'' if on == 'card' else '.Bg'}.TButton"
    if icon:
        text = f"{icon}  {text}" if text else icon
    return ttk.Button(parent, text=text, command=command, style=style, cursor="hand2", **kw)


class InfoBar(tk.Frame):
    """Плашка статуса Fluent: цветной фон, иконка, текст. kind: ok|warn|err|info."""

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=T["info_bg"], highlightthickness=1,
                         highlightbackground=T["stroke"], **kw)
        self._icon = tk.Label(self, font=FONTS["icon_sm"], bg=T["info_bg"])
        self._icon.pack(side="left", padx=(px(14), px(10)), pady=px(10), anchor="n")
        self._text = tk.Label(self, font=FONTS["body"], bg=T["info_bg"], fg=T["text"],
                              justify="left", anchor="w")
        self._text.pack(side="left", fill="x", expand=True, pady=px(10), padx=(0, px(14)))

    def set(self, kind, text):
        bg = T[f"{kind}_bg"]
        self.configure(bg=bg)
        self._icon.configure(text=ICON[kind], fg=T[kind], bg=bg)
        self._text.configure(text=text, bg=bg)


class Expander(tk.Frame):
    """Сворачиваемый блок: заголовок-кнопка с шевроном и тело."""

    def __init__(self, parent, title, bg="bg", expanded=False, **kw):
        super().__init__(parent, bg=T[bg], **kw)
        self._title = title
        self._open = expanded
        self.header = tk.Frame(self, bg=T[bg], cursor="hand2")
        self.header.pack(fill="x")
        self._chev = tk.Label(self.header, font=FONTS["icon_sm"], bg=T[bg], fg=T["text2"])
        self._chev.pack(side="left")
        self._lbl = tk.Label(self.header, font=FONTS["strong"], bg=T[bg], fg=T["text"])
        self._lbl.pack(side="left", padx=(px(8), 0))
        for w in (self.header, self._chev, self._lbl):
            w.bind("<Button-1>", lambda e: self.toggle())
        self.body = tk.Frame(self, bg=T[bg])
        self._render()

    def set_title(self, title):
        self._title = title
        self._render()

    def toggle(self, open_=None):
        self._open = (not self._open) if open_ is None else open_
        self._render()

    def _render(self):
        self._chev.configure(text=ICON["chev_down"] if self._open else ICON["chev_right"])
        self._lbl.configure(text=self._title)
        if self._open:
            self.body.pack(fill="both", expand=True, pady=(px(8), 0))
        else:
            self.body.pack_forget()


class FluentDialog(tk.Toplevel):
    """
    Базовый модальный диалог: заголовок (Subtitle), подзаголовок, тело на
    сером фоне, футер с кнопками справа. Esc — закрыть.
    """

    def __init__(self, parent, title, subtitle="", size=None, minsize=None):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=T["bg"])
        if size:
            self.geometry(f"{px(size[0])}x{px(size[1])}")
        if minsize:
            self.minsize(px(minsize[0]), px(minsize[1]))
        head = tk.Frame(self, bg=T["bg"])
        head.pack(fill="x", padx=px(24), pady=(px(20), px(4)))
        label(head, title, "subtitle", bg="bg").pack(anchor="w")
        if subtitle:
            label(head, subtitle, "caption", fg="text2", bg="bg",
                  justify="left", wraplength=px((size or (640, 0))[0] - 48)).pack(anchor="w", pady=(px(2), 0))

        self.footer = tk.Frame(self, bg=T["card"])
        self.footer.pack(fill="x", side="bottom")
        tk.Frame(self.footer, bg=T["divider"], height=1).pack(fill="x", side="top")
        self._footer_inner = tk.Frame(self.footer, bg=T["card"])
        self._footer_inner.pack(fill="x", padx=px(24), pady=px(14))
        self.hint = label(self._footer_inner, "", "caption", fg="text2")
        self.hint.pack(side="left")

        self.body = tk.Frame(self, bg=T["bg"])
        self.body.pack(fill="both", expand=True, padx=px(24), pady=(px(12), px(16)))
        self.bind("<Escape>", lambda e: self.destroy())

    def add_button(self, text, command, kind="Std"):
        b = button(self._footer_inner, text, command, kind=kind)
        b.pack(side="right", padx=(px(8), 0))
        return b

    def show_modal(self, parent):
        self.transient(parent)
        style_window(self)
        self.grab_set()
        self.focus_set()
