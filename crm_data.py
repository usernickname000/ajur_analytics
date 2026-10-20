# ============================================================
# crm_data.py — загрузка и очистка выгрузки CRM
# Колонки, справочники-фильтры, разбор денег/дат, нормализация
# клиентов и отраслей, общие маски. Используется analytics,
# comparison и dashboard — чтобы все три считали одинаково.
# ============================================================

import json
import os
import re
from datetime import datetime

import pandas as pd

from paths import data_path


# ==============================
# ИМЕНА КОЛОНОК
# ==============================

COL_MONTH       = 'Месяц'
COL_ORDER       = 'Заказ'
COL_DATE        = 'Дата заказа'
COL_MANAGER     = 'Менеджер'
COL_REVENUE     = 'Выручка без НДС'
COL_DISCOUNT_PCT = 'Скидка, %'
COL_PROJECT     = 'Проект'
COL_BARTER = 'Бартер'
COL_IS_REVENUE = 'Это выручка'
COL_REVENUE_PAID = 'Выручка без НДС опл'

COL_CLIENT      = 'Клиент'
COL_CLIENT_RA   = 'Клиент РА'
COL_REKLAMD     = 'Рекламодатель'

COL_INDUSTRY_CLIENT_RA = 'Отрасли КлиентаРА'
COL_INDUSTRY_CLIENT    = 'Отрасли Клиента'
COL_INDUSTRY_ADV       = 'Отрасли Рекламодателя'

COL_DESCRIPTION = 'Описание'
COL_POS_ORDER   = 'Позиция заказа'
COL_NOMEN       = 'Номенклатура'


# ==============================
# ПАРАМЕТРЫ (можно редактировать)
# ==============================

EXCLUDE_MANAGERS = [
    'Алмазова А.', 'Василенко В.', 'Большакова М.',
    'Савиных Е.', 'Пирожкова Е.',
    'Селиванова А.', 'Дмитриева Д.', 'Матвеева В.'
]

EXCLUDE_PROJECTS = [
    'Мероприятия (деловые)',
    'Мероприятия (городские)'
]

# ── Периметр аналитики ─────────────────────────────────────
# Клиенты, которые целиком исключаются из расчётов (подстроки в верхнем
# регистре, ищутся в КОНЕЧНЫЙ_КЛИЕНТ):
#   ФОНТАНКА ФЕСТ — внутригрупповой оборот (взаимозачёт ФФ/АМ), включая мероприятия;
#   АНО МЕДИА ДОМ — вне расчёта (в бухгалтерии идёт строкой ИРИ/АНО).
# Строки с 'Бартер' = 'Да' тоже исключаются: это те же сделки, что и
# 'Медийный бартер' в бухгалтерии, и в сравнениях бартер не участвует.
EXCLUDE_CLIENTS_PERIMETER = {
    'ФОНТАНКА ФЕСТ': 'ФОНТАНКА ФЕСТ (взаимозачёт)',
    'АНО МЕДИА ДОМ': 'АНО Медиа ДОМ (вне расчёта)',
}
EXCLUDE_REASON_BARTER = 'Бартер'
COL_EXCLUDED = 'ИСКЛЮЧЕНО'

DISCOUNT_BUCKETS = [0, 5, 10, 15, 20, 25, 50]

REVENUE_BUCKETS = [
    ('Менее 10 тыс. руб.', 0, 10_000),
    ('10-50 тыс. руб.', 10_000, 50_000),
    ('50-100 тыс. руб.', 50_000, 100_000),
    ('100-500 тыс. руб.', 100_000, 500_000),
    ('500 тыс. - 1 млн руб.', 500_000, 1_000_000),
    ('Более 1 млн руб.', 1_000_000, float('inf'))
]


# ==============================
# АЛИАСЫ КЛИЕНТОВ
# ==============================

# Рядом с .exe в собранной программе (см. paths.py)
_ALIASES_PATH = data_path('client_aliases.json')
_aliases_cache = None


def set_aliases_path(path):
    """Путь к client_aliases.json (app.py задаёт тот же файл, что правит редактор)."""
    global _ALIASES_PATH
    _ALIASES_PATH = path
    reload_aliases()


def _load_aliases():
    if _aliases_cache is None:
        reload_aliases()
    return _aliases_cache


def reload_aliases():
    global _aliases_cache
    try:
        with open(_ALIASES_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # Паттерн ищется с начала слова (?<!\w): 'ГАЗПРОМ' ловит «Газпромбанк» и
        # «ПАО Газпром», но не «Запсибгазпром», где это часть другого слова.
        # upper() — на случай паттернов, введённых строчными буквами.
        _aliases_cache = [
            (g['canonical'],
             [re.compile(r'(?<!\w)' + re.escape(str(p).strip().upper()))
              for p in g.get('patterns', []) if str(p).strip()])
            for g in data.get('groups', [])
        ]
    except FileNotFoundError:
        _aliases_cache = []
    except Exception as e:
        print(f"⚠ client_aliases.json не прочитан ({e}) — нормализация клиентов отключена")
        _aliases_cache = []


# ==============================
# РАЗБОР ЗНАЧЕНИЙ
# ==============================

def parse_money(x):
    if pd.isna(x):
        return None
    x = str(x).strip()
    # Валюта вместе с точкой ('руб.', 'р.', '₽') и все пробелы, включая
    # неразрывные (\xa0,  ), которые Excel/1С ставят между разрядами.
    # Раньше точка из 'руб.' оставалась и '1 234,56 руб.' читалось как 123456.
    x = re.sub(r'(?i)(руб\.?|р\.|₽)', '', x)
    x = re.sub(r"[\s']", '', x)
    x = x.replace("−", "-").replace("–", "-")  # типографский минус
    has_comma = "," in x
    has_dot = "." in x
    if has_comma and has_dot:
        # Оба разделителя встретились — дробная часть там, где разделитель стоит
        # ПОСЛЕДНИМ в строке: '1.234,56' → ',' дробная (европ.), '1,234.56' → '.'
        # дробная (амер./英). Раньше любой такой случай трактовался как европейский
        # формат, что портило американские числа вида '1,234.56' на два порядка.
        if x.rfind(",") > x.rfind("."):
            x = x.replace(".", "").replace(",", ".")
        else:
            x = x.replace(",", "")
    elif has_comma:
        # Одна запятая — дробная часть (рус. формат); несколько — разделители тысяч
        x = x.replace(",", ".") if x.count(",") == 1 else x.replace(",", "")
    elif x.count(".") > 1:
        # '1.234.567' — точки как разделители тысяч
        x = x.replace(".", "")
    try:
        return float(x)
    except Exception:
        return None


def parse_month(s):
    try:
        if pd.isna(s):
            return None
        s = str(s).strip()
        if '.' in s:
            mm, yy = s.split('.')
            yy = int(yy)
            if yy < 100:
                yy = 2000 + yy
            return datetime(year=yy, month=int(mm), day=1)
        return pd.to_datetime(s, errors='coerce')
    except:
        return None


def parse_dates(series):
    """
    Разбирает колонку дат из выгрузки CRM. Даты приходят строками 'ДД.ММ.ГГГГ'
    вперемешку с датами Excel. Без явного формата pandas угадывает его по первой
    строке: '05.03.2025' читается как 3 мая, а все дни >12 превращаются в NaT.
    """
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    result = pd.to_datetime(series.where(series.map(lambda v: isinstance(v, str))).str.strip(),
                            format='%d.%m.%Y', errors='coerce')
    rest = result.isna() & series.notna()
    if rest.any():
        result.loc[rest] = series.loc[rest].map(
            lambda v: pd.to_datetime(v, dayfirst=True, errors='coerce'))
    return pd.to_datetime(result, errors='coerce')


def categorize_revenue_amount(v):
    if pd.isna(v):
        return 'Нет данных'
    try:
        v = float(v)
    except:
        return 'прочие'
    for name, lo, hi in REVENUE_BUCKETS:
        if lo <= v < hi:
            return name
    return 'прочие'


def bucket_discount(x):
    if pd.isna(x):
        return 'без скидки'
    try:
        xv = float(x)
    except:
        return 'прочие'
    if xv == 0:
        return 'без скидки'
    if xv in DISCOUNT_BUCKETS:
        return f'{int(xv)}%'
    if 40 <= xv <= 75:
        return '40–75% (нестандарты)'
    return 'прочие'


# ==============================
# КЛИЕНТЫ И ОТРАСЛИ
# ==============================

def pick_client(row):
    for c in [COL_CLIENT_RA, COL_CLIENT, COL_REKLAMD]:
        if c in row and pd.notna(row[c]) and str(row[c]).strip() != '':
            return str(row[c]).strip()
    return ''


def normalize_client(name):
    if pd.isna(name):
        return name
    name_clean = str(name).upper()
    for canonical, patterns in _load_aliases():
        if any(p.search(name_clean) for p in patterns):
            return canonical
    return name


def pick_industry(row):
    for c in [COL_INDUSTRY_CLIENT_RA, COL_INDUSTRY_CLIENT, COL_INDUSTRY_ADV]:
        if c in row and pd.notna(row[c]) and str(row[c]).strip() != '':
            return str(row[c]).strip()
    return ''


def _kw_regex(stems, words=()):
    """Регулярка: stems — с начала слова (основы), words — только целым словом."""
    parts = [re.escape(s) for s in stems] + [re.escape(w) + r'(?!\w)' for w in words]
    return re.compile(r'(?<!\w)(?:' + '|'.join(parts) + ')')


_EXPO_RE = _kw_regex(['выстав', 'expo', 'expо', 'стенд', 'форум', 'конгресс', 'ярмарк',
                      'фестиваль', 'ивент'], ['event', 'events'])
_ADS_RE = _kw_regex(['реклам', 'размещ', 'баннер', 'промо', 'digital', 'диджитал',
                     'программатик', 'programmatic', 'performance', 'перфоманс',
                     'таргет', 'target', 'охват', 'спецпроект', 'интеграц', 'медиа'],
                    ['cpm', 'cpc', 'cpp', 'media', 'показ', 'показы', 'показов'])
_MARKETING_RE = _kw_regex(['маркет', 'бренд', 'branding', 'пиар', 'контент',
                           'исслед', 'аналит', 'стратег', 'консалт', 'коммуникац'],
                          ['pr', 'smm'])


def classify_industry(row):
    """
    Отрасль клиента из CRM как есть. Текст заказа (позиция, описание, номенклатура)
    используется только чтобы разбить отрасль «Реклама, маркетинг, выставки»
    (или пустую) на Выставки / Реклама / Маркетинг.

    Раньше эвристика применялась ко всем отраслям, а почти любой заказ медиахолдинга
    содержит «размещение»/«баннер»/«реклама» — поэтому банки, издательства,
    производство и т.д. массово уходили в «Рекламу».
    """
    orig = row.get('ОТРАСЛЬ_КЛИЕНТА', '')
    raw = str(orig if pd.notna(orig) else '').strip().lower()

    if any(x in raw for x in ['строит', 'строй', 'недвиж']):
        return 'Строительство + Недвижимость'

    if raw in ['маркетинг', 'реклама', 'выставки']:
        return raw.capitalize()

    is_marketing_industry = any(x in raw for x in ['реклам', 'маркет', 'выстав'])
    if raw and not is_marketing_industry:
        return orig

    text = ' '.join(
        str(v) for v in (row.get(COL_POS_ORDER), row.get(COL_DESCRIPTION),
                         row.get(COL_NOMEN), row.get('КОНЕЧНЫЙ_КЛИЕНТ'))
        if v is not None and pd.notna(v)
    ).lower()

    if _EXPO_RE.search(text):
        return 'Выставки'
    if _ADS_RE.search(text):
        return 'Реклама'
    if _MARKETING_RE.search(text):
        return 'Маркетинг'
    return orig


# ==============================
# МАСКИ И ЗАГРУЗКА
# ==============================

def mask_no_events(df):
    """Строки без проектов-мероприятий (для среднего/медианного чека, см. README)."""
    if COL_PROJECT in df.columns:
        return ~df[COL_PROJECT].fillna('').isin(EXCLUDE_PROJECTS)
    return pd.Series(True, index=df.index)


def mask_no_prog(df):
    """Строки без CRM-программатика (по названию проекта)."""
    if COL_PROJECT in df.columns:
        return ~df[COL_PROJECT].fillna('').str.contains(
            'программатик|programmatic', case=False, regex=True)
    return pd.Series(True, index=df.index)


def client_rank_mask(df):
    """
    Строки для рейтингов клиентов (топ клиентов, RFM): без проектов-мероприятий
    и без клиентов-мероприятий (фестивали). Общая для отчёта, дашборда и
    сравнения, чтобы «Топ клиентов» везде совпадал.
    \\b — граница слова: без неё 'ФЕСТ' ловил бы «Манифест групп» или «Инфест».
    """
    is_event_client = (
        df['КОНЕЧНЫЙ_КЛИЕНТ'].astype(str).str.upper()
        .str.contains(r'\b(?:ФЕСТ|МЕРОПРИЯТИ)', na=False, regex=True)
    )
    if COL_PROJECT in df.columns:
        return ~(df[COL_PROJECT].fillna('').isin(EXCLUDE_PROJECTS) | is_event_client)
    return ~is_event_client


# Подписи второй строки шапки выгрузки → имя колонки, под которым она нужна в коде
_SUBHEADER_NAMES = {'Рекламодатель', 'Отрасли Рекламодателя'}


def load_crm(path, with_clients=True, required=(COL_ORDER, COL_REVENUE)):
    """
    Читает выгрузку «Заказы за период» с общей для всех отчётов очисткой:
    заголовки без пробелов, без строки «Итого» и пустых строк, выручка — числом.
    with_clients — добавить КОНЕЧНЫЙ_КЛИЕНТ (выбор поля + алиасы).

    Шапка выгрузки двухстрочная: под 'Договор РА' во второй строке стоят
    'Рекламодатель' и 'Отрасли Рекламодателя'. pandas читает только первую
    строку шапки, поэтому колонки получали бы имена 'Договор РА' и 'Unnamed: N',
    а сама вторая строка оставалась бы в данных как фиктивный заказ.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"Файл не найден: {path}")
    df = pd.read_excel(path, dtype=object)
    df.columns = df.columns.map(lambda x: x.strip() if isinstance(x, str) else x)

    for i in range(min(3, len(df))):
        row = df.iloc[i]
        vals = {col: str(v).strip() for col, v in row.items() if pd.notna(v)}
        if not any(v in _SUBHEADER_NAMES for v in vals.values()):
            continue
        if COL_ORDER in df.columns and pd.notna(row.get(COL_ORDER)):
            continue  # это обычная строка данных, а не подзаголовок
        rename = {col: v for col, v in vals.items() if v in _SUBHEADER_NAMES and v not in df.columns}
        df = df.rename(columns=rename).drop(index=df.index[i])
        break

    if COL_MONTH in df.columns:
        df = df.loc[~df[COL_MONTH].astype(str).str.strip().str.lower().eq("итого")]
    df = df.dropna(how="all").reset_index(drop=True)

    # Обязательные колонки проверяем ДО разбора — иначе вместо понятного
    # сообщения был бы KeyError
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Критические ошибки в данных:\n - Отсутствуют обязательные колонки: "
                         f"{missing} ({os.path.basename(path)})")
    df[COL_REVENUE] = df[COL_REVENUE].apply(parse_money)
    if with_clients:
        df['КОНЕЧНЫЙ_КЛИЕНТ'] = df.apply(pick_client, axis=1).apply(normalize_client)
    return df


def mark_perimeter(df):
    """
    Колонка COL_EXCLUDED: причина исключения строки из периметра аналитики
    ('' — в периметре). Сначала клиенты вне периметра, затем бартер
    остальных клиентов — чтобы сумма по причинам не задваивалась.
    """
    reason = pd.Series('', index=df.index, dtype=object)
    client = df['КОНЕЧНЫЙ_КЛИЕНТ'].astype(str).str.upper().str.replace(r'\s+', ' ', regex=True).str.strip()
    for pattern, label in EXCLUDE_CLIENTS_PERIMETER.items():
        reason.loc[client.str.contains(pattern, regex=False) & reason.eq('')] = label
    if COL_BARTER in df.columns:
        barter = df[COL_BARTER].astype(str).str.strip().str.lower().eq('да') & reason.eq('')
        reason.loc[barter] = EXCLUDE_REASON_BARTER
    df[COL_EXCLUDED] = reason
    return df


def excluded_summary(excluded):
    """Сводка исключённых строк по причинам: Причина / Строк / Выручка, тыс. руб."""
    rows = []
    for label in list(EXCLUDE_CLIENTS_PERIMETER.values()) + [EXCLUDE_REASON_BARTER]:
        part = excluded.loc[excluded[COL_EXCLUDED] == label]
        rows.append({'Причина исключения': label, 'Строк': len(part),
                     'Выручка, тыс. руб.': round(part[COL_REVENUE].sum() / 1000, 2)})
    return pd.DataFrame(rows)


def prepare_crm(df, log=None):
    """
    Конечный клиент + периметр аналитики. Возвращает (в периметре, исключённые):
      в периметре — без ФОНТАНКА ФЕСТ, АНО Медиа ДОМ, бартера и позиций
                    'Это выручка' = 'Нет' с пустой выручкой (это не заказы
                    с нулевой суммой, а нерекламные/технические позиции);
      исключённые — строки вне периметра с колонкой COL_EXCLUDED (причина).
    df[COL_REVENUE] уже должен быть числом.
    """
    df = df.copy()
    if 'КОНЕЧНЫЙ_КЛИЕНТ' not in df.columns:
        df['КОНЕЧНЫЙ_КЛИЕНТ'] = df.apply(pick_client, axis=1).apply(normalize_client)
    df = mark_perimeter(df)
    excluded = df.loc[df[COL_EXCLUDED] != ''].reset_index(drop=True)

    keep = df[COL_EXCLUDED].eq('')
    if COL_IS_REVENUE in df.columns:
        keep &= ~(df[COL_REVENUE].isna()
                  & df[COL_IS_REVENUE].astype(str).str.strip().str.lower().eq('нет'))
    if log is not None:
        for _, r in excluded_summary(excluded).iterrows():
            if r['Строк']:
                log(f"Вне периметра: {r['Причина исключения']} — {r['Строк']} строк, "
                    f"{r['Выручка, тыс. руб.']:,.0f} тыс. руб.")
    return df.loc[keep].reset_index(drop=True), excluded


def count_orders(df, revenue_col=COL_REVENUE):
    """
    Число заказов = уникальные 'Заказ' среди строк с выручкой. Строка выгрузки —
    позиция заказа за месяц, один заказ занимает много строк.
    """
    if COL_ORDER not in df.columns:
        return int(df[revenue_col].notna().sum())
    return int(df.loc[df[revenue_col].notna(), COL_ORDER].nunique())


def order_totals(df, revenue_col=COL_REVENUE):
    """Выручка по каждому заказу (заказ → сумма) — база для среднего/медианного чека."""
    d = df.loc[df[revenue_col].notna()]
    if COL_ORDER not in d.columns:
        return d[revenue_col].astype(float)
    return d.groupby(COL_ORDER)[revenue_col].sum()


# ==============================
# КАЧЕСТВО ДАННЫХ
# ==============================

def _order_dup_key(df):
    """
    Колонка для проверки дублей строк. 'Заказ' — идентификатор родительского
    договора и ЗАКОНОМЕРНО повторяется на много строк выгрузки (по одной строке
    на месяц/позицию контракта) — дедуп по нему на реальных данных даёт огромный
    процент ложных "дубликатов". 'Позиция заказа' — уникальный ключ строки
    (одна позиция = одна услуга за период), используем её, если она есть.
    """
    if COL_POS_ORDER in df.columns and df[COL_POS_ORDER].notna().any():
        return COL_POS_ORDER
    return COL_ORDER


def _order_dup_mask(df, keep):
    """
    Маска дублей по _order_dup_key(df). Пустое/пробельное значение ключа
    считается уникальным само по себе (не дублем других пустых значений) —
    иначе при частично заполненной 'Позиция заказа' все строки с пустым
    значением ложно попадали бы в дубли друг к другу.

    Пустые строки просто исключаются из сравнения (а не заменяются на
    синтетический ключ вроде индекса строки) — синтетический ключ мог бы
    случайно совпасть с настоящим значением 'Позиция заказа' (напр. если
    там встречаются короткие числовые id) и создать новый ложный дубль.
    """
    dup_col = _order_dup_key(df)
    raw = df[dup_col]
    key = raw.astype(str).str.strip()
    # raw.isna() — а не сравнение строки с 'nan' — потому что astype(str) в
    # некоторых версиях pandas оставляет NaN/None настоящим NaN, а не строкой.
    is_blank = raw.isna() | key.eq('')
    dup = pd.Series(False, index=df.index)
    non_blank = key.index[~is_blank]
    dup.loc[non_blank] = key.loc[non_blank].duplicated(keep=keep)
    return dup, dup_col


def run_data_quality_checks(df, log=print):
    report = {}
    critical_errors = []

    required_columns = [COL_ORDER, COL_DATE, COL_REVENUE]
    missing_cols = [c for c in required_columns if c not in df.columns]

    report['Всего строк'] = len(df)
    report['Отсутствующие обязательные колонки'] = ", ".join(missing_cols) if missing_cols else "Нет"

    if missing_cols:
        critical_errors.append(f"Отсутствуют обязательные колонки: {missing_cols}")

    if COL_ORDER in df.columns:
        dup_mask, dup_col = _order_dup_mask(df, keep='first')
        dup_label = 'Дубликатов позиций заказа' if dup_col == COL_POS_ORDER else 'Дубликатов заказов'
        report[dup_label] = dup_mask.sum()
    else:
        report['Дубликатов заказов'] = "Колонка отсутствует"

    if COL_REVENUE in df.columns:
        report['Отрицательная выручка'] = (df[COL_REVENUE] < 0).sum()
        report['Нулевая выручка']       = (df[COL_REVENUE] == 0).sum()
        missing = df[COL_REVENUE].isna()
        if COL_IS_REVENUE in df.columns:
            not_rev = df[COL_IS_REVENUE].astype(str).str.strip().str.lower().eq('нет')
            report["Позиций 'Это выручка' = 'Нет' (не участвуют)"] = int((missing & not_rev).sum())
            missing = missing & ~not_rev
        report['Пропущенная выручка']   = int(missing.sum())
        # Не критично: отрицательная выручка — обычно легитимные сторно/корректировки.
        # Раньше это останавливало весь отчёт; теперь просто предупреждаем и продолжаем,
        # как и с нулевой выручкой — строки остаются в расчётах (см. лист 00_Data_Quality
        # и 18б_CRM_проблемные_строки для деталей).
        if report['Отрицательная выручка'] > 0:
            log(f"⚠ Предупреждение: {report['Отрицательная выручка']} строк с отрицательной "
                f"выручкой (сторно/корректировки?) — включены в расчёты, не критично")

    if COL_DATE in df.columns:
        temp_dates = parse_dates(df[COL_DATE])
        report['Некорректные/пропущенные даты'] = temp_dates.isna().sum()
        report['Будущие даты'] = (temp_dates > pd.Timestamp.today()).sum()
        if report['Будущие даты'] > 0:
            log(f"⚠ Предупреждение: {report['Будущие даты']} дат в будущем (не критично)")

    if any(c in df.columns for c in [COL_CLIENT, COL_CLIENT_RA, COL_REKLAMD]):
        # Строки, где не заполнено НИ ОДНО поле клиента (раньше NaN превращался
        # в строку 'nan' и счётчик почти всегда показывал 0).
        client_cols = [c for c in [COL_CLIENT_RA, COL_CLIENT, COL_REKLAMD] if c in df.columns]
        filled = pd.concat(
            [df[c].notna() & (df[c].astype(str).str.strip() != '') for c in client_cols], axis=1
        ).any(axis=1)
        report['Строк без клиента'] = int((~filled).sum())

    if COL_REVENUE in df.columns:
        try:
            q99 = df[COL_REVENUE].quantile(0.99)
            report['Аномально большие чеки'] = (df[COL_REVENUE] > q99 * 3).sum()
        except Exception:
            report['Аномально большие чеки'] = "Ошибка расчета"

    quality_df = pd.DataFrame(list(report.items()), columns=['Метрика', 'Значение'])
    return quality_df, critical_errors
