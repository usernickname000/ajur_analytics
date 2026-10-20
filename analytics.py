# ============================================================
# analytics.py — ядро аналитики заказов
# Запускается из app.py: run_analytics() строит Excel-отчёт.
# Загрузка/очистка CRM — crm_data.py, внешние доходы —
# external_income.py, оформление Excel — excel_style.py.
# ============================================================

import json
import os

import numpy as np
import pandas as pd

import external_income as ext_income
from crm_data import (
    COL_BARTER, COL_DATE, COL_DISCOUNT_PCT, COL_MANAGER, COL_MONTH, COL_NOMEN, COL_ORDER,
    COL_POS_ORDER, COL_PROJECT, COL_REVENUE, EXCLUDE_MANAGERS,
    _order_dup_mask, bucket_discount, categorize_revenue_amount, classify_industry,
    client_rank_mask, load_crm, mask_no_events as _mask_no_events, mask_no_prog as _mask_no_prog,
    normalize_client, parse_dates, parse_money, parse_month, pick_client, pick_industry,
    run_data_quality_checks,
)
from excel_style import add_report_charts, autosize_columns, style_workbook

# ── Верифицированные цифры из бухгалтерии ──────────────────
# Загружаются из verified_figures.json при запуске.
# Константы-fallback на случай если JSON отсутствует:
VERIFIED_FIGURES_JSON = 'verified_figures.json'
VERIFIED_TOTAL_WITH_BARTER_NO_PROG = 363_000_166   # fallback
VERIFIED_ADVERTISING_NO_EVENTS     = 243_005_820   # fallback
VERIFIED_TOTAL_WITH_PROG           = 482_404_000   # fallback
EXTERNAL_PROGRAMMATIC_TOTAL        = 119_403_897   # fallback
EXTERNAL_OTHER_INCOME              = 41_057_102    # fallback

EXTERNAL_INCOME_JSON = ext_income.EXTERNAL_INCOME_JSON

# ── Маппинг строк бухгалтерии на CRM-проекты ────────────────────
# Формат: (название строки в бухгалтерии, список проектов CRM которые туда входят)
# Если список пуст — строка берётся только из external_income.json (внешние данные).
# Порядок строк совпадает со структурой Доходы_2025.xlsx.
ACCOUNTING_ROWS = [
    # ── Блок: Реклама ФОНТАНКА ──
    ('Fontanka.ru - баннерная реклама',       ['Фонтанка.ру: Баннерная реклама']),
    ('Fontanka.ru - мобильная реклама',       ['Фонтанка.ру: Мобильная версия']),
    ('Fontanka.ru - ТЕКСТЫ',                  ['Фонтанка.ру: Тексты']),
    ('IC Доходы ФОНТАНКА',                    ['IC Доходы Фонтанка']),
    ('Fontanka.ru - НАТИВ-спецпроекты',       ['Фонтанка.ру: Спецпроекты']),
    ('IC Доходы ФОНТАНКА спецпроекты',        ['IC Доходы Фонтанка спецпроекты']),
    ('СММ Фонтанка',                          ['Фонтанка.ру: СММ']),
    ('IC Доходы ФОНТАНКА СММ',                ['IC Доходы Фонтанка СММ']),
    ('Программатик ФОНТАНКА ТГ',              []),  # внешний
    ('Медийный бартер ФОНТАНКА',              []),  # внешний
    ('Программатик ФОНТАНКА',                 []),  # внешний
    ('E-com ФОНТАНКА',                        []),  # внешний
    ('Рекомендательные системы ФОНТАНКА',     []),  # внешний
    # ── Блок: Реклама ДОКТОР ──
    ('doctorpiter.ru - реклама в десктоп',    ['ДокторПитер.ру: Реклама (баннеры + тексты)']),
    ('СММ Доктор',                            ['doctorpiter.ru: СММ']),
    ('doctorpiter.ru - информ.услуги',        ['ДокторПитер.ру: Информ. услуги']),
    ('doctorpiter.ru - НАТИВ-спецпроекты',    ['ДокторПитер.ру: Спецпроекты']),
    ('IC Доходы ДОКТОР',                      ['IC Доходы Доктор', 'IC Доходы Доктор (спецпроекты)']),
    ('Программатик ДОКТОР ТГ',                []),  # внешний
    ('Программатик ДОКТОР',                   []),  # внешний
    ('E-com ДОКТОР',                          []),  # внешний
    ('Рекомендательные системы ДОКТОР',       []),  # внешний
    # ── Блок: Мероприятия ──
    ('Мероприятия Массовые (ЭВЕНТЫ)',         ['Мероприятия (городские)']),
    ('Медийный бартер ЭВЕНТЫ',                []),  # внешний
    ('IC Доходы ЭВЕНТЫ',                      ['IC Доходы Эвенты']),
    ('Мероприятия КС ФОНТАНКА',               ['Мероприятия (деловые)']),
    ('IC Доходы КС ФОНТАНКА',                 []),  # внешний — появилась в Доходы_2026.xlsx
    ('Мероприятия КС Медицина ДОКТОР',        []),  # в CRM нет отдельно
    # ── Блок: Прочие доходы ──
    ('Выручка 47 (в план)',                   []),  # внешний — в бух часть 47 идёт отдельно
    ('Выручка 47 (закупка)',                  []),  # внешний
    ('ФФ/АМ взаимозачет',                     []),  # внешний
    ('Гранты Фонтанка',                       []),  # внешний
    ('Гранты Доктор',                         []),  # внешний
    ('ИРИ/АНО/Петроцентр',                    []),  # внешний
    # ── Корректировки ──
    ('Взаимозачет/Затраты',                   []),  # внешний
    ('Корректировка скидки (комиссия ХШМ)',   []),  # внешний (отрицательные суммы)
]

# Группировка строк бухгалтерии в блоки для промежуточных итогов
ACCOUNTING_BLOCKS = {
    'Total Ads Sales': [
        'Fontanka.ru - баннерная реклама', 'Fontanka.ru - мобильная реклама',
        'Fontanka.ru - ТЕКСТЫ', 'IC Доходы ФОНТАНКА',
        'Fontanka.ru - НАТИВ-спецпроекты', 'IC Доходы ФОНТАНКА спецпроекты',
        'СММ Фонтанка', 'IC Доходы ФОНТАНКА СММ',
        'Программатик ФОНТАНКА ТГ', 'Медийный бартер ФОНТАНКА',
        'Программатик ФОНТАНКА', 'E-com ФОНТАНКА', 'Рекомендательные системы ФОНТАНКА',
        'doctorpiter.ru - реклама в десктоп', 'СММ Доктор',
        'doctorpiter.ru - информ.услуги', 'doctorpiter.ru - НАТИВ-спецпроекты',
        'IC Доходы ДОКТОР', 'Программатик ДОКТОР ТГ', 'Программатик ДОКТОР',
        'E-com ДОКТОР', 'Рекомендательные системы ДОКТОР',
    ],
    'Total Events Sales': [
        'Мероприятия Массовые (ЭВЕНТЫ)', 'Медийный бартер ЭВЕНТЫ',
        'IC Доходы ЭВЕНТЫ', 'Мероприятия КС ФОНТАНКА', 'IC Доходы КС ФОНТАНКА',
        'Мероприятия КС Медицина ДОКТОР',
    ],
    'Total Прочие доходы': [
        'Выручка 47 (в план)', 'Выручка 47 (закупка)', 'ФФ/АМ взаимозачет',
        'Гранты Фонтанка', 'Гранты Доктор', 'ИРИ/АНО/Петроцентр',
    ],
}

# Вычеты из Total Commercial Sales → «Итого в бюджет отдела рекламы»
BUDGET_DEDUCTION_ROWS = [
    'Медийный бартер ФОНТАНКА', 'Медийный бартер ЭВЕНТЫ',
    'Взаимозачет/Затраты', 'ФФ/АМ взаимозачет',
    'Программатик ФОНТАНКА', 'Программатик ФОНТАНКА ТГ',
    'Программатик ДОКТОР', 'Программатик ДОКТОР ТГ',
    'Корректировка скидки (комиссия ХШМ)', 'ИРИ/АНО/Петроцентр',
    'Выручка 47 (закупка)',
]

# Маппинг направлений для KPI
KPI_DIRECTIONS = {
    'Фонтанка': [
        'Fontanka.ru - баннерная реклама', 'Fontanka.ru - мобильная реклама',
        'Fontanka.ru - ТЕКСТЫ', 'IC Доходы ФОНТАНКА',
        'Fontanka.ru - НАТИВ-спецпроекты', 'IC Доходы ФОНТАНКА спецпроекты',
    ],
    'Доктор': [
        'doctorpiter.ru - реклама в десктоп', 'СММ Доктор',
        'doctorpiter.ru - информ.услуги', 'doctorpiter.ru - НАТИВ-спецпроекты',
        'IC Доходы ДОКТОР',
    ],
    'СММ': ['СММ Фонтанка', 'IC Доходы ФОНТАНКА СММ'],
    'Мероприятия + КС': [
        'Мероприятия Массовые (ЭВЕНТЫ)', 'IC Доходы ЭВЕНТЫ',
        'Мероприятия КС ФОНТАНКА', 'Мероприятия КС Медицина ДОКТОР',
    ],
}

# ── Маппинг проектов CRM → бизнес-группа ────────────────────────
# Структура повторяет логику Доходы_2025.xlsx: Ads / Events / Прочее / Программатик.
# Если появится новый проект — добавь его сюда, иначе он попадёт в "НЕ КЛАССИФИЦИРОВАНО".
PROJECT_GROUP_MAP = {
    # ── Реклама Фонтанка ──
    'Фонтанка.ру: Тексты':              'Реклама Фонтанка',
    'Фонтанка.ру: СММ':                 'Реклама Фонтанка',
    'Фонтанка.ру: Спецпроекты':         'Реклама Фонтанка',
    'Фонтанка.ру: Баннерная реклама':   'Реклама Фонтанка',
    'Фонтанка.ру: Мобильная версия':    'Реклама Фонтанка',
    'Фонтанка.ру':                      'Реклама Фонтанка',
    'IC Доходы Фонтанка':               'Реклама Фонтанка',
    'IC Доходы Фонтанка спецпроекты':   'Реклама Фонтанка',
    'IC Доходы Фонтанка СММ':           'Реклама Фонтанка',
    # ── Реклама Доктор ──
    'ДокторПитер.ру: Реклама (баннеры + тексты)': 'Реклама Доктор',
    'ДокторПитер.ру: Спецпроекты':      'Реклама Доктор',
    'doctorpiter.ru: СММ':              'Реклама Доктор',
    'ДокторПитер.ру: Информ. услуги':   'Реклама Доктор',
    'IC Доходы Доктор':                 'Реклама Доктор',
    'IC Доходы Доктор (спецпроекты)':   'Реклама Доктор',
    # ── Программатик (CRM-часть, малая) ──
    'Программатик Фонтанка':            'Программатик',
    'Программатик ДокторПитер':         'Программатик',
    # ── Мероприятия ──
    'Мероприятия (деловые)':            'Мероприятия',
    'Мероприятия (городские)':          'Мероприятия',
    'IC Доходы Эвенты':                 'Мероприятия',
    # ── 47News / Прочее ──
    '47News':                           '47News / Прочее',
    '47News: Прочие услуги':            '47News / Прочее',
}

MONTH_TO_QUARTER = {f"{m:02d}": (m - 1) // 3 + 1 for m in range(1, 13)}
MONTH_TO_SEASON = {
    '12': 'Зима', '01': 'Зима', '02': 'Зима',
    '03': 'Весна', '04': 'Весна', '05': 'Весна',
    '06': 'Лето', '07': 'Лето', '08': 'Лето',
    '09': 'Осень', '10': 'Осень', '11': 'Осень',
}

# Пороги отклонения сверки, %: < WARN — OK, < BAD — проверить, иначе критично
RECON_WARN_PCT = 2
RECON_BAD_PCT = 5


# ==============================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ==============================

def _rev_stats(df, groupby_col, revenue_col, avg_mask=None):
    """Groupby с суммой/количеством/средним чеком и переводом в тыс. руб."""
    grp = df.groupby(groupby_col)
    stats = grp[revenue_col].agg(['sum', 'count']).round(2)
    stats.columns = ['Сумма выручки, руб.', 'Количество заказов']
    avg_df = df if avg_mask is None else df.loc[avg_mask]
    stats['Средний чек, руб.'] = avg_df.groupby(groupby_col)[revenue_col].mean().round(2)
    stats['Сумма выручки, тыс. руб.'] = (stats['Сумма выручки, руб.'] / 1000).round(2)
    return stats.reset_index().sort_values('Сумма выручки, тыс. руб.', ascending=False)


def load_verified_figures(path):
    """
    Загружает verified_figures.json. Возвращает dict с ключами:
    total_with_prog, total_with_barter_no_prog, advertising_no_events,
    programmatic_external, other_external_income, _год.
    Если файла нет — возвращает None.
    """
    if not os.path.exists(path):
        return None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None


def _pct_delta(fact, target):
    """Возвращает отклонение fact от target в процентах или None."""
    if target is None or target == 0:
        return None
    return (fact / target - 1) * 100


def _fmt_pct(pct):
    """'+1.23%' или '—'."""
    return '—' if pct is None else f"{round(pct, 2):+.2f}%"


def _recon_status(delta_pct, warn=RECON_WARN_PCT, bad=RECON_BAD_PCT):
    """Классифицирует качество сверки по процентному отклонению."""
    if delta_pct is None:
        return 'нет базы'
    abs_delta = abs(delta_pct)
    if abs_delta < warn:
        return 'OK'
    if abs_delta < bad:
        return 'Проверить'
    return 'Критично'


def recon_verdict(delta_pct):
    """Словесная оценка отклонения для журнала и GUI."""
    a = abs(delta_pct)
    return 'отлично' if a < RECON_WARN_PCT else ('в норме' if a < RECON_BAD_PCT else 'большое расхождение')


def _json_candidates(explicit, filename, input_path):
    """
    Порядок поиска JSON-справочника: явный путь (из GUI) → папка приложения
    (её правят редакторы) → папка входного файла → текущий каталог.
    Раньше папка приложения шла последней, и копия рядом с выгрузкой молча
    перекрывала правки, сделанные в редакторе.
    """
    return [p for p in (
        explicit,
        os.path.join(os.path.dirname(os.path.abspath(__file__)), filename),
        os.path.join(os.path.dirname(input_path), filename),
        filename,
    ) if p]


def _find_existing(candidates):
    return next((p for p in candidates if os.path.exists(p)), None)


# ==============================
# СВЕРКА И КОНТРОЛЬ
# ==============================

def build_unclassified_projects_report(df_full, revenue_col):
    """Возвращает сводку проектов, которые не попали в PROJECT_GROUP_MAP."""
    if COL_PROJECT not in df_full.columns or 'БИЗНЕС_ГРУППА' not in df_full.columns:
        return pd.DataFrame(columns=['Проект CRM', 'Заказов', 'Выручка, тыс. руб.', 'Доля CRM, %'])

    total_revenue = df_full[revenue_col].sum()
    unclassified = df_full.loc[df_full['БИЗНЕС_ГРУППА'] == 'НЕ КЛАССИФИЦИРОВАНО']
    if unclassified.empty:
        return pd.DataFrame([{
            'Проект CRM': 'Все проекты классифицированы',
            'Заказов': 0,
            'Выручка, тыс. руб.': 0,
            'Доля CRM, %': 0,
        }])

    report = (
        unclassified.groupby(COL_PROJECT, dropna=False)
        .agg(Заказов=(COL_ORDER, 'count'), Выручка_руб=(revenue_col, 'sum'))
        .reset_index()
        .rename(columns={COL_PROJECT: 'Проект CRM'})
    )
    report['Проект CRM'] = report['Проект CRM'].fillna('Пустой проект')
    report['Выручка, тыс. руб.'] = (report['Выручка_руб'] / 1000).round(2)
    report['Доля CRM, %'] = (
        report['Выручка_руб'] / total_revenue * 100 if total_revenue else 0
    ).round(2)
    return report.drop(columns=['Выручка_руб']).sort_values('Выручка, тыс. руб.', ascending=False)


def build_crm_issue_rows(df_full, revenue_col):
    """Формирует список CRM-строк, которые могут искажать итоговую сверку."""
    issue_frames = []
    cols = [
        col for col in [
            COL_ORDER, COL_POS_ORDER, COL_MONTH, COL_DATE, 'Дата_оплаты', COL_PROJECT,
            'БИЗНЕС_ГРУППА', COL_MANAGER, 'КОНЕЧНЫЙ_КЛИЕНТ', revenue_col,
            COL_BARTER, COL_DISCOUNT_PCT
        ]
        if col in df_full.columns
    ]

    def add_issue(mask, issue):
        if not mask.any():
            return
        part = df_full.loc[mask, cols].copy()
        part.insert(0, 'Проблема', issue)
        issue_frames.append(part)

    def blank(col):
        return df_full[col].isna() | df_full[col].astype(str).str.strip().eq('')

    add_issue(df_full[revenue_col].isna(), 'Пустая/нечисловая выручка')
    add_issue(df_full[revenue_col].fillna(0).eq(0), 'Нулевая выручка')
    add_issue(df_full[revenue_col].fillna(0).lt(0), 'Отрицательная выручка')

    if 'Дата_заказа' in df_full.columns:
        add_issue(df_full['Дата_заказа'].isna(), 'Некорректная дата заказа')
    if 'Дата_оплаты' in df_full.columns:
        add_issue(df_full['Дата_оплаты'].isna(), 'Нет даты оплаты')
    if COL_PROJECT in df_full.columns:
        add_issue(blank(COL_PROJECT), 'Пустой проект')
    if 'БИЗНЕС_ГРУППА' in df_full.columns:
        add_issue(df_full['БИЗНЕС_ГРУППА'].eq('НЕ КЛАССИФИЦИРОВАНО'), 'Проект не классифицирован')
    if 'КОНЕЧНЫЙ_КЛИЕНТ' in df_full.columns:
        add_issue(blank('КОНЕЧНЫЙ_КЛИЕНТ'), 'Пустой конечный клиент')
    if COL_ORDER in df_full.columns:
        dup_mask, dup_col = _order_dup_mask(df_full, keep=False)
        dup_issue = 'Дубликат позиции заказа' if dup_col == COL_POS_ORDER else 'Дубликат номера заказа'
        add_issue(dup_mask, dup_issue)

    if not issue_frames:
        return pd.DataFrame([{'Проблема': 'Явных проблемных строк не найдено'}])

    issues = pd.concat(issue_frames, ignore_index=True)
    if revenue_col in issues.columns:
        issues = issues.sort_values(revenue_col, ascending=False, na_position='last')
    return issues.head(500)


def build_reconciliation_bridge(
    rev_all_k,
    rev_bez_prog_k,
    rev_reklama_k,
    crm_prog_k,
    ext_cats,
    vf_total_with_prog_k,
    vf_total_no_prog_k,
    vf_ads_no_events_k,
):
    """Строит понятный мост от CRM-итога к бухгалтерским срезам."""
    ext_prog_k = ext_cats.get('programmatic', 0) / 1000
    ext_barter_k = ext_cats.get('barter', 0) / 1000
    ext_income_k = ext_cats.get('total_income', 0) / 1000
    ext_deduct_k = ext_cats.get('deductions', 0) / 1000

    rows = [
        ('CRM: все строки выгрузки', rev_all_k, '', '', ''),
        ('минус CRM-программатик', -crm_prog_k, '', '', ''),
        ('CRM без программатика', rev_bez_prog_k, '', '', ''),
        ('плюс внешний программатик', ext_prog_k, '', '', ''),
        ('плюс внешний бартер', ext_barter_k, '', '', ''),
        ('плюс прочие внешние доходы', ext_income_k - ext_prog_k - ext_barter_k, '', '', ''),
        ('справочно: внешние вычеты не в grand total', ext_deduct_k, '', '', ''),
    ]

    scenarios = [
        ('Итого: CRM + все внешние доходы', rev_all_k + ext_income_k, vf_total_with_prog_k),
        ('Итого: CRM без прогр. + бартер', rev_bez_prog_k + ext_barter_k, vf_total_no_prog_k),
        ('Итого: реклама без мероприятий', rev_reklama_k, vf_ads_no_events_k),
    ]
    for name, fact, target in scenarios:
        delta = fact - target if target else None
        pct = _pct_delta(fact, target)
        rows.append((name, fact, target or '', delta if delta is not None else '', _recon_status(pct)))

    return pd.DataFrame(
        rows,
        columns=['Шаг / срез', 'Факт, тыс. руб.', 'Бухгалтерия, тыс. руб.', 'Разница, тыс. руб.', 'Статус']
    )


def build_payment_order_reconciliation(monthly_stats):
    """Сравнивает месячный факт по дате заказа и по дате оплаты."""
    if monthly_stats is None or 'Выручка по дате оплаты, тыс. руб.' not in monthly_stats.columns:
        return None

    result = monthly_stats[['Период', 'Сумма выручки, тыс. руб.', 'Выручка по дате оплаты, тыс. руб.']].copy()
    result = result.rename(columns={
        'Сумма выручки, тыс. руб.': 'CRM по дате заказа, тыс. руб.',
        'Выручка по дате оплаты, тыс. руб.': 'CRM по дате оплаты, тыс. руб.',
    })
    result['Разница оплата-заказ, тыс. руб.'] = (
        result['CRM по дате оплаты, тыс. руб.'] - result['CRM по дате заказа, тыс. руб.']
    ).round(2)
    result['Разница, %'] = result.apply(
        lambda row: round(
            row['Разница оплата-заказ, тыс. руб.'] / row['CRM по дате заказа, тыс. руб.'] * 100, 2
        ) if row['CRM по дате заказа, тыс. руб.'] else None,
        axis=1,
    )
    return result


def build_accounting_table(df_full, revenue_col, ext, log=print):
    """
    Строит бухгалтерскую таблицу в стиле Доходы_2025.xlsx.
    ext — содержимое external_income.json. Возвращает DataFrame.
    """
    months = ext_income.MONTHS
    analysis_year = int(ext.get('_год', 2025))
    month_labels = {m: f"{m}.{analysis_year}" for m in months}

    # Помесячная выручка CRM по проектам (в тыс. руб.)
    df_full = df_full.copy()
    if COL_MONTH in df_full.columns:
        _ms = df_full[COL_MONTH].astype(str).str.strip()
        df_full['_month_num'] = _ms.str.split('.').str[0].str.zfill(2)
        # Год: формат 'MM.YY' или 'MM.YYYY'
        _yr = pd.to_numeric(_ms.str.split('.').str[1], errors='coerce').fillna(0).astype(int)
        df_full['_year_num'] = _yr.apply(lambda y: (2000 + y) if (0 < y < 100) else y)
    else:
        df_full['_month_num'] = '00'
        df_full['_year_num'] = 0
    year_mask = df_full['_year_num'] == analysis_year

    def row(name, values, total=None):
        return {'Показатель': name, **values,
                'Total': round(sum(values.values()), 2) if total is None else total}

    def sum_rows(names):
        return {m: sum(row_totals[rn].get(m, 0) for rn in names if rn in row_totals) for m in months}

    rows = []
    row_totals = {}  # для подсчётов блоков

    for row_name, crm_projects in ACCOUNTING_ROWS:
        month_values = {}
        total = 0.0
        for m in months:
            val = 0.0
            # 1. CRM-часть (только нужный год, чтобы не суммировать несколько лет)
            if crm_projects:
                mask = df_full[COL_PROJECT].isin(crm_projects) & (df_full['_month_num'] == m) & year_mask
                val += df_full.loc[mask, revenue_col].sum() / 1000  # в тыс. руб.
            # 2. Внешняя часть (из JSON)
            if row_name in ext and isinstance(ext[row_name], dict):
                val += (ext[row_name].get(m, 0) or 0) / 1000  # из руб. в тыс.
            month_values[m] = round(val, 2)
            total += val
        rows.append(row(row_name, month_values, round(total, 2)))
        row_totals[row_name] = month_values

    # Промежуточные итоги (Total Ads / Events / Прочие)
    for block_name, row_names in ACCOUNTING_BLOCKS.items():
        block_monthly = sum_rows(row_names)
        rows.append(row(block_name, {m: round(v, 2) for m, v in block_monthly.items()},
                        round(sum(block_monthly.values()), 2)))
        row_totals[block_name] = block_monthly

    # Total Commercial Sales = Ads + Events + Прочие
    commercial_monthly = sum_rows(['Total Ads Sales', 'Total Events Sales', 'Total Прочие доходы'])
    rows.append(row('Total Commercial Sales', {m: round(v, 2) for m, v in commercial_monthly.items()},
                    round(sum(commercial_monthly.values()), 2)))

    rows.append({'Показатель': '── Вычеты из бюджета ──', **{m: '' for m in months}, 'Total': ''})

    # Итог в бюджет отдела. Корректировка скидки хранится в JSON со знаком
    # минус — вычитаем её модуль, иначе «минус минус» увеличивал бюджет.
    budget_monthly = dict(commercial_monthly)
    for rn in BUDGET_DEDUCTION_ROWS:
        if rn in row_totals:
            for m in months:
                v = row_totals[rn].get(m, 0)
                budget_monthly[m] -= abs(v) if rn == 'Корректировка скидки (комиссия ХШМ)' else v
    budget_total = sum(budget_monthly.values())
    rows.append(row('Итого в бюджет отдела рекламы',
                    {m: round(v, 2) for m, v in budget_monthly.items()}, round(budget_total, 2)))

    # План группы и выполнение
    plan_monthly = ext.get('_план_группы_по_месяцам_руб', {})
    if plan_monthly:
        plan_row = {m: round(plan_monthly.get(m, 0) / 1000, 2) for m in months}
        plan_total = sum(plan_row.values())
        rows.append(row('План группы', plan_row, round(plan_total, 2)))
        perf_row = {m: round(budget_monthly.get(m, 0) / plan_row[m] * 100, 1) if plan_row[m] else None
                    for m in months}
        rows.append({'Показатель': 'Выполнение группы, %', **perf_row,
                     'Total': round(budget_total / plan_total * 100, 1) if plan_total else None})

    # KPI по направлениям
    kpi_plans = ext.get('_план_KPI_по_месяцам_руб', {})
    if kpi_plans and any(
        any(v for v in kpi_plans.get(d, {}).values() if isinstance(v, (int, float)) and v > 0)
        for d in KPI_DIRECTIONS
    ):
        rows.append({'Показатель': '── KPI ──', **{m: '' for m in months}, 'Total': ''})
        for direction, row_names in KPI_DIRECTIONS.items():
            fact_monthly = sum_rows(row_names)
            plan_dir = kpi_plans.get(direction, {})
            kpi_row = {}
            for m in months:
                p = (plan_dir.get(m, 0) or 0) / 1000
                kpi_row[m] = round(fact_monthly[m] / p * 100, 1) if p else None
            plan_total_dir = sum((plan_dir.get(m, 0) or 0) / 1000 for m in months)
            fact_total_dir = sum(fact_monthly.values())
            rows.append({'Показатель': f'KPI {direction}, %', **kpi_row,
                         'Total': round(fact_total_dir / plan_total_dir * 100, 1) if plan_total_dir else None})

    # Контроль: CRM-выручка отчётного года, не попавшая ни в одну строку таблицы
    # (проект отсутствует в ACCOUNTING_ROWS) — иначе она теряется молча.
    mapped_projects = {p for _, projects in ACCOUNTING_ROWS for p in projects}
    if COL_PROJECT in df_full.columns:
        unmapped = df_full.loc[year_mask & ~df_full[COL_PROJECT].isin(mapped_projects)]
        unmapped_monthly = {
            m: round(unmapped.loc[unmapped['_month_num'] == m, revenue_col].sum() / 1000, 2)
            for m in months
        }
        rows.append({'Показатель': '── Справочно ──', **{m: '' for m in months}, 'Total': ''})
        rows.append(row('CRM вне строк таблицы (проверить маппинг)', unmapped_monthly))

    columns = ['Показатель'] + [month_labels[m] for m in months] + ['Total']
    return pd.DataFrame([
        {'Показатель': r['Показатель'], **{month_labels[m]: r.get(m, '') for m in months}, 'Total': r.get('Total', '')}
        for r in rows
    ], columns=columns)


# ==============================
# СИГНАЛЫ
# ==============================

def _closed_months(monthly_stats):
    """
    Месячная статистика без текущего (незакрытого) и будущих месяцев,
    с колонкой '_ord' = год*12+месяц (для тренда с учётом пропущенных месяцев).
    Иначе неполный текущий месяц выглядит как «падение» и тянет прогноз вниз.
    """
    if monthly_stats is None or monthly_stats.empty:
        return monthly_stats
    dt = pd.to_datetime(monthly_stats['Период'], format='%m.%Y', errors='coerce')
    month_start = pd.Timestamp.today().normalize().replace(day=1)
    closed = dt < month_start
    ms = monthly_stats.loc[closed].copy()
    ms['_ord'] = dt[closed].dt.year * 12 + dt[closed].dt.month
    return ms.reset_index(drop=True)


def build_signals(df_full, revenue_col, monthly_stats, manager_stats, client_stats,
                  rfm_all, mask_no_events, log=print, client_mask=None):
    """
    Автоматически находит аномалии и сигналы в данных.
    Возвращает DataFrame с колонками: Приоритет, Категория, Сигнал, Детали.
    """
    signals = []

    def add(priority, category, signal, details=''):
        signals.append({'Приоритет': priority, 'Категория': category,
                        'Сигнал': signal, 'Детали': details})

    # ── 1. Аномалии по месяцам ───────────────────────────────
    monthly_stats = _closed_months(monthly_stats)
    if monthly_stats is not None and len(monthly_stats) >= 3:
        rev = monthly_stats['Сумма выручки, тыс. руб.']
        mean_rev = rev.mean()
        std_rev  = rev.std()
        periods  = monthly_stats['Период'].tolist()

        # При n<=5 отклонение одной точки от среднего математически не может
        # превысить 2σ (выборочное std) — проверка имеет смысл от 6 месяцев.
        if len(rev) >= 6:
            for period, val in zip(periods, rev):
                if std_rev > 0 and abs(val - mean_rev) > 2 * std_rev:
                    direction = '🔺 аномально высокая' if val > mean_rev else '🔻 аномально низкая'
                    add('🔴 Высокий', 'Месячная выручка',
                        f'{period}: {direction} выручка',
                        f'{val:,.0f} тыс. vs среднее {mean_rev:,.0f} тыс. (±{std_rev:,.0f})')

        # Падение три месяца подряд (только смежные месяцы, без пропусков)
        ords = monthly_stats['_ord'].tolist()
        if len(rev) >= 3 and ords[-1] - ords[-3] == 2:
            last3 = rev.iloc[-3:].tolist()
            if last3[2] < last3[1] < last3[0]:
                add('🟡 Средний', 'Тренд',
                    'Выручка падает три месяца подряд',
                    f'{periods[-3]}: {last3[0]:,.0f} → {periods[-2]}: {last3[1]:,.0f} → {periods[-1]}: {last3[2]:,.0f} тыс.')

        # Лучший и худший месяц
        idx_max = rev.idxmax()
        idx_min = rev.idxmin()
        add('🟢 Инфо', 'Месячная выручка',
            f'Лучший месяц: {monthly_stats.loc[idx_max, "Период"]}',
            f'{rev[idx_max]:,.0f} тыс. руб.')
        add('🟢 Инфо', 'Месячная выручка',
            f'Худший месяц: {monthly_stats.loc[idx_min, "Период"]}',
            f'{rev[idx_min]:,.0f} тыс. руб.')

    # ── 2. Концентрация выручки (топ-клиенты) ────────────────
    if client_stats is not None and len(client_stats) > 0:
        total_rev = client_stats['Сумма выручки, тыс. руб.'].sum()
        if total_rev > 0:
            top1_rev   = client_stats.iloc[0]['Сумма выручки, тыс. руб.']
            top1_share = top1_rev / total_rev * 100
            top5_share = client_stats.head(5)['Сумма выручки, тыс. руб.'].sum() / total_rev * 100
            top1_name  = client_stats.iloc[0]['КОНЕЧНЫЙ_КЛИЕНТ']

            if top1_share > 20:
                add('🔴 Высокий', 'Концентрация',
                    f'Один клиент даёт {top1_share:.1f}% выручки — высокий риск',
                    f'{top1_name}: {top1_rev:,.0f} тыс.')
            elif top1_share > 10:
                add('🟡 Средний', 'Концентрация',
                    f'Топ-клиент даёт {top1_share:.1f}% выручки',
                    f'{top1_name}: {top1_rev:,.0f} тыс.')

            if top5_share > 60:
                add('🟡 Средний', 'Концентрация',
                    f'Топ-5 клиентов дают {top5_share:.1f}% выручки',
                    ', '.join(client_stats.head(5)['КОНЕЧНЫЙ_КЛИЕНТ'].tolist()))

    # ── 3. Клиенты с одним заказом (риск оттока) ─────────────
    # Считаем уникальные заказы, а не строки выгрузки (у одного заказа много
    # строк-позиций), и по тем же клиентам, что и топ клиентов (без мероприятий).
    if 'КОНЕЧНЫЙ_КЛИЕНТ' in df_full.columns:
        src = df_full if client_mask is None else df_full.loc[client_mask]
        order_counts = src.groupby('КОНЕЧНЫЙ_КЛИЕНТ').agg(
            count=(COL_ORDER if COL_ORDER in src.columns else revenue_col, 'nunique'),
            sum=(revenue_col, 'sum'))
        single_order = order_counts[order_counts['count'] == 1]
        total_sum = order_counts['sum'].sum()
        single_rev_share = single_order['sum'].sum() / total_sum * 100 if total_sum > 0 else 0
        add('🟡 Средний' if single_rev_share > 30 else '🟢 Инфо', 'Лояльность',
            f'{len(single_order)} клиентов с единственным заказом ({single_rev_share:.1f}% выручки)',
            f'Потенциал повторных продаж: {single_order["sum"].sum()/1000:,.0f} тыс. руб.')

    # ── 4. RFM-сигналы ────────────────────────────────────────
    if rfm_all is not None and len(rfm_all) > 0:
        at_risk = rfm_all[rfm_all['Segment'] == 'At Risk']
        hibernating = rfm_all[rfm_all['Segment'] == 'Hibernating']

        if len(at_risk) > 0:
            add('🔴 Высокий', 'RFM / Отток',
                f'{len(at_risk)} клиентов "At Risk" — давно не покупали',
                f'Суммарная ценность: {at_risk["Monetary_thousands"].sum():,.0f} тыс. руб. | Топ: ' +
                ', '.join(at_risk.head(3)['КОНЕЧНЫЙ_КЛИЕНТ'].tolist()))

        if len(hibernating) > 0:
            add('🟡 Средний', 'RFM / Отток',
                f'{len(hibernating)} клиентов "Hibernating" — неактивны',
                f'Суммарная ценность: {hibernating["Monetary_thousands"].sum():,.0f} тыс. руб.')

        champions = rfm_all[rfm_all['Segment'] == 'Champions']
        if len(champions) > 0:
            add('🟢 Инфо', 'RFM / Рост',
                f'{len(champions)} клиентов-чемпионов — активно покупают',
                ', '.join(champions.head(5)['КОНЕЧНЫЙ_КЛИЕНТ'].tolist()))

    # ── 5. Менеджеры ─────────────────────────────────────────
    if manager_stats is not None and len(manager_stats) >= 2:
        top_mgr = manager_stats.iloc[0]
        bot_mgr = manager_stats.iloc[-1]
        top_rev = top_mgr['Сумма выручки, тыс. руб.']
        bot_rev = bot_mgr['Сумма выручки, тыс. руб.']
        gap = top_rev / bot_rev if bot_rev > 0 else 0

        if gap > 5:
            add('🟡 Средний', 'Менеджеры',
                f'Разрыв между топ и аутсайдером в {gap:.1f}x',
                f'{top_mgr[COL_MANAGER]}: {top_rev:,.0f} тыс. vs '
                f'{bot_mgr[COL_MANAGER]}: {bot_rev:,.0f} тыс.')

        add('🟢 Инфо', 'Менеджеры',
            f'Лучший менеджер: {top_mgr[COL_MANAGER]}',
            f'{top_rev:,.0f} тыс. руб.')

    # ── 6. Средний чек ────────────────────────────────────────
    rev_no_events = df_full.loc[mask_no_events, revenue_col].dropna()
    if len(rev_no_events) > 0:
        avg_check = rev_no_events.mean() / 1000
        median_check = rev_no_events.median() / 1000
        ratio = avg_check / median_check if median_check > 0 else 1

        if ratio > 3:
            add('🟡 Средний', 'Чеки',
                f'Средний чек в {ratio:.1f}x выше медианного — влияют крупные сделки',
                f'Среднее: {avg_check:,.1f} тыс. | Медиана: {median_check:,.1f} тыс.')

    # ── 7. Нулевые и аномальные суммы ────────────────────────
    zero_rev = (df_full[revenue_col].fillna(0) == 0).sum()
    if zero_rev > 0:
        add('🟡 Средний', 'Качество данных',
            f'{zero_rev} заказов с нулевой выручкой',
            'Проверь: возможно незакрытые сделки или технические строки')

    if not signals:
        add('🟢 Инфо', 'Общее', 'Явных аномалий не обнаружено', 'Данные выглядят стабильно')

    df_signals = pd.DataFrame(signals)
    priority_order = {'🔴 Высокий': 0, '🟡 Средний': 1, '🟢 Инфо': 2}
    df_signals['_sort'] = df_signals['Приоритет'].map(priority_order).fillna(3)
    return df_signals.sort_values('_sort').drop(columns='_sort').reset_index(drop=True)


def reconciliation_warnings(unclassified_projects, external_quality_report,
                            payment_order_reconciliation):
    """
    Замечания к сверке — одни и те же для листа сигналов и для GUI.
    Возвращает список dict: Приоритет/Категория/Сигнал/Детали + 'diag' (для GUI).
    """
    warnings = []
    if unclassified_projects is not None and not unclassified_projects.empty:
        unclass_total = unclassified_projects['Выручка, тыс. руб.'].sum()
        if unclass_total > 0:
            warnings.append({
                'Приоритет': '🔴 Высокий',
                'Категория': 'Сверка',
                'Сигнал': f'Есть неклассифицированные проекты: {unclass_total:,.0f} тыс. руб.',
                'Детали': 'Проверь лист 18_Неклассифицированные: эти проекты не попали в PROJECT_GROUP_MAP.',
                'diag': f"неклассифицированные проекты: {unclass_total:,.0f} тыс.",
            })

    if external_quality_report is not None and not external_quality_report.empty:
        bad_checks = external_quality_report[
            external_quality_report['Статус'].isin(['Критично', 'Проверить'])
        ]
        if not bad_checks.empty:
            warnings.append({
                'Приоритет': '🔴 Высокий' if (bad_checks['Статус'] == 'Критично').any() else '🟡 Средний',
                'Категория': 'Внешние доходы',
                'Сигнал': f'Проверки external_income.json: {len(bad_checks)} замечаний',
                'Детали': '; '.join(bad_checks['Проверка'].head(5).tolist()),
                'diag': f"external_income.json: {len(bad_checks)} замечаний",
            })

    if payment_order_reconciliation is not None and not payment_order_reconciliation.empty:
        total_order_k = payment_order_reconciliation['CRM по дате заказа, тыс. руб.'].sum()
        total_payment_k = payment_order_reconciliation['CRM по дате оплаты, тыс. руб.'].sum()
        payment_delta_pct = _pct_delta(total_payment_k, total_order_k)
        if payment_delta_pct is not None and abs(payment_delta_pct) > RECON_WARN_PCT:
            warnings.append({
                'Приоритет': '🟡 Средний',
                'Категория': 'Дата оплаты',
                'Сигнал': f'По оплате и по заказу разные итоги: {payment_delta_pct:+.2f}%',
                'Детали': 'Проверь лист 19_Заказ_vs_Оплата: бухгалтерия обычно ближе к оплате/закрытию.',
                'diag': f"заказ vs оплата: {payment_delta_pct:+.2f}%",
            })
    return warnings


# ==============================
# ЭТАПЫ РАСЧЁТА
# ==============================

def load_reference_data(input_path, verified_path, external_path, log):
    """verified_figures.json и external_income.json: (verified|None, ext|None, путь к ext)."""
    verified = None
    for vf_path in _json_candidates(verified_path, VERIFIED_FIGURES_JSON, input_path):
        if os.path.exists(vf_path):
            verified = load_verified_figures(vf_path)
            if verified:
                log(f"Верифицированные цифры подгружены из {os.path.abspath(vf_path)}")
                break
    if not verified:
        log("⚠ verified_figures.json не найден — использую константы из analytics.py")

    ext_path = _find_existing(_json_candidates(external_path, EXTERNAL_INCOME_JSON, input_path))
    ext = ext_income.load(ext_path)
    return verified, ext, ext_path


def verified_targets(verified):
    """Цели сверки в руб. из verified_figures.json (или fallback-константы)."""
    v = verified or {}
    return {
        'total_with_prog': v.get('total_with_prog', VERIFIED_TOTAL_WITH_PROG),
        'total_no_prog':   v.get('total_with_barter_no_prog', VERIFIED_TOTAL_WITH_BARTER_NO_PROG),
        'ads_no_events':   v.get('advertising_no_events', VERIFIED_ADVERTISING_NO_EVENTS),
        'prog_external':   v.get('programmatic_external', EXTERNAL_PROGRAMMATIC_TOTAL),
        'other_external':  v.get('other_external_income', EXTERNAL_OTHER_INCOME),
    }


def prepare_crm(df_full):
    """Отрасли, категории выручки и скидок (in place)."""
    df_full['ОТРАСЛЬ_КЛИЕНТА'] = df_full.apply(pick_industry, axis=1)
    df_full['ОТРАСЛЬ_КЛИЕНТА_НОРМ'] = df_full.apply(classify_industry, axis=1)
    df_full['Категория выручки'] = df_full[COL_REVENUE].apply(categorize_revenue_amount)
    if COL_DISCOUNT_PCT in df_full.columns:
        df_full['Скидка_%_число'] = pd.to_numeric(
            df_full[COL_DISCOUNT_PCT].astype(str)
            .str.replace(',', '.', regex=False)
            .str.replace('%', '', regex=False)
            .str.strip(),
            errors='coerce'
        )
        df_full['Категория_скидки'] = df_full['Скидка_%_число'].apply(bucket_discount)


def parse_payments(df_full, revenue_col, log):
    """
    Колонка "ДатаСуммаОплаты_" — одна или НЕСКОЛЬКО строк вида
    "25.02.2025 135000" (по одному платежу на строку, через перенос строки),
    если заказ гасится траншами в разные месяцы, напр.:
      "19.03.2025 23618,75\\n06.10.2025 190618,75"
    Возвращает (Дата_оплаты — дата ПЕРВОГО платежа по строке, payments_df —
    платежи по траншам: Дата_оплаты / Сумма_оплаты).
    """
    import re
    date_pat = re.compile(r'(\d{2}\.\d{2}\.\d{4})')

    def parse_pay_lines(s):
        """
        (pairs, ok). pairs — (дата, сумма) по строкам ячейки, где дата
        распозналась (сумма может быть None). ok=False, если хотя бы одна
        непустая строка НЕ разобралась целиком — сигнал не доверять разбивке
        по траншам и откатиться на резерв (иначе сумма тихо теряется).
        """
        if pd.isna(s):
            return [], True
        pairs = []
        ok = True
        for line in str(s).split('\n'):
            line = line.strip()
            if not line:
                continue
            m = date_pat.match(line)
            if not m:
                ok = False
                continue
            try:
                dt = pd.to_datetime(m.group(1), format='%d.%m.%Y')
            except Exception:
                ok = False
                continue
            rest = line[m.end():].strip()
            amount = parse_money(rest) if rest else None
            if amount is None:
                ok = False
            pairs.append((dt, amount))
        return pairs, ok

    if 'ДатаСуммаОплаты_' in df_full.columns:
        parsed = df_full['ДатаСуммаОплаты_'].apply(parse_pay_lines)
    else:
        parsed = pd.Series([([], True)] * len(df_full), index=df_full.index)

    first_dates = parsed.apply(lambda r: r[0][0][0] if r[0] else pd.NaT)

    # Каждый транш — в СВОЙ месяц со своей суммой, а не вся сумма заказа в месяц
    # первого платежа. Если хоть одна сумма в ячейке не распозналась — относим
    # ВСЮ выручку заказа на дату первого платежа (не рискуем потерять итог).
    pay_rows = []
    fallback_rows = 0
    for idx, (pairs, ok) in parsed.items():
        if not pairs:
            continue
        if ok:
            pay_rows.extend(pairs)
        else:
            pay_rows.append((pairs[0][0], df_full.at[idx, revenue_col]))
            fallback_rows += 1
    if fallback_rows:
        log(f"⚠ {fallback_rows} строк(и) 'ДатаСуммаОплаты_' с нераспознанной строкой/суммой "
            f"платежа — вся выручка заказа отнесена на дату первого распознанного платежа")
    return first_dates, pd.DataFrame(pay_rows, columns=['Дата_оплаты', 'Сумма_оплаты'])


def build_monthly_stats(df_full, revenue_col, payments_df, ext_monthly, ext_year,
                        mask_no_events, log):
    """
    01_Месячная_статистика: CRM по колонке 'Месяц', по дате оплаты, YoY,
    внешние доходы. Добавляет в df_full колонку 'Дата_месяц'.
    """
    df_full['Дата_месяц'] = df_full[COL_MONTH].apply(parse_month)
    dated = df_full.dropna(subset=['Дата_месяц'])

    monthly_stats = dated.groupby('Дата_месяц').agg({revenue_col: ['sum', 'count']}).round(2)
    monthly_stats.columns = ['Сумма выручки, руб.', 'Количество заказов']
    monthly_stats['Средний чек, руб.'] = (
        df_full.loc[mask_no_events].dropna(subset=['Дата_месяц'])
        .groupby('Дата_месяц')[revenue_col].mean().round(2)
    )
    monthly_stats['Сумма выручки, тыс. руб.'] = monthly_stats['Сумма выручки, руб.'] / 1000

    # Сортируем по реальной дате, а не по строке 'MM.YYYY' — иначе на границе года
    # (напр. 11.2025..02.2026) строки идут не по хронологии, что ломает прогноз,
    # тренд-сигналы и порядок точек на графике.
    monthly_stats = monthly_stats.reset_index().sort_values('Дата_месяц').reset_index(drop=True)
    monthly_stats['Период'] = monthly_stats['Дата_месяц'].dt.strftime('%m.%Y')
    monthly_stats = monthly_stats.drop(columns=['Дата_месяц'])
    monthly_stats = monthly_stats[['Период'] + [c for c in monthly_stats.columns if c != 'Период']]

    # Выручка по дате оплаты — сумма фактически оплаченных траншей (payments_df),
    # а не полная сумма заказа в каждом месяце, где встретилась дата платежа.
    if not payments_df.empty:
        paydate_sums = (
            payments_df
            .groupby(payments_df['Дата_оплаты'].dt.strftime('%m.%Y'))['Сумма_оплаты']
            .sum() / 1000
        ).round(2).to_dict()
        monthly_stats['Выручка по дате оплаты, тыс. руб.'] = (
            monthly_stats['Период'].map(paydate_sums).fillna(0).round(2)
        )

    # YoY (год к году), если в данных больше одного года
    if len(df_full['Дата_месяц'].dt.year.dropna().unique()) > 1:
        rev_by_period = monthly_stats.set_index('Период')['Сумма выручки, тыс. руб.'].to_dict()

        def prev_year(p):
            mm, yy = p.split('.')
            return rev_by_period.get(f"{mm}.{int(yy) - 1}")

        monthly_stats['Пред. год, тыс. руб.'] = monthly_stats['Период'].map(prev_year)
        monthly_stats['YoY, %'] = monthly_stats.apply(
            lambda r: round((r['Сумма выручки, тыс. руб.'] / r['Пред. год, тыс. руб.'] - 1) * 100, 1)
            if pd.notna(r.get('Пред. год, тыс. руб.')) and r['Пред. год, тыс. руб.'] != 0
            else None,
            axis=1
        )

    if ext_monthly is not None:
        def get_ext(period):
            # '07.2025' → ключ '07'. Внешние доходы в JSON — за один год ('_год'),
            # поэтому к месяцам других лет их не добавляем.
            m, y = period.split('.')
            if ext_year and int(y) != ext_year:
                return 0
            return ext_monthly.get(m, 0) / 1000  # в тыс. руб.

        other_years = sorted({int(p.split('.')[1]) for p in monthly_stats['Период']} - {ext_year}) \
            if ext_year else []
        if other_years:
            log(f"⚠ external_income.json за {ext_year} год — к месяцам "
                f"{', '.join(map(str, other_years))} г. внешние доходы не добавлены")

        monthly_stats['Внешние доходы, тыс. руб.'] = monthly_stats['Период'].apply(get_ext).round(2)
        monthly_stats['Итого с внешними, тыс. руб.'] = (
            monthly_stats['Сумма выручки, тыс. руб.'] + monthly_stats['Внешние доходы, тыс. руб.']
        ).round(2)
        if 'Выручка по дате оплаты, тыс. руб.' in monthly_stats.columns:
            monthly_stats['Итого по оплате + внешние, тыс. руб.'] = (
                monthly_stats['Выручка по дате оплаты, тыс. руб.'] + monthly_stats['Внешние доходы, тыс. руб.']
            ).round(2)

    return monthly_stats


def build_group_stats(df_full, revenue_col):
    """14_Выручка_по_группам (добавляет в df_full колонку 'БИЗНЕС_ГРУППА')."""
    df_full['БИЗНЕС_ГРУППА'] = (
        df_full[COL_PROJECT].map(PROJECT_GROUP_MAP).fillna('НЕ КЛАССИФИЦИРОВАНО')
        if COL_PROJECT in df_full.columns else 'НЕ КЛАССИФИЦИРОВАНО'
    )
    group_stats = df_full.groupby('БИЗНЕС_ГРУППА').agg({revenue_col: ['sum', 'count']}).round(2)
    group_stats.columns = ['Выручка CRM, руб.', 'Заказов']
    group_stats['Выручка CRM, тыс. руб.'] = (group_stats['Выручка CRM, руб.'] / 1000).round(2)
    return group_stats.reset_index().sort_values('Выручка CRM, тыс. руб.', ascending=False)


def reconciliation_scope(df_full, verified, ext_year, log):
    """
    Верифицированные цифры и external_income.json — годовые. CRM для сверки
    берём только за этот год (по колонке 'Месяц'), иначе выгрузка за 13+
    месяцев или за другой год сравнивается с годовым итогом бухгалтерии.
    Возвращает (recon_year, recon_mask, recon_applicable, recon_label).
    """
    recon_year = None
    if isinstance(verified, dict) and isinstance(verified.get('_год'), (int, float)):
        recon_year = int(verified['_год'])
    elif ext_year:
        recon_year = ext_year
    recon_mask = pd.Series(True, index=df_full.index)
    applicable = True
    if recon_year and 'Дата_месяц' in df_full.columns:
        in_year = df_full['Дата_месяц'].map(lambda d: pd.notna(d) and d.year == recon_year).astype(bool)
        if in_year.any():
            recon_mask = in_year
            skipped = int((~in_year).sum())
            if skipped:
                log(f"Сверка с бухгалтерией за {recon_year} г.: {skipped} строк других "
                    f"периодов в сверку не включены")
        else:
            applicable = False
            log(f"⚠ В выгрузке нет строк за {recon_year} г. (год verified_figures/"
                f"external_income) — сверка с бухгалтерией не применима")
    label = f' ({recon_year})' if recon_year and applicable else ''
    return recon_year, recon_mask, applicable, label


def build_group_summary(df_full, revenue_col, recon_mask, recon_label, rev_bez_prog, rev_reklama,
                        ext_cats, targets):
    """15_Сводка_по_группам — CRM по группам vs бухгалтерия (без попытки сложить)."""
    grp = df_full['БИЗНЕС_ГРУППА']
    rev = df_full[revenue_col].where(recon_mask)
    crm_total = rev.sum()

    gs = {k: ext_cats.get(k, 0) / 1000 for k in
          ('programmatic', 'barter', 'iri_grants', 'recsys', 'ecom', '47_plan', 'deductions', 'total_income')}
    has_ext = bool(ext_cats)
    crm_plus_prog = crm_total / 1000 + gs['programmatic']
    crm_plus_all = crm_total / 1000 + gs['total_income']
    crm_no_prog_bar = rev_bez_prog + gs['barter']

    vf_prog_k = targets['total_with_prog'] / 1000 if targets['total_with_prog'] else None
    vf_no_prog_k = targets['total_no_prog'] / 1000 if targets['total_no_prog'] else None
    vf_ads_k = targets['ads_no_events'] / 1000 if targets['ads_no_events'] else None

    def ext_val(key):
        return round(gs[key], 2) if has_ext else '—'

    rows = [
        (f'═══ CRM (расчёт по выгрузке){recon_label} ═══', ''),
        ('Реклама Фонтанка + Доктор, тыс. руб.', round(rev[grp.isin(['Реклама Фонтанка', 'Реклама Доктор'])].sum() / 1000, 2)),
        ('Программатик (CRM-часть), тыс. руб.',  round(rev[grp == 'Программатик'].sum() / 1000, 2)),
        ('Мероприятия, тыс. руб.',               round(rev[grp == 'Мероприятия'].sum() / 1000, 2)),
        ('47News / Прочее, тыс. руб.',           round(rev[grp == '47News / Прочее'].sum() / 1000, 2)),
        ('НЕ классифицированное, тыс. руб.',     round(rev[grp == 'НЕ КЛАССИФИЦИРОВАНО'].sum() / 1000, 2)),
        ('Итого CRM, тыс. руб.',                 round(crm_total / 1000, 2)),
        ('═══ Внешние доходы (external_income.json) ═══', ''),
        ('Программатик (вне CRM), тыс. руб.',     ext_val('programmatic')),
        ('Медийный бартер, тыс. руб.',            ext_val('barter')),
        ('ИРИ / Гранты, тыс. руб.',               ext_val('iri_grants')),
        ('Рекомендательные системы, тыс. руб.',   ext_val('recsys')),
        ('E-com, тыс. руб.',                      ext_val('ecom')),
        ('47News (в план), тыс. руб.',            ext_val('47_plan')),
        ('Вычеты (взаимозачёты/корректировки), тыс. руб.', ext_val('deductions')),
        ('Итого внешних доходов (без вычетов), тыс. руб.', ext_val('total_income')),
        ('═══ Сверка с верифицированными цифрами ═══', ''),
        ('CRM + программатик (факт), тыс. руб.',
            round(crm_plus_prog, 2) if has_ext else round(crm_total / 1000, 2)),
        ('CRM + все внешние доходы (факт), тыс. руб.',
            round(crm_plus_all, 2) if has_ext else round(crm_total / 1000, 2)),
        ('Верифицировано: всего с прогр. (бух.), тыс. руб.', round(vf_prog_k, 2) if vf_prog_k else '—'),
        ('Расхождение: всего с прогр., тыс. руб.',
            round(crm_plus_all - vf_prog_k, 2) if vf_prog_k and has_ext else '—'),
        ('Расхождение: всего с прогр., %',
            _fmt_pct(_pct_delta(crm_plus_all, vf_prog_k)) if vf_prog_k and has_ext else '—'),
        ('── Без программатика ──', ''),
        ('CRM (без прогр.) + бартер (факт), тыс. руб.', round(crm_no_prog_bar, 2) if has_ext else '—'),
        ('Верифицировано: с бартером без прогр. (бух.), тыс. руб.',
            round(vf_no_prog_k, 2) if vf_no_prog_k else '—'),
        ('Расхождение: без программатика, тыс. руб.',
            round(crm_no_prog_bar - vf_no_prog_k, 2) if vf_no_prog_k and has_ext else '—'),
        ('Расхождение: без программатика, %',
            _fmt_pct(_pct_delta(crm_no_prog_bar, vf_no_prog_k)) if vf_no_prog_k and has_ext else '—'),
        ('── Рекламная без мероприятий ──', ''),
        ('CRM реклама (без мер.), тыс. руб.', round(rev_reklama, 2)),
        ('Верифицировано: реклама без мер. (бух.), тыс. руб.', round(vf_ads_k, 2) if vf_ads_k else '—'),
        ('Расхождение: реклама, тыс. руб.', round(rev_reklama - vf_ads_k, 2) if vf_ads_k else '—'),
        ('Расхождение: реклама, %', _fmt_pct(_pct_delta(rev_reklama, vf_ads_k)) if vf_ads_k else '—'),
        ('═══ Справочно ═══', ''),
        ('Программатик полный (verified_figures), тыс. руб.',
            round(targets['prog_external'] / 1000, 2) if targets['prog_external'] else '—'),
        ('Прочие доходы (verified_figures), тыс. руб.',
            round(targets['other_external'] / 1000, 2) if targets['other_external'] else '—'),
    ]
    return pd.DataFrame(rows, columns=['Показатель', 'Значение']), crm_total


def build_seasonality(df_full, revenue_col, mask_no_events, ext_monthly, ext_year, log):
    """05_Сезонность и 06_Кварталы по дате заказа (+ внешние доходы, если применимо)."""
    df_full['Месяц_число'] = df_full['Дата_заказа'].dt.month
    df_full['Квартал'] = df_full['Дата_заказа'].dt.quarter
    # Без даты — None (выпадает из группировки, как и в кварталах);
    # раньше такие строки попадали в «Осень».
    df_full['Сезон'] = df_full['Месяц_число'].map(
        lambda x: None if pd.isna(x) else MONTH_TO_SEASON[f"{int(x):02d}"])

    seasonal = (
        df_full.groupby('Сезон')
        .agg({revenue_col: ['sum', 'count'], 'КОНЕЧНЫЙ_КЛИЕНТ': 'nunique'}).round(2)
    )
    seasonal.columns = ['Выручка, руб.', 'Количество заказов', 'Уникальных клиентов']
    seasonal['Средний чек, руб.'] = df_full.loc[mask_no_events].groupby('Сезон')[revenue_col].mean().round(2)
    seasonal['Выручка, тыс. руб.'] = seasonal['Выручка, руб.'] / 1000
    seasonal = seasonal.reset_index()

    quarterly = df_full.groupby('Квартал').agg({revenue_col: ['sum', 'count']}).round(2)
    quarterly.columns = ['Выручка, руб.', 'Количество заказов']
    quarterly['Выручка, тыс. руб.'] = quarterly['Выручка, руб.'] / 1000
    quarterly['Коэф. сезонности'] = (quarterly['Выручка, тыс. руб.'] / quarterly['Выручка, тыс. руб.'].mean()).round(2)
    quarterly = quarterly.reset_index()

    # Внешние доходы — только если выгрузка целиком за год внешних доходов:
    # кварталы/сезоны не разделены по годам.
    crm_years = (set(df_full['Дата_месяц'].dropna().map(lambda d: d.year))
                 if 'Дата_месяц' in df_full.columns else set())
    ext_fits = ext_year is None or not crm_years or crm_years == {ext_year}
    if ext_monthly is not None and not ext_fits:
        log(f"⚠ Кварталы/сезоны: выгрузка за {sorted(crm_years)}, внешние доходы за "
            f"{ext_year} — внешние не добавлены")
    if ext_monthly is not None and ext_fits:
        ext_by_quarter = {q: 0 for q in range(1, 5)}
        ext_by_season = {'Зима': 0, 'Весна': 0, 'Лето': 0, 'Осень': 0}
        for m, v in ext_monthly.items():
            ext_by_quarter[MONTH_TO_QUARTER[m]] += v / 1000
            ext_by_season[MONTH_TO_SEASON[m]] += v / 1000
        for table, key, mapping in ((quarterly, 'Квартал', ext_by_quarter),
                                    (seasonal, 'Сезон', ext_by_season)):
            table['Внешние доходы, тыс. руб.'] = table[key].map(mapping).round(2)
            table['Итого с внешними, тыс. руб.'] = (
                table['Выручка, тыс. руб.'] + table['Внешние доходы, тыс. руб.']
            ).round(2)
    return seasonal, quarterly


def build_rfm(rfm_source, revenue_col):
    """07–09: RFM по клиентам без мероприятий. Возвращает (все, сегменты, не-топ)."""
    current_date = rfm_source['Дата_заказа'].max()
    rfm = (
        rfm_source.groupby('КОНЕЧНЫЙ_КЛИЕНТ')
        # nunique, а не count: у одного заказа много строк-позиций
        .agg({'Дата_заказа': 'max', COL_ORDER: 'nunique', revenue_col: 'sum'})
        .reset_index()
        .rename(columns={'Дата_заказа': 'Последняя_покупка', COL_ORDER: 'Частота',
                         revenue_col: 'Денежная_ценность_руб'})
    )
    rfm['Recency_days']       = (current_date - rfm['Последняя_покупка']).dt.days
    rfm['Frequency_orders']   = rfm['Частота']
    rfm['Monetary_thousands'] = rfm['Денежная_ценность_руб'] / 1000
    rfm = rfm.dropna(subset=['Recency_days', 'Frequency_orders', 'Monetary_thousands'])

    q = rfm[['Recency_days', 'Frequency_orders', 'Monetary_thousands']].quantile([0.25, 0.5, 0.75])

    def score(x, col, ascending):
        # Квартиль 1..4; для давности меньше = лучше
        s = 1 if x <= q[col][0.25] else 2 if x <= q[col][0.5] else 3 if x <= q[col][0.75] else 4
        return s if ascending else 5 - s

    rfm['R'] = rfm['Recency_days'].apply(score, args=('Recency_days', False))
    rfm['F'] = rfm['Frequency_orders'].apply(score, args=('Frequency_orders', True))
    rfm['M'] = rfm['Monetary_thousands'].apply(score, args=('Monetary_thousands', True))
    rfm['RFM_Score'] = rfm[['R', 'F', 'M']].sum(axis=1)

    # Сегмент — по сочетанию давности (R) и ценности (F/M), а не по сумме
    # баллов: при сумме «At Risk» получали и недавние мелкие клиенты, хотя
    # смысл сегмента — ценные клиенты, которые давно не покупали.
    def segment(r):
        recent = r['R'] >= 3
        valuable = r['F'] >= 3 or r['M'] >= 3
        if recent and r['F'] >= 3 and r['M'] >= 3: return 'Champions'
        if recent and valuable: return 'Loyal Customers'
        if recent: return 'Potential Loyalists'
        if valuable: return 'At Risk'
        return 'Hibernating'

    rfm['Segment'] = rfm.apply(segment, axis=1)
    rfm_all = rfm.sort_values(['RFM_Score', 'Monetary_thousands'], ascending=[False, False]).reset_index(drop=True)

    segments = (
        rfm_all.groupby('Segment')
        .agg(Клиентов=('КОНЕЧНЫЙ_КЛИЕНТ', 'count'),
             Выручка_тыс=('Monetary_thousands', 'sum'),
             Средний_RFM=('RFM_Score', 'mean'))
        .reset_index()
    )
    segments['Доля клиентов, %'] = (segments['Клиентов'] / rfm_all['КОНЕЧНЫЙ_КЛИЕНТ'].nunique() * 100).round(1)
    segments['Доля выручки, %'] = (segments['Выручка_тыс'] / rfm_all['Monetary_thousands'].sum() * 100).round(1)
    non_top = rfm_all[~rfm_all['Segment'].isin(['Champions', 'Loyal Customers'])].copy()
    return rfm_all, segments, non_top


def build_order_distribution(df_full, revenue_col):
    """10_Распределение_заказов: сколько клиентов сделали N уникальных заказов."""
    per_client = df_full.groupby('КОНЕЧНЫЙ_КЛИЕНТ').agg({COL_ORDER: 'nunique', revenue_col: 'sum'})
    freq = per_client[COL_ORDER].value_counts().sort_index()
    return pd.DataFrame({
        'Количество заказов': freq.index,
        'Количество клиентов': freq.values,
        'Доля клиентов, %': (freq.values / freq.sum() * 100).round(1)
    })


def build_summary(df_full, revenue_col, mask_no_events, recon_label, rev_all, rev_bez_prog,
                  rev_reklama, targets):
    """11_Расширенный_сводный: CRM-расчёт рядом с верифицированными цифрами."""
    rows = [
        (f'— CRM (расчёт по выгрузке){recon_label} —', ''),
        ('Выручка: все заказы в CRM, тыс. руб.', round(rev_all, 2)),
        ('Выручка: CRM без программатика, тыс. руб.', round(rev_bez_prog, 2)),
        ('Выручка: CRM рекламная без мероприятий, тыс. руб.', round(rev_reklama, 2)),
        ('— Верифицированные (из бухгалтерии) —', ''),
    ]
    for key, label in (('total_with_prog', 'всего с программатиком'),
                       ('total_no_prog', 'с бартером без программатика'),
                       ('ads_no_events', 'рекламная без мероприятий')):
        if targets[key]:
            rows.append((f'Верифицированная: {label}, тыс. руб.', round(targets[key] / 1000, 2)))
    no_events_rev = df_full.loc[mask_no_events, revenue_col]
    rows.extend([
        ('— Прочее —', ''),
        ('Средний чек БЕЗ мероприятий, тыс. руб.', round(no_events_rev.mean() / 1000, 2)),
        ('Медианный чек БЕЗ мероприятий, тыс. руб.', round(no_events_rev.median() / 1000, 2)),
        ('Количество заказов', len(df_full)),
        ('Уникальных клиентов', df_full['КОНЕЧНЫЙ_КЛИЕНТ'].nunique()),
        ('Уникальных менеджеров', df_full[COL_MANAGER].nunique() if COL_MANAGER in df_full.columns else None),
    ])
    return pd.DataFrame(rows, columns=['Метрика', 'Значение'])


def build_comparison_metrics(df_full, recon_label, rev_all, rev_bez_prog, rev_reklama,
                             ext_cats, targets, monthly_stats):
    """12_Сравнение: детальная сверка CRM vs бухгалтерия + помесячно."""
    k = {key: ext_cats.get(key, 0) / 1000 for key in
         ('programmatic', 'barter', 'iri_grants', 'recsys', 'ecom', '47_plan', 'other',
          'deductions', 'total_income', 'total_net')}

    # «Строк в исходнике» — без строк с пустым/служебным 'Месяц'
    if COL_MONTH in df_full.columns:
        rows_src = int((~df_full[COL_MONTH].astype(str).str.strip().str.lower()
                        .isin(['итого', 'nan', 'none'])).sum())
    else:
        rows_src = len(df_full)

    rows = [
        ('═══ CRM (выгрузка) ═══', '', '', ''),
        ('Строк в исходнике',           rows_src,                      '',    ''),
        ('Строк после очистки',         len(df_full),                  '',    ''),
        (f'CRM итого{recon_label}, тыс. руб.', round(rev_all, 2),      '',    ''),
        ('CRM без программатика, тыс.', round(rev_bez_prog, 2),        '',    ''),
        ('CRM рекламная (без мер.), тыс.', round(rev_reklama, 2),      '',    ''),
    ]
    if ext_cats:
        rows += [
            ('═══ Внешние доходы (external_income.json) ═══', '', '', ''),
            ('Программатик (все), тыс.',     round(k['programmatic'], 2), '', ''),
            ('Бартер медийный, тыс.',        round(k['barter'], 2), '', ''),
            ('ИРИ / Гранты, тыс.',           round(k['iri_grants'], 2), '', ''),
            ('Рекомендательные системы, тыс.', round(k['recsys'], 2), '', ''),
            ('E-com, тыс.',                  round(k['ecom'], 2), '', ''),
            ('47News (в план), тыс.',        round(k['47_plan'], 2), '', ''),
            ('Прочие доходные, тыс.',        round(k['other'], 2), '', ''),
            ('── Вычеты (закупка/взаимозачёт), тыс.', round(k['deductions'], 2), '', ''),
            ('Итого внешние доходы (без вычетов), тыс.', round(k['total_income'], 2), '', ''),
            ('Итого внешние нетто (с вычетами), тыс.',   round(k['total_net'], 2), '', ''),
        ]

    rows.append(('═══ Сверка с бухгалтерией ═══', '', '', ''))
    slices = [
        ('── Срез: всего с программатиком ──', 'CRM + внешние доходы, тыс.',
         rev_all + k['total_income'], targets['total_with_prog']),
        ('── Срез: с бартером, без программатика ──', 'CRM (без прогр.) + бартер, тыс.',
         rev_bez_prog + k['barter'], targets['total_no_prog']),
        ('── Срез: рекламная без мероприятий ──', 'CRM рекламная (без мер.), тыс.',
         rev_reklama, targets['ads_no_events']),
    ]
    for title, fact_label, fact, target in slices:
        if not target:
            continue
        target_k = target / 1000
        rows += [
            (title, '', '', ''),
            (fact_label,                    round(fact, 2), '', ''),
            ('Верифицировано (бух.), тыс.', round(target_k, 2), '', ''),
            ('Расхождение абс., тыс.',      round(fact - target_k, 2), '', ''),
            ('Расхождение, %',              _fmt_pct(_pct_delta(fact, target_k)), '', ''),
        ]

    if monthly_stats is not None and ext_cats:
        rows.append(('═══ Помесячно: CRM + программатик ═══', '', '', ''))
        for _, mrow in monthly_stats.iterrows():
            crm_m = mrow['Сумма выручки, тыс. руб.']
            ext_m = mrow.get('Внешние доходы, тыс. руб.', 0) or 0
            total_m = mrow.get('Итого с внешними, тыс. руб.', crm_m + ext_m)
            rows.append((mrow['Период'], round(crm_m, 2), round(ext_m, 2), round(total_m, 2)))

    return pd.DataFrame(
        rows, columns=['Показатель / Период', 'CRM, тыс. руб.', 'Внешние, тыс. руб.', 'Итого, тыс. руб.'])


def build_forecast(monthly_stats):
    """
    13_Прогноз: линейный тренд по закрытым месяцам. Ось X — номер месяца
    (год*12+мес), а не индекс строки, чтобы пропуски не сжимали шкалу.
    """
    closed = _closed_months(monthly_stats)
    if closed is None or len(closed) < 3:
        return None
    x = closed['_ord'].to_numpy(dtype=float)
    y = closed['Сумма выручки, тыс. руб.'].to_numpy(dtype=float)
    coeff = np.polyfit(x, y, 1)
    next_ord = int(x[-1]) + 1
    return pd.DataFrame({
        'Метрика': [
            'Прогноз выручки на след. месяц, тыс. руб.',
            'Прогнозируемый период',
            'Направление тренда',
            'Сила тренда, % от среднего',
            'Месяцев в расчёте (закрытых)',
        ],
        'Значение': [
            round(coeff[0] * next_ord + coeff[1], 2),
            f"{(next_ord - 1) % 12 + 1:02d}.{(next_ord - 1) // 12}",
            "↑ растёт" if coeff[0] > 0 else "↓ падает",
            round(abs(coeff[0]) / y.mean() * 100 if y.mean() else 0, 2),
            len(closed),
        ]
    })


def build_manager_plan(manager_stats, manager_plan):
    """03а_План_менеджеров: факт / план / % / статус."""
    plan_df = manager_stats.copy()
    # Менеджеры с планом, но без продаж в выгрузке, тоже должны быть в таблице
    missing = [m for m, p in manager_plan.items() if p and m not in set(plan_df[COL_MANAGER])]
    if missing:
        plan_df = pd.concat([plan_df, pd.DataFrame({
            COL_MANAGER: missing,
            'Сумма выручки, руб.': 0.0,
            'Сумма выручки, тыс. руб.': 0.0,
        })], ignore_index=True)
    plan_df['План, руб.'] = plan_df[COL_MANAGER].map(manager_plan).fillna(0)
    plan_df['План, тыс. руб.'] = (plan_df['План, руб.'] / 1000).round(2)
    has_plan = plan_df['План, руб.'] > 0
    plan_df['Выполнение, %'] = (plan_df['Сумма выручки, руб.'] / plan_df['План, руб.'] * 100).round(1).where(has_plan)
    plan_df['Остаток до плана, тыс. руб.'] = (
        (plan_df['План, руб.'] - plan_df['Сумма выручки, руб.']) / 1000).round(2).where(has_plan)
    plan_df['Статус'] = plan_df['Выполнение, %'].apply(
        lambda x: '— нет плана' if pd.isna(x) else
                  '✅ Выполнен' if x >= 100 else '⚠ В работе' if x >= 70 else '❌ Отстаёт'
    )
    return plan_df[[
        COL_MANAGER, 'Сумма выручки, тыс. руб.', 'План, тыс. руб.',
        'Выполнение, %', 'Остаток до плана, тыс. руб.', 'Статус'
    ]].sort_values('Выполнение, %', ascending=False, na_position='last')


def build_nomenclature(df_full, revenue_col, mask_no_events):
    """04а_Номенклатура: топ-50 без мероприятий."""
    stats = (
        df_full.loc[mask_no_events].groupby(COL_NOMEN)
        .agg(Выручка_руб=(revenue_col, 'sum'), Заказов=(COL_ORDER, 'count'),
             Клиентов=('КОНЕЧНЫЙ_КЛИЕНТ', 'nunique'))
        .reset_index()
        .sort_values('Выручка_руб', ascending=False)
    )
    stats['Выручка, тыс. руб.'] = (stats['Выручка_руб'] / 1000).round(2)
    stats['Средний чек, тыс. руб.'] = (stats['Выручка_руб'] / stats['Заказов'] / 1000).round(2)
    stats['Доля выручки, %'] = (stats['Выручка_руб'] / stats['Выручка_руб'].sum() * 100).round(1)
    return stats[[COL_NOMEN, 'Выручка, тыс. руб.', 'Доля выручки, %', 'Заказов', 'Клиентов',
                  'Средний чек, тыс. руб.']].head(50)


def build_discounts(df_full, revenue_col, mask_no_events):
    """17*: скидки — сводка, по клиентам, менеджерам, бакетам. None, если скидок нет."""
    disc_col = 'Скидка_%_число'
    df_disc = df_full.loc[mask_no_events & df_full[disc_col].notna() & (df_full[disc_col] > 0)]
    if df_disc.empty:
        return None

    def by(key, sort_col, head=None):
        t = (df_disc.groupby(key)
             .agg(Заказов=(COL_ORDER, 'count'), Выручка_руб=(revenue_col, 'sum'),
                  Ср_скидка=(disc_col, 'mean'), Макс_скидка=(disc_col, 'max'))
             .reset_index().sort_values(sort_col, ascending=False))
        if head:
            t = t.head(head)
        t['Выручка, тыс. руб.'] = (t['Выручка_руб'] / 1000).round(2)
        t['Ср. скидка, %'] = t['Ср_скидка'].round(1)
        t['Макс. скидка, %'] = t['Макс_скидка'].round(1)
        return t[[key, 'Выручка, тыс. руб.', 'Заказов', 'Ср. скидка, %', 'Макс. скидка, %']]

    buckets = (
        df_disc.groupby('Категория_скидки')
        .agg(Заказов=(COL_ORDER, 'count'), Выручка_руб=(revenue_col, 'sum'))
        .reset_index().sort_values('Выручка_руб', ascending=False)
    )
    buckets['Выручка, тыс. руб.'] = (buckets['Выручка_руб'] / 1000).round(2)
    total_disc_rev = buckets['Выручка_руб'].sum()
    buckets['Доля, %'] = (buckets['Выручка_руб'] / total_disc_rev * 100).round(1) if total_disc_rev > 0 else 0

    total_orders = int(mask_no_events.sum())
    no_events_rev = df_full.loc[mask_no_events, revenue_col].sum()
    summary = pd.DataFrame([
        ('Заказов без мероприятий (всего)', total_orders),
        ('Заказов со скидкой', len(df_disc)),
        ('Доля заказов со скидкой, %', round(len(df_disc) / total_orders * 100, 1) if total_orders else 0),
        ('Средняя скидка по скидочным сделкам, %', round(df_disc[disc_col].mean(), 1)),
        ('Медианная скидка, %', round(df_disc[disc_col].median(), 1)),
        ('Максимальная скидка, %', round(df_disc[disc_col].max(), 1)),
        ('Выручка со скидкой, тыс. руб.', round(df_disc[revenue_col].sum() / 1000, 2)),
        ('Доля выручки со скидкой, %',
            round(df_disc[revenue_col].sum() / no_events_rev * 100, 1) if no_events_rev > 0 else 0),
    ], columns=['Метрика', 'Значение'])

    return {
        'summary': summary,
        'by_client': by('КОНЕЧНЫЙ_КЛИЕНТ', 'Выручка_руб', head=30),
        'by_manager': by(COL_MANAGER, 'Ср_скидка') if COL_MANAGER in df_disc.columns else None,
        'by_bucket': buckets[['Категория_скидки', 'Заказов', 'Выручка, тыс. руб.', 'Доля, %']],
    }


def export_report(export_data, output_path, log):
    """Пишет листы в xlsx, оформляет и добавляет лист с графиками."""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name, df_sheet in export_data.items():
            safe_name = sheet_name[:31]
            try:
                df_sheet.to_excel(writer, sheet_name=safe_name, index=False)
                autosize_columns(writer.sheets[safe_name], df_sheet)
            except Exception as e:
                log(f"Ошибка при сохранении листа {safe_name}: {e}")

    log("Строю графики...")
    try:
        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        style_workbook(wb, log=log)
        add_report_charts(wb)
        wb.save(output_path)
        log("Графики добавлены ✅")
    except Exception as e:
        log(f"⚠ Графики не добавлены: {e}")


# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ — запускается из app.py
# ============================================================

def run_analytics(input_path: str, output_path: str, log=print,
                  manager_plan: dict = None, date_by: str = 'order',
                  verified_path: str = None, external_path: str = None):
    """
    Запускает полный цикл аналитики.
    input_path   — путь к входному .xlsx
    output_path  — путь к выходному .xlsx
    log          — функция для вывода сообщений (print или GUI-лог)
    manager_plan — dict {имя_менеджера: план_руб} или None
    date_by      — 'order' (дата заказа, по умолчанию) или 'payment' (дата оплаты)
    verified_path / external_path — явные пути к JSON-справочникам (из GUI)

    Возвращает: dict с ключами:
        'output_path'       — путь к созданному файлу
        'crm_total'         — итог CRM за год сверки, тыс. руб.
        'crm_paydate_total' — итог CRM по дате оплаты, тыс. руб.
        'external_total'    — внешние доходы для месячной аналитики, тыс. руб.
        'full_external_total' — все внешние доходы для сверки, тыс. руб.
        'grand_total'       — CRM + внешние, тыс. руб. (по выбранной дате)
        'verified_total'    — верифицированная цель, тыс. руб.
        'deviation_pct'     — процент отклонения от верифицированной цели
        'date_by', 'diagnostics'
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    # ── 0. Справочники ───────────────────────────────────────
    verified, ext, ext_path = load_reference_data(input_path, verified_path, external_path, log)
    targets = verified_targets(verified)
    ext_year = ext_income.year(ext)
    ext_monthly = ext_income.monthly_totals(ext)
    if ext_monthly:
        log(f"Внешние доходы подгружены: {sum(ext_monthly.values())/1000:,.0f} тыс. руб. "
            f"({os.path.abspath(ext_path)})")
    vf_year = verified.get('_год') if isinstance(verified, dict) else None
    if isinstance(vf_year, (int, float)) and ext_year and int(vf_year) != ext_year:
        log(f"⚠ Несоответствие годов: verified_figures.json={int(vf_year)}, "
            f"external_income.json={ext_year}. "
            "Проверьте, что оба файла относятся к одному отчетному году.")
    ext_cats = ext_income.totals_by_category(ext)
    external_quality_report = ext_income.quality_report(ext, verified, _pct_delta, _recon_status)

    # ── 1. Загрузка и контроль качества ──────────────────────
    log("Загрузка файла...")
    df_full = load_crm(input_path, with_clients=False)
    log(f"Загружено строк: {len(df_full)}")

    log("Проверка качества данных...")
    quality_report, critical_errors = run_data_quality_checks(df_full, log=log)
    if critical_errors:
        error_text = "\n".join(f" - {e}" for e in critical_errors)
        raise ValueError(f"Критические ошибки в данных:\n{error_text}")
    log("Качество данных: ОК")

    # ── 2. Подготовка ────────────────────────────────────────
    revenue_col = COL_REVENUE
    df_full['КОНЕЧНЫЙ_КЛИЕНТ'] = df_full.apply(pick_client, axis=1).apply(normalize_client)
    prepare_crm(df_full)
    df_full['Дата_оплаты'], payments_df = parse_payments(df_full, revenue_col, log)
    mask_no_events = _mask_no_events(df_full)
    mask_no_prog = _mask_no_prog(df_full)

    # ── 3. Месячная статистика ───────────────────────────────
    monthly_stats = payment_order_reconciliation = None
    if COL_MONTH in df_full.columns:
        log("Считаю месячную статистику...")
        monthly_stats = build_monthly_stats(df_full, revenue_col, payments_df, ext_monthly,
                                            ext_year, mask_no_events, log)
        payment_order_reconciliation = build_payment_order_reconciliation(monthly_stats)

    # ── 4. Клиенты, менеджеры, отрасли, группы ───────────────
    log("Считаю топ клиентов...")
    client_df_for_rank = df_full.loc[client_rank_mask(df_full)]
    client_stats = _rev_stats(client_df_for_rank, 'КОНЕЧНЫЙ_КЛИЕНТ', revenue_col)

    manager_stats = None
    if COL_MANAGER in df_full.columns:
        log("Считаю статистику по менеджерам...")
        manager_df = df_full.loc[~df_full[COL_MANAGER].isin(EXCLUDE_MANAGERS)]
        manager_stats = _rev_stats(manager_df, COL_MANAGER, revenue_col,
                                   avg_mask=manager_df.index.isin(df_full.index[mask_no_events]))

    log("Считаю по отраслям...")
    industry_stats = _rev_stats(df_full, 'ОТРАСЛЬ_КЛИЕНТА_НОРМ', revenue_col, avg_mask=mask_no_events)

    log("Считаю выручку по группам...")
    group_stats = build_group_stats(df_full, revenue_col)
    unclassified_projects = build_unclassified_projects_report(df_full, revenue_col)

    # ── 5. Сверка с бухгалтерией (за год сверки) ─────────────
    recon_year, recon_mask, recon_applicable, recon_label = \
        reconciliation_scope(df_full, verified, ext_year, log)
    # rev_bez_prog — через mask_no_prog (regex по названию проекта), а не через
    # БИЗНЕС_ГРУППА=='Программатик': иначе «CRM минус программатик» в мосте
    # не равнялся бы «CRM без программатика».
    rev_reklama  = df_full.loc[mask_no_prog & mask_no_events & recon_mask, revenue_col].sum() / 1000
    rev_bez_prog = df_full.loc[mask_no_prog & recon_mask, revenue_col].sum() / 1000
    group_summary_df, crm_total = build_group_summary(
        df_full, revenue_col, recon_mask, recon_label, rev_bez_prog, rev_reklama, ext_cats, targets)
    rev_all = crm_total / 1000

    # Дата заказа — один раз, для сезонности, RFM и проблемных строк
    if COL_DATE in df_full.columns:
        df_full['Дата_заказа'] = parse_dates(df_full[COL_DATE])
    crm_issue_rows = build_crm_issue_rows(df_full, revenue_col)

    # ── 6. Сезонность, RFM, лояльность ───────────────────────
    seasonal_stats = quarterly_stats = None
    if 'Дата_заказа' in df_full.columns:
        log("Считаю сезонность...")
        seasonal_stats, quarterly_stats = build_seasonality(
            df_full, revenue_col, mask_no_events, ext_monthly, ext_year, log)

    rfm_all = rfm_segment_extended = rfm_non_top = None
    if 'Дата_заказа' in df_full.columns:
        try:
            log("Провожу RFM-анализ...")
            rfm_all, rfm_segment_extended, rfm_non_top = build_rfm(client_df_for_rank, revenue_col)
        except Exception as e:
            log(f"RFM-анализ не выполнен: {e}")

    log("Считаю лояльность клиентов...")
    order_freq_df = build_order_distribution(df_full, revenue_col)

    # ── 7. Сводки и сверка ───────────────────────────────────
    summary_df = build_summary(df_full, revenue_col, mask_no_events, recon_label,
                               rev_all, rev_bez_prog, rev_reklama, targets)
    comparison_metrics = build_comparison_metrics(df_full, recon_label, rev_all, rev_bez_prog,
                                                  rev_reklama, ext_cats, targets, monthly_stats)

    def k(v):
        return v / 1000 if v else None
    reconciliation_bridge = build_reconciliation_bridge(
        rev_all_k=rev_all,
        rev_bez_prog_k=rev_bez_prog,
        rev_reklama_k=rev_reklama,
        # rev_all - rev_bez_prog (не crm_prog!) — согласовано по построению
        # с «CRM без программатика» (оба через mask_no_prog).
        crm_prog_k=rev_all - rev_bez_prog,
        ext_cats=ext_cats,
        vf_total_with_prog_k=k(targets['total_with_prog']),
        vf_total_no_prog_k=k(targets['total_no_prog']),
        vf_ads_no_events_k=k(targets['ads_no_events']),
    )

    # ── 8. Прогноз, план, номенклатура, скидки ───────────────
    forecast_summary = build_forecast(monthly_stats)

    plan_stats = None
    if manager_plan and manager_stats is not None and len(manager_stats) > 0:
        log("Считаю выполнение плана...")
        plan_stats = build_manager_plan(manager_stats, manager_plan)

    nomen_stats = None
    if COL_NOMEN in df_full.columns:
        log("Считаю топ номенклатур...")
        nomen_stats = build_nomenclature(df_full, revenue_col, mask_no_events)

    discount_stats = None
    if 'Скидка_%_число' in df_full.columns:
        log("Считаю анализ скидок...")
        discount_stats = build_discounts(df_full, revenue_col, mask_no_events)

    # ── 9. Сигналы ───────────────────────────────────────────
    log("Ищу сигналы и аномалии...")
    signals_df = build_signals(
        df_full, revenue_col, monthly_stats, manager_stats, client_stats,
        rfm_all, mask_no_events, log=log, client_mask=client_df_for_rank.index
    )
    warnings = reconciliation_warnings(unclassified_projects, external_quality_report,
                                       payment_order_reconciliation)
    if warnings:
        signals_df = pd.concat([pd.DataFrame([{k_: v for k_, v in w.items() if k_ != 'diag'}
                                              for w in warnings]), signals_df], ignore_index=True)

    log("Строю бухгалтерскую таблицу...")
    accounting_table = None
    if ext is not None:
        accounting_table = build_accounting_table(df_full, revenue_col, ext, log=log)
    else:
        log(f"⚠ Файл {EXTERNAL_INCOME_JSON} не найден — бухгалтерская таблица не построена")

    # ── 10. Экспорт ──────────────────────────────────────────
    log("Сохраняю отчёт...")
    sheets = [
        ('00_Сигналы', signals_df),
        ('00_Data_Quality', quality_report),
        ('01_Месячная_статистика', monthly_stats),
        ('02_Топ_клиентов', client_stats.head(50)),
        ('03_Топ_менеджеров', manager_stats.head(30) if manager_stats is not None else None),
        ('03а_План_менеджеров', plan_stats),
        ('04_Отрасли', industry_stats),
        ('04а_Номенклатура', nomen_stats),
        ('05_Сезонность', seasonal_stats),
        ('06_Кварталы', quarterly_stats),
        ('07_RFM_все', rfm_all),
        ('08_RFM_сегменты', rfm_segment_extended),
        ('09_RFM_не_топ', rfm_non_top),
        ('10_Распределение_заказов', order_freq_df),
        ('11_Расширенный_сводный', summary_df),
        ('12_Сравнение', comparison_metrics),
        ('13_Прогноз', forecast_summary),
        ('14_Выручка_по_группам', group_stats),
        ('15_Сводка_по_группам', group_summary_df),
        ('16_Бухгалтерская_таблица', accounting_table),
    ]
    if discount_stats is not None:
        sheets += [
            ('17_Скидки_сводка', discount_stats['summary']),
            ('17а_Скидки_клиенты', discount_stats['by_client']),
            ('17б_Скидки_менеджеры', discount_stats['by_manager']),
            ('17в_Скидки_по_бакетам', discount_stats['by_bucket']),
        ]
    sheets += [
        ('18_Мост_расхождений', reconciliation_bridge),
        ('18а_Неклассифицированные', unclassified_projects),
        ('18б_CRM_проблемные_строки', crm_issue_rows),
        ('18в_External_QA', external_quality_report),
        ('19_Заказ_vs_Оплата', payment_order_reconciliation),
    ]
    export_report({name: df for name, df in sheets if df is not None}, output_path, log)
    log(f"✅ Готово! Отчёт сохранён: {output_path}")

    # ── 11. Сводка расхождения для GUI ───────────────────────
    # Для сверки — только год сверки (recon_mask); оплаты — по году даты оплаты.
    crm_total_k = df_full.loc[recon_mask, revenue_col].sum() / 1000
    pay_recon = payments_df
    if recon_year and recon_applicable and not payments_df.empty:
        pay_recon = payments_df.loc[payments_df['Дата_оплаты'].dt.year == recon_year]
    crm_paydate_total_k = pay_recon['Сумма_оплаты'].sum() / 1000 if not pay_recon.empty else 0

    # external_total_k — то, что добавляется в месячную аналитику (программатик);
    # full_external_k — все внешние доходы без закупочных/технических статей.
    external_total_k = sum(ext_monthly.values()) / 1000 if ext_monthly else 0
    full_external_k = ext_income.full_total(ext) / 1000 if ext is not None else 0.0

    by_payment = date_by == 'payment' and crm_paydate_total_k > 0
    crm_base_k = crm_paydate_total_k if by_payment else crm_total_k
    grand_total_k = crm_base_k + full_external_k
    verified_total_k = targets['total_with_prog'] / 1000 if targets['total_with_prog'] else 0
    deviation_pct = (
        (grand_total_k / verified_total_k - 1) * 100
        if verified_total_k > 0 and recon_applicable else None
    )
    diagnostics = [w['diag'] for w in warnings]

    date_label = "по дате оплаты" if by_payment else "по дате заказа"
    log(f"📊 CRM ({date_label}): {crm_base_k:,.0f} тыс. | "
        f"Прогр. (ежемес.): {external_total_k:,.0f} тыс. | "
        f"Все внешние доходы: {full_external_k:,.0f} тыс.")
    if deviation_pct is not None:
        log(f"📊 Итого CRM + внешние: {grand_total_k:,.0f} тыс. | "
            f"Верифицировано: {verified_total_k:,.0f} тыс. | "
            f"Отклонение: {deviation_pct:+.2f}% — {recon_verdict(deviation_pct)}")
    else:
        log(f"📊 Итого: {grand_total_k:,.0f} тыс.")
    if diagnostics:
        log("🔎 Диагностика сверки: " + " | ".join(diagnostics))

    return {
        'output_path':           output_path,
        'crm_total':             round(crm_total_k, 2),
        'crm_paydate_total':     round(crm_paydate_total_k, 2),
        'external_total':        round(external_total_k, 2),
        'full_external_total':   round(full_external_k, 2),
        'grand_total':           round(grand_total_k, 2),
        'verified_total':        round(verified_total_k, 2),
        'deviation_pct':         round(deviation_pct, 2) if deviation_pct is not None else None,
        'date_by':               date_by,
        'diagnostics':           diagnostics,
    }
