# ============================================================
# external_income.py — внешние доходы (external_income.json)
# То, чего нет в CRM: программатик, бартер, ИРИ/гранты, 47News и т.д.
# Файл читается один раз (load), остальные функции получают dict.
# ============================================================

import json
import os

import pandas as pd

EXTERNAL_INCOME_JSON = 'external_income.json'

MONTHS = [f"{m:02d}" for m in range(1, 13)]

# Статьи, которые не являются реальным доходом (закупка, взаимозачёты,
# затраты, корректировки) — если в JSON не задан '_не_включать_в_grand_total'
DEFAULT_DEDUCTIONS = [
    'Выручка 47 (закупка)',
    'Взаимозачет/Затраты',
    'ФФ/АМ взаимозачет',
    'Корректировка скидки (комиссия ХШМ)',
]

PROG_KEYS   = {'Программатик ФОНТАНКА', 'Программатик ФОНТАНКА ТГ',
               'Программатик ДОКТОР', 'Программатик ДОКТОР ТГ'}
BARTER_KEYS = {'Медийный бартер ФОНТАНКА', 'Медийный бартер ЭВЕНТЫ'}
IRI_KEYS    = {'ИРИ/АНО/Петроцентр', 'Гранты Фонтанка', 'Гранты Доктор'}
RECSYS_KEYS = {'Рекомендательные системы ФОНТАНКА', 'Рекомендательные системы ДОКТОР'}
ECOM_KEYS   = {'E-com ФОНТАНКА', 'E-com ДОКТОР'}
PLAN47_KEYS = {'Выручка 47 (в план)'}
DEDUCT_KEYS = set(DEFAULT_DEDUCTIONS)

# Статьи, которые нельзя прибавлять к CRM: дублируют строки выгрузки (47News есть
# в CRM проектами '47News*', бартер — строками с 'Бартер' = 'Да') либо находятся
# вне периметра (ИРИ/АНО). В сверке бартер и ИРИ/АНО показываются отдельно.
EXTERNAL_DUPLICATES_OF_CRM = {'Выручка 47 (в план)', 'Медийный бартер ФОНТАНКА', 'Медийный бартер ЭВЕНТЫ'}
EXTERNAL_OUT_OF_PERIMETER = {'ИРИ/АНО/Петроцентр', 'Гранты Фонтанка', 'Гранты Доктор'}


def load(path, errors=None):
    """
    Содержимое external_income.json или None (нет файла / не читается).
    errors — список, куда добавляется текст ошибки чтения (для отчёта качества).
    """
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        if errors is not None:
            errors.append(str(e))
        return None


def data_rows(ext):
    """{статья: {'01': руб, ...}} — без служебных ключей '_...'."""
    return {k: v for k, v in (ext or {}).items()
            if not str(k).startswith('_') and isinstance(v, dict)}


def row_sum(values):
    """Сумма числовых месяцев статьи."""
    return sum(v for v in (values.get(m) for m in MONTHS) if isinstance(v, (int, float)))


def year(ext):
    """Отчётный год ('_год') или None."""
    y = (ext or {}).get('_год')
    return int(y) if isinstance(y, (int, float)) else None


def monthly_totals(ext):
    """
    {месяц: сумма_руб} тех статей, что перечислены в '_включать_в_аналитику'
    (реальная выручка, а не вычеты). Если ключ задан — берём только его
    (пустой список = ничего не добавлять); если ключа нет — все статьи.
    Раньше пустой список трактовался как «ключа нет», и в аналитику
    попадали все статьи, включая вычеты и закупку.
    """
    if ext is None:
        return None
    rows = data_rows(ext)
    if '_включать_в_аналитику' in ext:
        rows = {r: rows[r] for r in (ext.get('_включать_в_аналитику') or []) if r in rows}
    totals = {m: 0.0 for m in MONTHS}
    for values in rows.values():
        for m in MONTHS:
            v = values.get(m, 0) or 0
            if isinstance(v, (int, float)):
                totals[m] += v
    return totals


def full_total(ext):
    """
    Внешние доходы (руб.), прибавляемые к периметру CRM — без статей из
    '_не_включать_в_grand_total' (закупка, взаимозачёты, затраты — это не
    реальный доход, а технические/расходные проводки).
    """
    exclude = set((ext or {}).get('_не_включать_в_grand_total', DEFAULT_DEDUCTIONS))
    # Дубли CRM и статьи вне периметра тоже не прибавляются к CRM
    exclude |= EXTERNAL_DUPLICATES_OF_CRM | EXTERNAL_OUT_OF_PERIMETER
    return sum(row_sum(v) for k, v in data_rows(ext).items() if k not in exclude)


def totals_by_category(ext):
    """
    Разбивка внешних доходов по категориям для сверки (руб.):
    programmatic, barter, iri_grants, recsys, ecom, 47_plan, other,
    deductions — сумма модулей вычетов (в JSON закупка/взаимозачёты
    положительные, корректировка скидки отрицательная, но все они
    уменьшают доход), total_income — всё кроме вычетов,
    total_net — total_income - deductions.
    """
    if ext is None:
        return {}
    cats = {k: 0.0 for k in ('programmatic', 'barter', 'iri_grants', 'recsys',
                             'ecom', '47_plan', 'other', 'deductions')}
    for key, values in data_rows(ext).items():
        s = row_sum(values)
        if key in PROG_KEYS:
            cats['programmatic'] += s
        elif key in BARTER_KEYS:
            cats['barter'] += s
        elif key in IRI_KEYS:
            cats['iri_grants'] += s
        elif key in RECSYS_KEYS:
            cats['recsys'] += s
        elif key in ECOM_KEYS:
            cats['ecom'] += s
        elif key in PLAN47_KEYS:
            cats['47_plan'] += s
        elif key in DEDUCT_KEYS:
            cats['deductions'] += abs(s)
        else:
            cats['other'] += s
    cats['total_income'] = sum(cats[k] for k in ('programmatic', 'barter', 'iri_grants',
                                                 'recsys', 'ecom', '47_plan', 'other'))
    cats['total_net'] = cats['total_income'] - cats['deductions']
    return cats


def perimeter_total(cats):
    """
    Внешние доходы (руб.), которые прибавляются к периметру CRM: программатик,
    рекомендательные системы, e-com и прочие доходные статьи. Без бартера и
    'Выручка 47 (в план)' (дубли CRM), без ИРИ/АНО/грантов (вне периметра)
    и без вычетов.
    """
    return sum(cats.get(k, 0) for k in ('programmatic', 'recsys', 'ecom', 'other'))


def quality_report(ext, verified_data=None, pct_delta=None, recon_status=None, other_keys=(),
                   load_error=None):
    """
    Проверяет external_income.json на ошибки, которые чаще всего ломают сверку.
    Возвращает DataFrame: Проверка / Статус / Детали.
    pct_delta, recon_status — функции из analytics (чтобы не дублировать);
    other_keys — статьи блока 'Total Прочие доходы' бухгалтерии.
    """
    rows = []

    def add(check, status, details):
        rows.append({'Проверка': check, 'Статус': status, 'Детали': details})

    if load_error:
        add('Чтение external_income.json', 'Критично', load_error)
        return pd.DataFrame(rows)
    if ext is None:
        add('Файл external_income.json', 'Критично', 'Файл не найден')
        return pd.DataFrame(rows)

    rows_data = data_rows(ext)
    add('Строки внешних доходов', 'OK' if rows_data else 'Критично',
        f'Найдено строк: {len(rows_data)}')

    verified_year = verified_data.get('_год') if isinstance(verified_data, dict) else None
    external_year = ext.get('_год')
    if verified_year and external_year and int(verified_year) != int(external_year):
        add('Год сверки', 'Критично',
            f'verified_figures.json={verified_year}, external_income.json={external_year}')
    else:
        add('Год сверки', 'OK', f'Год: {external_year or verified_year or "не задан"}')

    missing_include = [r for r in ext.get('_включать_в_аналитику', []) if r not in rows_data]
    add('Список _включать_в_аналитику',
        'OK' if not missing_include else 'Критично',
        'Все строки найдены' if not missing_include else ', '.join(missing_include))

    missing_months, non_numeric = [], []
    for row_name, values in rows_data.items():
        for month in MONTHS:
            if month not in values:
                missing_months.append(f'{row_name}: {month}')
                continue
            value = values.get(month)
            if value not in ('', None) and not isinstance(value, (int, float)):
                non_numeric.append(f'{row_name}: {month}={value}')

    add('Наличие 12 месяцев',
        'OK' if not missing_months else 'Проверить',
        'Все строки имеют 12 месяцев' if not missing_months else '; '.join(missing_months[:20]))
    add('Числовой формат сумм',
        'OK' if not non_numeric else 'Критично',
        'Все суммы числовые или пустые' if not non_numeric else '; '.join(non_numeric[:20]))

    exclude = set(ext.get('_не_включать_в_grand_total', []))
    negative = [f'{name}: {row_sum(values):,.0f} руб.'
                for name, values in rows_data.items()
                if name not in exclude and row_sum(values) < 0]
    add('Отрицательные доходные статьи',
        'OK' if not negative else 'Проверить',
        'Не обнаружены' if not negative else '; '.join(negative))

    cats = totals_by_category(ext)
    if isinstance(verified_data, dict):
        verified_prog = verified_data.get('programmatic_external')
        if verified_prog:
            prog = cats.get('programmatic', 0)
            add('Программатик vs verified_figures',
                recon_status(pct_delta(prog, verified_prog), warn=1, bad=3),
                f'JSON={prog:,.0f} руб.; verified={verified_prog:,.0f} руб.; '
                f'delta={prog - verified_prog:,.0f} руб.')

        verified_other = verified_data.get('other_external_income')
        if verified_other:
            # other_external_income — это блок 'Total Прочие доходы' бухгалтерии
            # (включая закупку 47 и взаимозачёт ФФ/АМ): сравниваем с теми же статьями
            other = sum(row_sum(rows_data.get(k, {})) for k in other_keys)
            add('Прочие доходы (блок) vs verified_figures',
                recon_status(pct_delta(other, verified_other), warn=1, bad=3),
                f'JSON={other:,.0f} руб.; verified={verified_other:,.0f} руб.; '
                f'delta={other - verified_other:,.0f} руб.')

    return pd.DataFrame(rows)
