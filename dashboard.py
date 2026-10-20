# ============================================================
# dashboard.py — генерация HTML-дашборда
# ============================================================

import os
import json
import webbrowser
from datetime import datetime

import pandas as pd

from crm_data import (
    COL_MANAGER, COL_MONTH, COL_REVENUE, EXCLUDE_MANAGERS,
    classify_industry, client_rank_mask, load_crm, mask_no_events, parse_month, pick_industry,
)

# Шаблон рядом с модулем (в собранном exe — в папке распаковки, см. build.py)
DASHBOARD_TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                  'templates', 'dashboard.html')


def _load_and_prepare(path: str) -> pd.DataFrame:
    """
    Та же загрузка и очистка, что у отчёта «Анализ» (crm_data.load_crm) —
    иначе KPI дашборда (Выручка, Заказов, Клиентов) с ним не совпадают.
    """
    df = load_crm(path)
    # Колонка должна называться именно 'ОТРАСЛЬ_КЛИЕНТА' — её читает classify_industry()
    df['ОТРАСЛЬ_КЛИЕНТА'] = df.apply(pick_industry, axis=1)
    df['ОТРАСЛЬ_НОРМ'] = df.apply(classify_industry, axis=1)
    if COL_MONTH in df.columns:
        df['_month_dt'] = df[COL_MONTH].apply(parse_month)
    return df


def _collect_data(df: pd.DataFrame) -> dict:
    rev_total = round(df[COL_REVENUE].sum() / 1000, 1)
    orders    = len(df)
    clients   = df['КОНЕЧНЫЙ_КЛИЕНТ'].nunique()

    # Средний чек — без мероприятий (см. README), как в run_analytics.
    avg_check = round(df.loc[mask_no_events(df), COL_REVENUE].mean() / 1000, 1)

    monthly = []
    if '_month_dt' in df.columns:
        # Группируем и сортируем по реальной дате, а не по строке 'MM.YYYY' —
        # иначе на границе года (напр. 11.2025..02.2026) месяцы идут не по
        # порядку, что ломает график и KPI "изменение к пред. месяцу".
        grp = (df.dropna(subset=['_month_dt'])
                 .groupby('_month_dt')[COL_REVENUE].sum()
                 .reset_index()
                 .sort_values('_month_dt'))
        grp['month'] = grp['_month_dt'].dt.strftime('%m.%Y')
        grp['revenue'] = (grp[COL_REVENUE] / 1000).round(1)
        monthly = grp[['month', 'revenue']].to_dict('records')

    # Без мероприятий — как лист 02_Топ_клиентов в отчёте
    top_cli = (df.loc[client_rank_mask(df)].groupby('КОНЕЧНЫЙ_КЛИЕНТ')[COL_REVENUE]
                 .sum().nlargest(10).reset_index())
    top_cli.columns = ['name', 'revenue']
    top_cli['revenue'] = (top_cli['revenue'] / 1000).round(1)
    top_cli['name'] = top_cli['name'].astype(str).str[:35]

    top_mgr = pd.DataFrame()
    if COL_MANAGER in df.columns:
        # Менеджеры вне отдела исключаются только здесь (см. README), не из общих итогов.
        mgr_df = df.loc[~df[COL_MANAGER].isin(EXCLUDE_MANAGERS)]
        top_mgr = (mgr_df.groupby(COL_MANAGER)[COL_REVENUE]
                     .sum().nlargest(10).reset_index())
        top_mgr.columns = ['name', 'revenue']
        top_mgr['revenue'] = (top_mgr['revenue'] / 1000).round(1)

    industries_full = df.groupby('ОТРАСЛЬ_НОРМ')[COL_REVENUE].sum()
    industries_full = industries_full[industries_full.index.astype(str).str.strip() != '']
    # Общая сумма по ВСЕМ отраслям (не только показанным топ-8) — используется как
    # знаменатель для % на графике, иначе при >8 отраслях проценты завышаются.
    industries_total_k = round(industries_full.sum() / 1000, 1)
    industries = industries_full.nlargest(8).reset_index()
    industries.columns = ['name', 'revenue']
    industries['revenue'] = (industries['revenue'] / 1000).round(1)

    # Тренд: изменение к предыдущему месяцу
    trend_pct = None
    if len(monthly) >= 2:
        prev = monthly[-2]['revenue']
        curr = monthly[-1]['revenue']
        if prev > 0:
            trend_pct = round((curr - prev) / prev * 100, 1)

    return {
        'meta': {
            'filename': '',
            'generated': datetime.now().strftime('%d.%m.%Y %H:%M'),
        },
        'kpi': {
            'revenue':   rev_total,
            'orders':    orders,
            'clients':   clients,
            'avg_check': avg_check,
            'trend_pct': trend_pct,
        },
        'monthly':      monthly,
        'top_clients':  top_cli.to_dict('records'),
        'top_managers': top_mgr.to_dict('records') if not top_mgr.empty else [],
        'industries':   industries.to_dict('records'),
        'industries_total': industries_total_k,
    }


def _generate_html(data: dict, filename: str) -> str:
    data['meta']['filename'] = os.path.basename(filename)
    # '</' экранируем: клиент с '</script>' в названии иначе ломал страницу
    d = json.dumps(data, ensure_ascii=False).replace('</', '<\\/')

    with open(DASHBOARD_TEMPLATE, encoding='utf-8') as f:
        html = f.read()

    return html.replace('__DATA_PLACEHOLDER__', d)


def generate_dashboard(input_path: str, log=print) -> str:
    log(f"Загружаю данные: {os.path.basename(input_path)}")
    df = _load_and_prepare(input_path)
    log(f"Строк после фильтров: {len(df)}")

    log("Формирую дашборд...")
    data = _collect_data(df)

    out_dir = os.path.dirname(input_path)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(out_dir, f"dashboard_{ts}.html")

    html = _generate_html(data, input_path)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    log(f"✅ Дашборд сохранён: {out_path}")
    webbrowser.open(f"file:///{out_path.replace(os.sep, '/')}")
    return out_path
