# ============================================================
# comparison.py — сравнительный анализ двух периодов
# Вызывается из app.py. Загрузка и фильтры — общие с отчётом
# «Анализ» (crm_data), оформление — excel_style.
# ============================================================

import os

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import Reference
from openpyxl.styles import Font

from crm_data import (
    COL_MANAGER, COL_ORDER, COL_REVENUE, EXCLUDE_MANAGERS,
    client_rank_mask, load_crm, mask_no_events,
)
from excel_style import (
    ACCENT, ACCENT_2, HEADER_FILL, HEADER_FONT, LOST_FILL, NEG_FONT, NEW_FILL, POS_FONT, POSITIVE,
    autosize_sheet, color_series, make_chart, style_body_cell, style_header_row,
)

DELTA_COL = 'Δ выручка, тыс.'


def _load_period(path: str, log) -> pd.DataFrame:
    """
    Загружает файл с той же очисткой, что и run_analytics. EXCLUDE_MANAGERS
    и мероприятия не вырезаются из общего датасета — как и в отчёте «Анализ»,
    они влияют только на конкретные срезы (менеджеры / клиенты / средний чек).
    """
    df = load_crm(path)
    log(f"  Загружено строк после фильтров: {len(df)}")
    return df


# ============================================================
# РАСЧЁТЫ ПО ПЕРИОДУ
# ============================================================

def _stats(df: pd.DataFrame, key: str) -> pd.DataFrame:
    """Выручка (тыс.) и число строк-заказов по ключу."""
    stats = (
        df.groupby(key)
        .agg(Выручка_руб=(COL_REVENUE, 'sum'), Заказов=(COL_ORDER, 'count'))
        .reset_index()
        .sort_values('Выручка_руб', ascending=False)
    )
    stats['Выручка_тыс'] = (stats['Выручка_руб'] / 1000).round(2)
    return stats


def _manager_stats(df: pd.DataFrame) -> pd.DataFrame:
    if COL_MANAGER not in df.columns:
        return pd.DataFrame(columns=[COL_MANAGER, 'Выручка_руб', 'Заказов', 'Выручка_тыс'])
    # Менеджеры вне отдела исключаются только из этого среза (см. README)
    return _stats(df.loc[~df[COL_MANAGER].isin(EXCLUDE_MANAGERS)], COL_MANAGER)


def _client_stats(df: pd.DataFrame) -> pd.DataFrame:
    # Без мероприятий — как лист 02_Топ_клиентов в отчёте «Анализ»
    return _stats(df.loc[client_rank_mask(df)], 'КОНЕЧНЫЙ_КЛИЕНТ')


# ============================================================
# ЛИСТЫ
# ============================================================

def _append_row(ws, values, fill=None):
    ws.append(list(values))
    for cell in ws[ws.max_row]:
        style_body_cell(cell, fill)


def _color_delta(ws, delta_idx, pct_idx=None, first_row=2):
    for r in range(first_row, ws.max_row + 1):
        val = ws.cell(r, delta_idx).value
        if val is not None:
            font = POS_FONT if float(val) >= 0 else NEG_FONT
            ws.cell(r, delta_idx).font = font
            if pct_idx:
                ws.cell(r, pct_idx).font = font


def _sheet_summary(wb, df_a, df_b, label_a, label_b):
    ws = wb.create_sheet("01_Сводка")
    ws.freeze_panes = "A2"
    ws.append(["Метрика", label_a, label_b, "Δ абс.", "Δ %"])
    style_header_row(ws, 1)

    # Средний/медианный чек — без мероприятий (см. README), как в run_analytics
    mask_a, mask_b = mask_no_events(df_a), mask_no_events(df_b)
    metrics = [
        ("Выручка, тыс. руб.", round(df_a[COL_REVENUE].sum() / 1000, 2), round(df_b[COL_REVENUE].sum() / 1000, 2)),
        ("Количество заказов", len(df_a), len(df_b)),
        ("Средний чек, тыс. руб.", round(df_a.loc[mask_a, COL_REVENUE].mean() / 1000, 2),
                                   round(df_b.loc[mask_b, COL_REVENUE].mean() / 1000, 2)),
        ("Медианный чек, тыс. руб.", round(df_a.loc[mask_a, COL_REVENUE].median() / 1000, 2),
                                     round(df_b.loc[mask_b, COL_REVENUE].median() / 1000, 2)),
        ("Уникальных клиентов", df_a['КОНЕЧНЫЙ_КЛИЕНТ'].nunique(), df_b['КОНЕЧНЫЙ_КЛИЕНТ'].nunique()),
    ]
    if COL_MANAGER in df_a.columns and COL_MANAGER in df_b.columns:
        metrics.append(("Активных менеджеров", df_a[COL_MANAGER].nunique(), df_b[COL_MANAGER].nunique()))

    for name, va, vb in metrics:
        delta_abs = round(vb - va, 2)
        delta_pct = round((vb - va) / abs(va) * 100, 1) if va else None
        _append_row(ws, [name, va, vb, delta_abs, delta_pct])
    _color_delta(ws, 4, 5)

    ws.row_dimensions[1].height = 22
    autosize_sheet(ws)


def _merge_periods(stats_a, stats_b, key, label_a, label_b, sort_by_delta):
    """Слияние статистики двух периодов по ключу + Δ выручки."""
    col_a, col_b = f'Выручка_{label_a}_тыс', f'Выручка_{label_b}_тыс'
    merged = pd.merge(
        stats_a.rename(columns={'Выручка_тыс': col_a, 'Заказов': f'Заказов_{label_a}'})
               [[key, col_a, f'Заказов_{label_a}']],
        stats_b.rename(columns={'Выручка_тыс': col_b, 'Заказов': f'Заказов_{label_b}'})
               [[key, col_b, f'Заказов_{label_b}']],
        on=key, how='outer'
    ).fillna(0)
    merged[DELTA_COL] = (merged[col_b] - merged[col_a]).round(2)
    merged['Δ %'] = merged.apply(
        lambda r: round((r[col_b] - r[col_a]) / abs(r[col_a]) * 100, 1) if r[col_a] != 0 else None, axis=1
    )
    return merged.sort_values(DELTA_COL if sort_by_delta else col_b, ascending=False).reset_index(drop=True)


def _sheet_table(wb, title, merged):
    ws = wb.create_sheet(title)
    ws.freeze_panes = "A2"
    headers = list(merged.columns)
    ws.append(headers)
    style_header_row(ws, 1)
    for row in merged.itertuples(index=False):
        _append_row(ws, row)
    _color_delta(ws, headers.index(DELTA_COL) + 1, headers.index('Δ %') + 1)
    autosize_sheet(ws)


def _client_block(ws, title, df, clients, label, fill):
    """Блок «новые/ушедшие клиенты»: заголовок + список с выручкой."""
    ws.append([title])
    ws.cell(ws.max_row, 1).font = Font(bold=True, name="Segoe UI", size=10)
    ws.append(["Клиент", f"Выручка {label}, тыс.", "Заказов"])
    style_header_row(ws, ws.max_row)
    detail = (
        df[df['КОНЕЧНЫЙ_КЛИЕНТ'].isin(clients)]
        .groupby('КОНЕЧНЫЙ_КЛИЕНТ')
        .agg(Выручка=(COL_REVENUE, 'sum'), Заказов=(COL_ORDER, 'count'))
        .reset_index()
        .sort_values('Выручка', ascending=False)
    )
    detail['Выручка'] = (detail['Выручка'] / 1000).round(2)
    for _, row in detail.iterrows():
        _append_row(ws, [row['КОНЕЧНЫЙ_КЛИЕНТ'], row['Выручка'], row['Заказов']], fill)


def _sheet_movement(wb, df_a, df_b, label_a, label_b):
    ws = wb.create_sheet("04_Движение_клиентов")

    clients_a = set(df_a['КОНЕЧНЫЙ_КЛИЕНТ'].dropna().unique())
    clients_b = set(df_b['КОНЕЧНЫЙ_КЛИЕНТ'].dropna().unique())
    new_clients, lost_clients, kept_clients = clients_b - clients_a, clients_a - clients_b, clients_a & clients_b

    rev_a = df_a.groupby('КОНЕЧНЫЙ_КЛИЕНТ')[COL_REVENUE].sum()
    rev_b = df_b.groupby('КОНЕЧНЫЙ_КЛИЕНТ')[COL_REVENUE].sum()

    def rev_k(rev, clients):
        return round(rev[rev.index.isin(clients)].sum() / 1000, 2)

    ws.append(["Категория", "Клиентов", f"Выручка {label_a}, тыс.", f"Выручка {label_b}, тыс."])
    style_header_row(ws, 1)
    _append_row(ws, ["Постоянные клиенты (были в обоих периодах)", len(kept_clients),
                     rev_k(rev_a, kept_clients), rev_k(rev_b, kept_clients)])
    _append_row(ws, [f"Новые клиенты (только в {label_b})", len(new_clients),
                     0, rev_k(rev_b, new_clients)], NEW_FILL)
    _append_row(ws, [f"Ушедшие клиенты (только в {label_a})", len(lost_clients),
                     rev_k(rev_a, lost_clients), 0], LOST_FILL)
    ws.append([])

    _client_block(ws, f"НОВЫЕ КЛИЕНТЫ — появились в {label_b}", df_b, new_clients, label_b, NEW_FILL)
    ws.append([])
    _client_block(ws, f"УШЕДШИЕ КЛИЕНТЫ — были в {label_a}, нет в {label_b}",
                  df_a, lost_clients, label_a, LOST_FILL)
    autosize_sheet(ws)


def _chart_table(ws, start_row, headers, rows):
    """Вспомогательная таблица под график: шапка + строки, возвращает число строк."""
    for c, h in enumerate(headers, 1):
        cell = ws.cell(start_row, c, h)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
    for i, values in enumerate(rows):
        for c, v in enumerate(values, 1):
            ws.cell(start_row + 1 + i, c, v)
    return len(rows)


def _sheet_charts(wb, mgr_merged, cli_merged, label_a, label_b):
    ws = wb.create_sheet("05_Графики")
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"СРАВНЕНИЕ: {label_a}  vs  {label_b}"
    ws["A1"].font = Font(bold=True, size=14, name="Segoe UI", color=ACCENT)
    ws.row_dimensions[1].height = 28

    col_a, col_b = f"Выручка_{label_a}_тыс", f"Выручка_{label_b}_тыс"

    def add_grouped(start_row, key_label, rows, kind, title, anchor, axis):
        n = _chart_table(ws, start_row, [key_label, label_a, label_b], rows)
        chart = make_chart(kind, title, width=24, height=16, **axis)
        chart.grouping = "clustered"
        chart.add_data(Reference(ws, min_col=2, max_col=3, min_row=start_row, max_row=start_row + n),
                       titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=start_row + 1, max_row=start_row + n))
        color_series(chart, ACCENT, ACCENT_2)
        ws.add_chart(chart, anchor)
        return n

    # 1. Менеджеры: сгруппированный горизонтальный
    mgr_top = mgr_merged.head(10).reset_index(drop=True)
    n_mgr = add_grouped(3, "Менеджер",
                        [(r[COL_MANAGER], float(r[col_a]), float(r[col_b])) for _, r in mgr_top.iterrows()],
                        "bar", f"Менеджеры: {label_a} vs {label_b}", "E3", {'x_title': "тыс. руб."})

    # 2. Клиенты: топ-10 по сумме двух периодов, вертикальный
    start_row = 4 + n_mgr + 3
    cli_top = cli_merged.assign(_total=cli_merged[col_a] + cli_merged[col_b]) \
                        .nlargest(10, "_total").reset_index(drop=True)
    n_cli = add_grouped(start_row, "Клиент",
                        [(str(r["КОНЕЧНЫЙ_КЛИЕНТ"])[:30], float(r[col_a]), float(r[col_b]))
                         for _, r in cli_top.iterrows()],
                        "col", f"Топ-10 клиентов: {label_a} vs {label_b}", f"A{start_row}",
                        {'y_title': "тыс. руб."})

    # 3. Дельта выручки менеджеров
    delta_row = start_row + n_cli + 5
    n = _chart_table(ws, delta_row, ["Менеджер", DELTA_COL],
                     [(r[COL_MANAGER], float(r[DELTA_COL])) for _, r in mgr_top.iterrows()])
    chart = make_chart("col", "Изменение выручки по менеджерам (Δ тыс. руб.)", width=24, height=14,
                       y_title="тыс. руб.")
    chart.add_data(Reference(ws, min_col=2, min_row=delta_row, max_row=delta_row + n), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=delta_row + 1, max_row=delta_row + n))
    color_series(chart, POSITIVE)
    ws.add_chart(chart, f"E{delta_row}")

    autosize_sheet(ws)


# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================

def run_comparison(
    path_a: str,
    path_b: str,
    output_path: str,
    label_a: str = "Период А",
    label_b: str = "Период Б",
    log=print
):
    """
    Сравнивает два периода и сохраняет отчёт в output_path.
    path_a / path_b — пути к Excel-выгрузкам из CRM
    label_a / label_b — произвольные метки (например, "Янв 2024")
    """
    # Метки становятся частью имён колонок — одинаковые метки ломают слияние
    label_a = (label_a or '').strip() or "Период А"
    label_b = (label_b or '').strip() or "Период Б"
    if label_a == label_b:
        label_a, label_b = f"{label_a} (А)", f"{label_b} (Б)"

    log(f"Загружаю период А: {os.path.basename(path_a)}")
    df_a = _load_period(path_a, log)
    log(f"Загружаю период Б: {os.path.basename(path_b)}")
    df_b = _load_period(path_b, log)

    log("Формирую отчёт...")
    wb = Workbook()
    wb.remove(wb.active)  # удаляем дефолтный лист

    _sheet_summary(wb, df_a, df_b, label_a, label_b)
    mgr_merged = _merge_periods(_manager_stats(df_a), _manager_stats(df_b), COL_MANAGER,
                                label_a, label_b, sort_by_delta=False)
    _sheet_table(wb, "02_Менеджеры", mgr_merged)
    cli_merged = _merge_periods(_client_stats(df_a), _client_stats(df_b), 'КОНЕЧНЫЙ_КЛИЕНТ',
                                label_a, label_b, sort_by_delta=True)
    _sheet_table(wb, "03_Клиенты", cli_merged)
    _sheet_movement(wb, df_a, df_b, label_a, label_b)

    if not mgr_merged.empty and not cli_merged.empty:
        _sheet_charts(wb, mgr_merged, cli_merged, label_a, label_b)

    wb.save(output_path)
    log(f"✅ Готово! Отчёт сравнения сохранён: {output_path}")
    return output_path
