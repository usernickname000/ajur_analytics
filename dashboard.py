# ============================================================
# dashboard.py — генерация HTML-дашборда
# Вызывается из app.py, открывает результат в браузере
# ============================================================

import os
import json
import webbrowser
import tempfile
from datetime import datetime

import pandas as pd

from analytics import (
    parse_money, pick_client, normalize_client, pick_industry, classify_industry,
    EXCLUDE_MANAGERS, EXCLUDE_PROJECTS,
    COL_REVENUE, COL_MONTH, COL_MANAGER, COL_PROJECT,
    COL_ORDER, COL_DATE, COL_CLIENT, COL_CLIENT_RA, COL_REKLAMD,
    parse_month
)


def _load_and_prepare(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=object)
    df.columns = df.columns.map(lambda x: x.strip() if isinstance(x, str) else x)
    if COL_MONTH in df.columns:
        df = df[~df[COL_MONTH].astype(str).str.strip().str.lower().eq('итого')]
    df = df.dropna(how='all').reset_index(drop=True)
    df[COL_REVENUE] = df[COL_REVENUE].apply(parse_money)
    df = df.dropna(subset=[COL_REVENUE])
    df = df[df[COL_REVENUE] > 0].reset_index(drop=True)
    df['КОНЕЧНЫЙ_КЛИЕНТ'] = df.apply(pick_client, axis=1).apply(normalize_client)
    df['ОТРАСЛЬ'] = df.apply(pick_industry, axis=1)
    df['ОТРАСЛЬ_НОРМ'] = df.apply(classify_industry, axis=1)
    if COL_MANAGER in df.columns:
        df = df[~df[COL_MANAGER].isin(EXCLUDE_MANAGERS)].reset_index(drop=True)
    if COL_PROJECT in df.columns:
        df = df[~df[COL_PROJECT].fillna('').isin(EXCLUDE_PROJECTS)].reset_index(drop=True)
    if COL_MONTH in df.columns:
        df['_month_dt'] = df[COL_MONTH].apply(parse_month)
    return df


def _collect_data(df: pd.DataFrame) -> dict:
    rev_total = round(df[COL_REVENUE].sum() / 1000, 1)
    orders    = len(df)
    clients   = df['КОНЕЧНЫЙ_КЛИЕНТ'].nunique()
    avg_check = round(df[COL_REVENUE].mean() / 1000, 1)

    # Месячная выручка
    monthly = []
    if '_month_dt' in df.columns:
        grp = (df.dropna(subset=['_month_dt'])
                 .groupby(df['_month_dt'].dt.strftime('%m.%Y'))
                 [COL_REVENUE].sum()
                 .reset_index())
        grp.columns = ['month', 'revenue']
        grp['revenue'] = (grp['revenue'] / 1000).round(1)
        monthly = grp.to_dict('records')

    # Топ-10 клиентов
    top_cli = (df.groupby('КОНЕЧНЫЙ_КЛИЕНТ')[COL_REVENUE]
                 .sum().nlargest(10).reset_index())
    top_cli.columns = ['name', 'revenue']
    top_cli['revenue'] = (top_cli['revenue'] / 1000).round(1)
    top_cli['name'] = top_cli['name'].astype(str).str[:35]

    # Топ-10 менеджеров
    top_mgr = pd.DataFrame()
    if COL_MANAGER in df.columns:
        top_mgr = (df.groupby(COL_MANAGER)[COL_REVENUE]
                     .sum().nlargest(10).reset_index())
        top_mgr.columns = ['name', 'revenue']
        top_mgr['revenue'] = (top_mgr['revenue'] / 1000).round(1)

    # Отрасли
    industries = (df.groupby('ОТРАСЛЬ_НОРМ')[COL_REVENUE]
                    .sum().nlargest(8).reset_index())
    industries.columns = ['name', 'revenue']
    industries['revenue'] = (industries['revenue'] / 1000).round(1)
    industries = industries[industries['name'].astype(str).str.strip() != '']

    return {
        'meta': {
            'filename': '',
            'generated': datetime.now().strftime('%d.%m.%Y %H:%M'),
        },
        'kpi': {
            'revenue': rev_total,
            'orders':  orders,
            'clients': clients,
            'avg_check': avg_check,
        },
        'monthly':    monthly,
        'top_clients': top_cli.to_dict('records'),
        'top_managers': top_mgr.to_dict('records') if not top_mgr.empty else [],
        'industries':  industries.to_dict('records'),
    }


def _generate_html(data: dict, filename: str) -> str:
    data['meta']['filename'] = os.path.basename(filename)

    d = json.dumps(data, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Аналитика — Фонтанка.ру</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>
  :root {{
    --orange: #F38120;
    --orange-dark: #D06A10;
    --orange-light: rgba(243,129,32,0.12);
    --bg: #0C0C0C;
    --surface: #161616;
    --surface2: #1E1E1E;
    --border: #2A2A2A;
    --text: #F0F0F0;
    --text2: #AAAAAA;
    --green: #22C55E;
    --blue: #60A5FA;
    --purple: #A78BFA;
    --pink: #F472B6;
  }}

  * {{ margin: 0; padding: 0; box-sizing: border-box; }}

  body {{
    font-family: 'Manrope', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
  }}

  /* ── Header ── */
  header {{
    background: linear-gradient(135deg, #1A0A00 0%, #0C0C0C 60%);
    border-bottom: 1px solid var(--border);
    padding: 0 40px;
    height: 72px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 100;
    backdrop-filter: blur(20px);
  }}

  .logo {{
    display: flex;
    align-items: center;
    gap: 16px;
  }}

  .logo-text {{
    font-family: Georgia, serif;
    font-size: 22px;
    font-style: italic;
    font-weight: bold;
    color: var(--orange);
    letter-spacing: -0.5px;
  }}

  .logo-sep {{
    width: 1px;
    height: 28px;
    background: var(--border);
  }}

  .logo-sub {{
    font-size: 11px;
    font-weight: 600;
    color: var(--text2);
    letter-spacing: 2px;
    text-transform: uppercase;
  }}

  .header-meta {{
    font-size: 12px;
    color: var(--text2);
    text-align: right;
    line-height: 1.6;
  }}

  .header-meta span {{
    color: var(--orange);
    font-weight: 600;
  }}

  /* ── Layout ── */
  main {{
    max-width: 1400px;
    margin: 0 auto;
    padding: 40px;
  }}

  /* ── KPI Cards ── */
  .kpi-grid {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 16px;
    margin-bottom: 40px;
  }}

  .kpi-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 28px 24px;
    position: relative;
    overflow: hidden;
    transition: transform 0.2s, border-color 0.2s;
  }}

  .kpi-card:hover {{
    transform: translateY(-2px);
    border-color: var(--orange);
  }}

  .kpi-card::before {{
    content: '';
    position: absolute;
    top: 0; left: 0;
    width: 100%; height: 3px;
    background: var(--accent, var(--orange));
  }}

  .kpi-card:nth-child(2) {{ --accent: var(--blue); }}
  .kpi-card:nth-child(3) {{ --accent: var(--green); }}
  .kpi-card:nth-child(4) {{ --accent: var(--purple); }}

  .kpi-label {{
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: var(--text2);
    margin-bottom: 12px;
  }}

  .kpi-value {{
    font-size: 36px;
    font-weight: 800;
    color: var(--text);
    line-height: 1;
    margin-bottom: 6px;
  }}

  .kpi-unit {{
    font-size: 13px;
    color: var(--text2);
    font-weight: 400;
  }}

  .kpi-icon {{
    position: absolute;
    top: 24px; right: 24px;
    font-size: 28px;
    opacity: 0.15;
  }}

  /* ── Charts Grid ── */
  .charts-grid {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 20px;
    margin-bottom: 20px;
  }}

  .chart-full {{ grid-column: 1 / -1; }}

  .chart-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 28px;
    transition: border-color 0.2s;
  }}

  .chart-card:hover {{ border-color: #3A3A3A; }}

  .chart-title {{
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: var(--text2);
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    gap: 10px;
  }}

  .chart-title::before {{
    content: '';
    width: 3px;
    height: 16px;
    background: var(--orange);
    border-radius: 2px;
  }}

  .chart-wrap {{
    position: relative;
    height: 280px;
  }}

  .chart-wrap-tall {{
    position: relative;
    height: 340px;
  }}

  /* ── Footer ── */
  footer {{
    text-align: center;
    padding: 40px;
    color: var(--text2);
    font-size: 12px;
    border-top: 1px solid var(--border);
    margin-top: 20px;
  }}

  /* ── Animations ── */
  @keyframes fadeUp {{
    from {{ opacity: 0; transform: translateY(20px); }}
    to   {{ opacity: 1; transform: translateY(0); }}
  }}

  .kpi-card {{ animation: fadeUp 0.4s ease both; }}
  .kpi-card:nth-child(1) {{ animation-delay: 0.05s; }}
  .kpi-card:nth-child(2) {{ animation-delay: 0.10s; }}
  .kpi-card:nth-child(3) {{ animation-delay: 0.15s; }}
  .kpi-card:nth-child(4) {{ animation-delay: 0.20s; }}
  .chart-card {{ animation: fadeUp 0.5s ease both; animation-delay: 0.25s; }}

  /* ── Scrollbar ── */
  ::-webkit-scrollbar {{ width: 6px; }}
  ::-webkit-scrollbar-track {{ background: var(--bg); }}
  ::-webkit-scrollbar-thumb {{ background: var(--border); border-radius: 3px; }}
</style>
</head>
<body>

<header>
  <div class="logo">
    <span class="logo-text">фонтанка.ру</span>
    <div class="logo-sep"></div>
    <span class="logo-sub">Аналитика коммерческого отдела</span>
  </div>
  <div class="header-meta">
    <div>Файл: <span id="fname"></span></div>
    <div>Сформировано: <span id="fdate"></span></div>
  </div>
</header>

<main>
  <!-- KPI -->
  <div class="kpi-grid">
    <div class="kpi-card">
      <div class="kpi-icon">💰</div>
      <div class="kpi-label">Выручка</div>
      <div class="kpi-value" id="kpi-rev">—</div>
      <div class="kpi-unit">тыс. руб.</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-icon">📋</div>
      <div class="kpi-label">Заказов</div>
      <div class="kpi-value" id="kpi-ord">—</div>
      <div class="kpi-unit">шт.</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-icon">🏢</div>
      <div class="kpi-label">Клиентов</div>
      <div class="kpi-value" id="kpi-cli">—</div>
      <div class="kpi-unit">уникальных</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-icon">📊</div>
      <div class="kpi-label">Средний чек</div>
      <div class="kpi-value" id="kpi-avg">—</div>
      <div class="kpi-unit">тыс. руб.</div>
    </div>
  </div>

  <!-- Выручка по месяцам — полная ширина -->
  <div class="charts-grid">
    <div class="chart-card chart-full">
      <div class="chart-title">Выручка по месяцам</div>
      <div class="chart-wrap">
        <canvas id="chartMonthly"></canvas>
      </div>
    </div>

    <!-- Топ клиентов -->
    <div class="chart-card">
      <div class="chart-title">Топ-10 клиентов</div>
      <div class="chart-wrap-tall">
        <canvas id="chartClients"></canvas>
      </div>
    </div>

    <!-- Топ менеджеров -->
    <div class="chart-card">
      <div class="chart-title">Топ-10 менеджеров</div>
      <div class="chart-wrap-tall">
        <canvas id="chartManagers"></canvas>
      </div>
    </div>

    <!-- Отрасли -->
    <div class="chart-card chart-full">
      <div class="chart-title">Выручка по отраслям</div>
      <div class="chart-wrap">
        <canvas id="chartIndustries"></canvas>
      </div>
    </div>
  </div>
</main>

<footer>
  Фонтанка.ру · Коммерческий отдел · Аналитика заказов
</footer>

<script>
const DATA = {d};

// Meta
document.getElementById('fname').textContent = DATA.meta.filename;
document.getElementById('fdate').textContent = DATA.meta.generated;

// KPI
function fmt(n) {{
  return n.toLocaleString('ru-RU');
}}
document.getElementById('kpi-rev').textContent = fmt(DATA.kpi.revenue);
document.getElementById('kpi-ord').textContent = fmt(DATA.kpi.orders);
document.getElementById('kpi-cli').textContent = fmt(DATA.kpi.clients);
document.getElementById('kpi-avg').textContent = fmt(DATA.kpi.avg_check);

// Палитра
const ORANGE   = '#F38120';
const ORANGE_T = 'rgba(243,129,32,0.15)';
const BLUE     = '#60A5FA';
const GREEN    = '#22C55E';
const PURPLE   = '#A78BFA';
const PINK     = '#F472B6';
const TEAL     = '#2DD4BF';
const YELLOW   = '#FBBF24';
const COLORS   = [ORANGE, BLUE, GREEN, PURPLE, PINK, TEAL, YELLOW,
                  '#FB7185', '#34D399', '#818CF8'];

const DEFAULTS = {{
  color: '#AAAAAA',
  font: {{ family: 'Manrope', size: 11 }},
}};
Chart.defaults.color = DEFAULTS.color;
Chart.defaults.font  = DEFAULTS.font;

const GRID = {{
  color: 'rgba(255,255,255,0.05)',
  drawBorder: false,
}};

// ── 1. Выручка по месяцам ────────────────────────────────
if (DATA.monthly.length > 0) {{
  const ctx = document.getElementById('chartMonthly').getContext('2d');
  const gradient = ctx.createLinearGradient(0, 0, 0, 280);
  gradient.addColorStop(0, 'rgba(243,129,32,0.3)');
  gradient.addColorStop(1, 'rgba(243,129,32,0)');

  new Chart(ctx, {{
    type: 'line',
    data: {{
      labels: DATA.monthly.map(r => r.month),
      datasets: [{{
        label: 'Выручка, тыс. руб.',
        data: DATA.monthly.map(r => r.revenue),
        borderColor: ORANGE,
        backgroundColor: gradient,
        borderWidth: 2.5,
        pointBackgroundColor: ORANGE,
        pointBorderColor: '#0C0C0C',
        pointBorderWidth: 2,
        pointRadius: 5,
        pointHoverRadius: 8,
        fill: true,
        tension: 0.4,
      }}]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{
          backgroundColor: '#1E1E1E',
          borderColor: '#2A2A2A',
          borderWidth: 1,
          titleColor: '#F0F0F0',
          bodyColor: ORANGE,
          padding: 12,
          callbacks: {{
            label: ctx => ' ' + ctx.parsed.y.toLocaleString('ru-RU') + ' тыс. руб.'
          }}
        }}
      }},
      scales: {{
        x: {{ grid: GRID, ticks: {{ maxRotation: 45 }} }},
        y: {{ grid: GRID, ticks: {{
          callback: v => v.toLocaleString('ru-RU')
        }} }},
      }}
    }}
  }});
}}

// ── 2. Топ клиентов (горизонтальный) ────────────────────
if (DATA.top_clients.length > 0) {{
  const sorted = [...DATA.top_clients].sort((a,b) => a.revenue - b.revenue);
  new Chart(document.getElementById('chartClients'), {{
    type: 'bar',
    data: {{
      labels: sorted.map(r => r.name),
      datasets: [{{
        label: 'тыс. руб.',
        data: sorted.map(r => r.revenue),
        backgroundColor: COLORS.slice(0, sorted.length),
        borderRadius: 4,
        borderSkipped: false,
      }}]
    }},
    options: {{
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{
          backgroundColor: '#1E1E1E',
          borderColor: '#2A2A2A',
          borderWidth: 1,
          callbacks: {{
            label: ctx => ' ' + ctx.parsed.x.toLocaleString('ru-RU') + ' тыс. руб.'
          }}
        }}
      }},
      scales: {{
        x: {{ grid: GRID, ticks: {{ callback: v => v.toLocaleString('ru-RU') }} }},
        y: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 10 }} }} }},
      }}
    }}
  }});
}}

// ── 3. Топ менеджеров (вертикальный) ────────────────────
if (DATA.top_managers.length > 0) {{
  new Chart(document.getElementById('chartManagers'), {{
    type: 'bar',
    data: {{
      labels: DATA.top_managers.map(r => r.name),
      datasets: [{{
        label: 'тыс. руб.',
        data: DATA.top_managers.map(r => r.revenue),
        backgroundColor: ORANGE,
        hoverBackgroundColor: '#FFB060',
        borderRadius: 6,
        borderSkipped: false,
      }}]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{
          backgroundColor: '#1E1E1E',
          borderColor: '#2A2A2A',
          borderWidth: 1,
          callbacks: {{
            label: ctx => ' ' + ctx.parsed.y.toLocaleString('ru-RU') + ' тыс. руб.'
          }}
        }}
      }},
      scales: {{
        x: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 10 }}, maxRotation: 30 }} }},
        y: {{ grid: GRID, ticks: {{ callback: v => v.toLocaleString('ru-RU') }} }},
      }}
    }}
  }});
}}

// ── 4. Отрасли (пончик) ─────────────────────────────────
if (DATA.industries.length > 0) {{
  new Chart(document.getElementById('chartIndustries'), {{
    type: 'doughnut',
    data: {{
      labels: DATA.industries.map(r => r.name),
      datasets: [{{
        data: DATA.industries.map(r => r.revenue),
        backgroundColor: COLORS,
        borderColor: '#0C0C0C',
        borderWidth: 3,
        hoverOffset: 8,
      }}]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      cutout: '65%',
      plugins: {{
        legend: {{
          position: 'right',
          labels: {{ padding: 16, font: {{ size: 11 }}, usePointStyle: true }}
        }},
        tooltip: {{
          backgroundColor: '#1E1E1E',
          borderColor: '#2A2A2A',
          borderWidth: 1,
          callbacks: {{
            label: ctx => ' ' + ctx.parsed.toLocaleString('ru-RU') + ' тыс. руб.'
          }}
        }}
      }}
    }}
  }});
}}
</script>
</body>
</html>"""

    return html.replace('{d}', d)


def generate_dashboard(input_path: str, log=print) -> str:
    """
    Парсит файл, генерирует HTML-дашборд и открывает в браузере.
    Возвращает путь к HTML-файлу.
    """
    log(f"Загружаю данные: {os.path.basename(input_path)}")
    df = _load_and_prepare(input_path)
    log(f"Строк после фильтров: {len(df)}")

    log("Формирую дашборд...")
    data = _collect_data(df)

    # Сохраняем рядом с входным файлом
    out_dir = os.path.dirname(input_path)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(out_dir, f"dashboard_{ts}.html")

    html = _generate_html(data, input_path)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    log(f"✅ Дашборд сохранён: {out_path}")
    webbrowser.open(f"file:///{out_path.replace(os.sep, '/')}")
    return out_path
