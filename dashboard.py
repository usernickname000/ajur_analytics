# ============================================================
# dashboard.py — генерация HTML-дашборда
# ============================================================

import os
import json
import webbrowser
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

    monthly = []
    if '_month_dt' in df.columns:
        grp = (df.dropna(subset=['_month_dt'])
                 .groupby(df['_month_dt'].dt.strftime('%m.%Y'))
                 [COL_REVENUE].sum()
                 .reset_index())
        grp.columns = ['month', 'revenue']
        grp['revenue'] = (grp['revenue'] / 1000).round(1)
        monthly = grp.to_dict('records')

    top_cli = (df.groupby('КОНЕЧНЫЙ_КЛИЕНТ')[COL_REVENUE]
                 .sum().nlargest(10).reset_index())
    top_cli.columns = ['name', 'revenue']
    top_cli['revenue'] = (top_cli['revenue'] / 1000).round(1)
    top_cli['name'] = top_cli['name'].astype(str).str[:35]

    top_mgr = pd.DataFrame()
    if COL_MANAGER in df.columns:
        top_mgr = (df.groupby(COL_MANAGER)[COL_REVENUE]
                     .sum().nlargest(10).reset_index())
        top_mgr.columns = ['name', 'revenue']
        top_mgr['revenue'] = (top_mgr['revenue'] / 1000).round(1)

    industries = (df.groupby('ОТРАСЛЬ_НОРМ')[COL_REVENUE]
                    .sum().nlargest(8).reset_index())
    industries.columns = ['name', 'revenue']
    industries['revenue'] = (industries['revenue'] / 1000).round(1)
    industries = industries[industries['name'].astype(str).str.strip() != '']

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
    }


def _generate_html(data: dict, filename: str) -> str:
    data['meta']['filename'] = os.path.basename(filename)
    d = json.dumps(data, ensure_ascii=False)

    html = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Аналитика — Фонтанка.ру</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>
  /* ── Reset & Base ── */
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --orange:      #F38120;
    --orange-glow: rgba(243,129,32,0.25);
    --orange-dim:  rgba(243,129,32,0.10);
    --bg:          #080808;
    --bg2:         #0E0E0E;
    --surface:     #141414;
    --surface2:    #1C1C1C;
    --surface3:    #242424;
    --border:      #262626;
    --border2:     #333;
    --text:        #F2F2F2;
    --text2:       #999;
    --text3:       #555;
    --green:       #22C55E;
    --blue:        #3B82F6;
    --purple:      #8B5CF6;
    --pink:        #EC4899;
    --teal:        #14B8A6;
    --amber:       #F59E0B;
    --red:         #EF4444;
    --radius:      12px;
    --radius-lg:   18px;
    --shadow:      0 0 0 1px var(--border);
    --shadow-lg:   0 8px 32px rgba(0,0,0,0.5);
  }

  html { font-size: 14px; scroll-behavior: smooth; }

  body {
    font-family: 'Inter', system-ui, -apple-system, sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    -webkit-font-smoothing: antialiased;
  }

  /* ── Scrollbar ── */
  ::-webkit-scrollbar { width: 5px; height: 5px; }
  ::-webkit-scrollbar-track { background: var(--bg); }
  ::-webkit-scrollbar-thumb { background: var(--border2); border-radius: 3px; }
  ::-webkit-scrollbar-thumb:hover { background: #444; }

  /* ── Header ── */
  header {
    position: sticky;
    top: 0;
    z-index: 100;
    background: rgba(8,8,8,0.85);
    backdrop-filter: blur(24px) saturate(1.8);
    -webkit-backdrop-filter: blur(24px) saturate(1.8);
    border-bottom: 1px solid var(--border);
    height: 64px;
    display: flex;
    align-items: center;
    padding: 0 40px;
    gap: 20px;
  }

  .logo-mark {
    width: 36px; height: 36px;
    background: var(--orange);
    border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    font-family: Georgia, serif;
    font-size: 16px; font-weight: bold; font-style: italic;
    color: #fff;
    flex-shrink: 0;
  }

  .logo-text {
    font-family: Georgia, serif;
    font-size: 18px;
    font-style: italic;
    font-weight: bold;
    color: var(--text);
    letter-spacing: -0.3px;
  }

  .logo-dot { color: var(--orange); }

  .header-sep {
    width: 1px; height: 24px;
    background: var(--border2);
    margin: 0 4px;
  }

  .header-badge {
    font-size: 10px;
    font-weight: 600;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: var(--text3);
  }

  .header-right {
    margin-left: auto;
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 2px;
  }

  .header-file {
    font-size: 11px;
    color: var(--text2);
  }

  .header-file span {
    color: var(--orange);
    font-weight: 500;
  }

  /* ── Layout ── */
  main {
    max-width: 1360px;
    margin: 0 auto;
    padding: 36px 40px 60px;
  }

  /* ── Section title ── */
  .section-title {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 16px;
  }

  .section-title::before {
    content: '';
    width: 3px;
    height: 14px;
    background: var(--orange);
    border-radius: 2px;
    flex-shrink: 0;
  }

  .section-title h2 {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: var(--text2);
  }

  /* ── KPI grid ── */
  .kpi-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 12px;
    margin-bottom: 40px;
  }

  .kpi-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius-lg);
    padding: 22px 20px;
    position: relative;
    overflow: hidden;
    cursor: default;
    transition: border-color .2s, transform .2s, box-shadow .2s;
  }

  .kpi-card:hover {
    border-color: var(--card-color, var(--orange));
    transform: translateY(-2px);
    box-shadow: 0 12px 40px rgba(0,0,0,.4),
                0 0 0 1px var(--card-color, var(--orange));
  }

  /* Фоновый градиент-кружок */
  .kpi-card::after {
    content: '';
    position: absolute;
    right: -20px; top: -20px;
    width: 100px; height: 100px;
    background: radial-gradient(circle, var(--card-color, var(--orange)) 0%, transparent 70%);
    opacity: .06;
    pointer-events: none;
  }

  .kpi-card .card-accent {
    position: absolute;
    left: 0; top: 0; bottom: 0;
    width: 3px;
    background: var(--card-color, var(--orange));
    border-radius: 2px 0 0 2px;
  }

  .kpi-label {
    font-size: 10px;
    font-weight: 600;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: var(--text3);
    margin-bottom: 10px;
  }

  .kpi-value {
    font-size: 30px;
    font-weight: 800;
    color: var(--text);
    line-height: 1;
    letter-spacing: -1px;
    margin-bottom: 4px;
  }

  .kpi-unit {
    font-size: 11px;
    color: var(--text2);
    font-weight: 400;
  }

  .kpi-trend {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    margin-top: 8px;
    font-size: 11px;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 20px;
  }

  .kpi-trend.up   { color: var(--green); background: rgba(34,197,94,.1); }
  .kpi-trend.down { color: var(--red);   background: rgba(239,68,68,.1); }

  /* ── Chart cards ── */
  .charts-section { margin-bottom: 40px; }

  .charts-grid {
    display: grid;
    gap: 12px;
  }

  .charts-row-2 { grid-template-columns: 1fr 1fr; }
  .charts-row-1 { grid-template-columns: 1fr; }

  .chart-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius-lg);
    padding: 24px;
    transition: border-color .2s;
  }

  .chart-card:hover { border-color: var(--border2); }

  .chart-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 20px;
  }

  .chart-title {
    font-size: 12px;
    font-weight: 700;
    letter-spacing: .5px;
    color: var(--text2);
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .chart-title::before {
    content: '';
    width: 3px; height: 12px;
    background: var(--orange);
    border-radius: 2px;
  }

  .chart-badge {
    font-size: 10px;
    font-weight: 600;
    color: var(--text3);
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 3px 8px;
    letter-spacing: .5px;
  }

  .chart-wrap        { position: relative; height: 260px; }
  .chart-wrap-tall   { position: relative; height: 320px; }
  .chart-wrap-short  { position: relative; height: 200px; }

  /* ── Footer ── */
  footer {
    text-align: center;
    padding: 32px 40px;
    color: var(--text3);
    font-size: 11px;
    border-top: 1px solid var(--border);
    letter-spacing: .5px;
  }

  /* ── Animations ── */
  @keyframes fadeUp {
    from { opacity: 0; transform: translateY(16px); }
    to   { opacity: 1; transform: translateY(0); }
  }

  @keyframes countUp {
    from { opacity: 0; }
    to   { opacity: 1; }
  }

  .kpi-card { animation: fadeUp .5s ease both; }
  .kpi-card:nth-child(1) { animation-delay: .05s; }
  .kpi-card:nth-child(2) { animation-delay: .10s; }
  .kpi-card:nth-child(3) { animation-delay: .15s; }
  .kpi-card:nth-child(4) { animation-delay: .20s; }
  .chart-card { animation: fadeUp .6s ease both; animation-delay: .3s; }

  /* ── Responsive ── */
  @media (max-width: 900px) {
    .kpi-grid { grid-template-columns: repeat(2, 1fr); }
    .charts-row-2 { grid-template-columns: 1fr; }
    main { padding: 20px; }
    header { padding: 0 20px; }
  }
</style>
</head>
<body>

<header>
  <div class="logo-mark">Ф</div>
  <span class="logo-text">фонтанка<span class="logo-dot">.</span>ру</span>
  <div class="header-sep"></div>
  <span class="header-badge">Коммерческий отдел · Аналитика</span>
  <div class="header-right">
    <div class="header-file">Файл: <span id="fname"></span></div>
    <div class="header-file">Сформировано: <span id="fdate"></span></div>
  </div>
</header>

<main>

  <!-- KPI -->
  <div style="margin-bottom:32px">
    <div class="section-title"><h2>Ключевые показатели</h2></div>
    <div class="kpi-grid" id="kpi-grid">
      <div class="kpi-card" style="--card-color: #F38120">
        <div class="card-accent"></div>
        <div class="kpi-label">Выручка</div>
        <div class="kpi-value" id="kpi-rev">—</div>
        <div class="kpi-unit">тыс. руб.</div>
        <div id="kpi-trend" style="display:none" class="kpi-trend"></div>
      </div>
      <div class="kpi-card" style="--card-color: #3B82F6">
        <div class="card-accent"></div>
        <div class="kpi-label">Заказов</div>
        <div class="kpi-value" id="kpi-ord">—</div>
        <div class="kpi-unit">в периоде</div>
      </div>
      <div class="kpi-card" style="--card-color: #22C55E">
        <div class="card-accent"></div>
        <div class="kpi-label">Клиентов</div>
        <div class="kpi-value" id="kpi-cli">—</div>
        <div class="kpi-unit">уникальных</div>
      </div>
      <div class="kpi-card" style="--card-color: #8B5CF6">
        <div class="card-accent"></div>
        <div class="kpi-label">Средний чек</div>
        <div class="kpi-value" id="kpi-avg">—</div>
        <div class="kpi-unit">тыс. руб.</div>
      </div>
    </div>
  </div>

  <!-- Выручка по месяцам -->
  <div class="charts-section">
    <div class="section-title"><h2>Динамика выручки</h2></div>
    <div class="charts-grid charts-row-1">
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">Выручка по месяцам</div>
          <div class="chart-badge">тыс. руб.</div>
        </div>
        <div class="chart-wrap">
          <canvas id="chartMonthly"></canvas>
        </div>
      </div>
    </div>
  </div>

  <!-- Клиенты и менеджеры -->
  <div class="charts-section">
    <div class="section-title"><h2>Топ-10</h2></div>
    <div class="charts-grid charts-row-2">
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">Клиенты</div>
          <div class="chart-badge">топ-10</div>
        </div>
        <div class="chart-wrap-tall">
          <canvas id="chartClients"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">Менеджеры</div>
          <div class="chart-badge">топ-10</div>
        </div>
        <div class="chart-wrap-tall">
          <canvas id="chartManagers"></canvas>
        </div>
      </div>
    </div>
  </div>

  <!-- Отрасли -->
  <div class="charts-section">
    <div class="section-title"><h2>Структура по отраслям</h2></div>
    <div class="charts-grid charts-row-1">
      <div class="chart-card">
        <div class="chart-header">
          <div class="chart-title">Выручка по отраслям</div>
          <div class="chart-badge">тыс. руб.</div>
        </div>
        <div class="chart-wrap">
          <canvas id="chartIndustries"></canvas>
        </div>
      </div>
    </div>
  </div>

</main>

<footer>
  Фонтанка.ру &nbsp;·&nbsp; Коммерческий отдел &nbsp;·&nbsp; Аналитика заказов
</footer>

<script>
const DATA = __DATA_PLACEHOLDER__;

// Meta
document.getElementById('fname').textContent = DATA.meta.filename;
document.getElementById('fdate').textContent = DATA.meta.generated;

// KPI
function fmt(n) { return n.toLocaleString('ru-RU'); }

document.getElementById('kpi-rev').textContent = fmt(DATA.kpi.revenue);
document.getElementById('kpi-ord').textContent = fmt(DATA.kpi.orders);
document.getElementById('kpi-cli').textContent = fmt(DATA.kpi.clients);
document.getElementById('kpi-avg').textContent = fmt(DATA.kpi.avg_check);

// Тренд
if (DATA.kpi.trend_pct !== null && DATA.kpi.trend_pct !== undefined) {
  const trendEl = document.getElementById('kpi-trend');
  const up = DATA.kpi.trend_pct >= 0;
  trendEl.textContent = (up ? '↑ +' : '↓ ') + DATA.kpi.trend_pct + '% к пред. месяцу';
  trendEl.className = 'kpi-trend ' + (up ? 'up' : 'down');
  trendEl.style.display = 'inline-flex';
}

// ── Глобальные настройки Chart.js ────────────────────────
Chart.defaults.color          = '#666';
Chart.defaults.font.family    = "'Inter', system-ui, sans-serif";
Chart.defaults.font.size      = 11;

const GRID = {
  color: 'rgba(255,255,255,0.04)',
  drawBorder: false,
};

const TOOLTIP = {
  backgroundColor: '#1C1C1C',
  borderColor: '#333',
  borderWidth: 1,
  titleColor: '#CCC',
  bodyColor: '#F38120',
  padding: 12,
  cornerRadius: 8,
  displayColors: false,
};

const PALETTE = [
  '#F38120', '#3B82F6', '#22C55E', '#8B5CF6',
  '#EC4899', '#14B8A6', '#F59E0B', '#EF4444',
  '#6366F1', '#10B981',
];

// ── 1. Выручка по месяцам ────────────────────────────────
if (DATA.monthly.length > 0) {
  const ctx = document.getElementById('chartMonthly').getContext('2d');

  const grad = ctx.createLinearGradient(0, 0, 0, 260);
  grad.addColorStop(0,   'rgba(243,129,32,0.22)');
  grad.addColorStop(0.6, 'rgba(243,129,32,0.04)');
  grad.addColorStop(1,   'rgba(243,129,32,0)');

  new Chart(ctx, {
    type: 'line',
    data: {
      labels: DATA.monthly.map(r => r.month),
      datasets: [{
        data: DATA.monthly.map(r => r.revenue),
        borderColor: '#F38120',
        backgroundColor: grad,
        borderWidth: 2,
        pointBackgroundColor: '#F38120',
        pointBorderColor: '#080808',
        pointBorderWidth: 2,
        pointRadius: 4,
        pointHoverRadius: 7,
        pointHoverBackgroundColor: '#FFB060',
        fill: true,
        tension: 0.4,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          ...TOOLTIP,
          callbacks: {
            label: c => ' ' + c.parsed.y.toLocaleString('ru-RU') + ' тыс. руб.'
          }
        }
      },
      scales: {
        x: {
          grid: GRID,
          ticks: { maxRotation: 40, color: '#555' },
          border: { display: false },
        },
        y: {
          grid: GRID,
          ticks: {
            callback: v => v.toLocaleString('ru-RU'),
            color: '#555',
          },
          border: { display: false },
        },
      }
    }
  });
}

// ── 2. Топ клиентов ──────────────────────────────────────
if (DATA.top_clients.length > 0) {
  const sorted = [...DATA.top_clients].sort((a,b) => a.revenue - b.revenue);
  new Chart(document.getElementById('chartClients'), {
    type: 'bar',
    data: {
      labels: sorted.map(r => r.name),
      datasets: [{
        data: sorted.map(r => r.revenue),
        backgroundColor: PALETTE.slice(0, sorted.length).map(c => c + 'CC'),
        hoverBackgroundColor: PALETTE.slice(0, sorted.length),
        borderRadius: 4,
        borderSkipped: false,
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          ...TOOLTIP,
          callbacks: {
            label: c => ' ' + c.parsed.x.toLocaleString('ru-RU') + ' тыс. руб.'
          }
        }
      },
      scales: {
        x: {
          grid: GRID,
          ticks: { callback: v => v.toLocaleString('ru-RU'), color: '#555' },
          border: { display: false },
        },
        y: {
          grid: { display: false },
          ticks: { font: { size: 10 }, color: '#777' },
          border: { display: false },
        },
      }
    }
  });
}

// ── 3. Топ менеджеров ────────────────────────────────────
if (DATA.top_managers.length > 0) {
  const mgr = DATA.top_managers;
  new Chart(document.getElementById('chartManagers'), {
    type: 'bar',
    data: {
      labels: mgr.map(r => r.name),
      datasets: [{
        data: mgr.map(r => r.revenue),
        backgroundColor: 'rgba(243,129,32,0.75)',
        hoverBackgroundColor: '#F38120',
        borderRadius: 6,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          ...TOOLTIP,
          callbacks: {
            label: c => ' ' + c.parsed.y.toLocaleString('ru-RU') + ' тыс. руб.'
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { font: { size: 10 }, maxRotation: 35, color: '#777' },
          border: { display: false },
        },
        y: {
          grid: GRID,
          ticks: { callback: v => v.toLocaleString('ru-RU'), color: '#555' },
          border: { display: false },
        },
      }
    }
  });
}

// ── 4. Отрасли (горизонтальный bar) ─────────────────────
if (DATA.industries.length > 0) {
  const total = DATA.industries.reduce((s, r) => s + r.revenue, 0);
  const sorted = [...DATA.industries].sort((a,b) => a.revenue - b.revenue);
  new Chart(document.getElementById('chartIndustries'), {
    type: 'bar',
    data: {
      labels: sorted.map(r => r.name),
      datasets: [{
        data: sorted.map(r => r.revenue),
        backgroundColor: PALETTE.slice(0, sorted.length).map(c => c + 'BB'),
        hoverBackgroundColor: PALETTE.slice(0, sorted.length),
        borderRadius: 5,
        borderSkipped: false,
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          ...TOOLTIP,
          callbacks: {
            label: c => {
              const pct = total > 0 ? ((c.parsed.x / total) * 100).toFixed(1) : '0';
              return ' ' + c.parsed.x.toLocaleString('ru-RU') + ' тыс. руб.  (' + pct + '%)';
            }
          }
        }
      },
      scales: {
        x: {
          grid: GRID,
          ticks: { callback: v => v.toLocaleString('ru-RU'), color: '#555' },
          border: { display: false },
        },
        y: {
          grid: { display: false },
          ticks: { font: { size: 10 }, color: '#777' },
          border: { display: false },
        },
      }
    }
  });
}
</script>
</body>
</html>"""

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
