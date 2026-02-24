// =============================================
// CONFIGURACIÓN DE MINERALES
// =============================================
const MINERALS = [
  { key: 'Oro', unit: 'USD/kg', color: '#f59e0b', gradient: ['rgba(245,158,11,.25)', 'rgba(245,158,11,0)'] },
  { key: 'Plata', unit: 'USD/kg', color: '#94a3b8', gradient: ['rgba(148,163,184,.25)', 'rgba(148,163,184,0)'] },
  { key: 'Cobre', unit: 'USD/kg', color: '#f97316', gradient: ['rgba(249,115,22,.25)', 'rgba(249,115,22,0)'] },
  { key: 'Hierro', unit: 'USD/kg', color: '#ef4444', gradient: ['rgba(239,68,68,.25)', 'rgba(239,68,68,0)'] },
  { key: 'Litio', unit: 'USD/kg', color: '#a855f7', gradient: ['rgba(168,85,247,.25)', 'rgba(168,85,247,0)'] },
  { key: 'Aluminio', unit: 'USD/kg', color: '#06b6d4', gradient: ['rgba(6,182,212,.25)', 'rgba(6,182,212,0)'] },
  { key: 'Níquel', unit: 'USD/kg', color: '#10b981', gradient: ['rgba(16,185,129,.25)', 'rgba(16,185,129,0)'] },
  { key: 'Platino', unit: 'USD/kg', color: '#7e8a9a', gradient: ['rgba(126,138,154,.25)', 'rgba(126,138,154,0)'] },
  { key: 'Paladio', unit: 'USD/kg', color: '#ec4899', gradient: ['rgba(236,72,153,.25)', 'rgba(236,72,153,0)'] },
  { key: 'Zinc', unit: 'USD/kg', color: '#84cc16', gradient: ['rgba(132,204,22,.25)', 'rgba(132,204,22,0)'] },
];

// =============================================
// DATOS EXACTOS DEL EXCEL 1976–2026
// =============================================
const rawData = [
  [1976,4018.8403,138.2481,1.4,0.012,1.5,0.8,4.8,5144.1156,1286.0289,0.7],
  [1977,7933.1908,269.2623,1.575,0.0135,1.5357,1,5.1,9323.7096,2572.0578,0.7125],
  [1978,11847.5413,400.2765,1.75,0.015,1.5714,1.2,5.4,13503.3035,3858.0867,0.725],
  [1979,15761.8917,531.2907,1.925,0.0165,1.6071,1.4,5.7,17682.8974,5144.1156,0.7375],
  [1980,19676.2422,662.3049,2.1,0.018,1.6429,1.6,6,21862.4914,6430.1445,0.75],
  [1981,18939.9907,611.5067,2.15,0.0186,1.6786,1.6,6.28,21187.3262,6172.9387,0.825],
  [1982,18203.7391,560.7086,2.2,0.0192,1.7143,1.6,6.56,20512.161,5915.733,0.9],
  [1983,17467.4876,509.9105,2.25,0.0198,1.75,1.6,6.84,19836.9958,5658.5272,0.975],
  [1984,16731.236,459.1123,2.3,0.0204,1.7857,1.6,7.12,19161.8307,5401.3214,1.05],
  [1985,15994.9845,408.3142,2.35,0.021,1.8214,1.6,7.4,18486.6655,5144.1156,1.125],
  [1986,15258.7329,357.516,2.4,0.0216,1.8571,1.6,7.68,17811.5003,4886.9098,1.2],
  [1987,14522.4814,306.7179,2.45,0.0222,1.8929,1.6,7.96,17136.3351,4629.7041,1.275],
  [1988,13786.2298,255.9198,2.5,0.0228,1.9286,1.6,8.24,16461.17,4372.4983,1.35],
  [1989,13049.9783,205.1216,2.55,0.0234,1.9643,1.6,8.52,15786.0048,4115.2925,1.425],
  [1990,12313.7268,154.3235,2.6,0.024,2,1.6,8.8,15110.8396,3858.0867,1.5],
  [1991,11979.3592,154.645,2.6667,0.0264,2.15,1.62,10.4588,17075.606,5261.0273,1.6062],
  [1992,11644.9917,154.9665,2.7333,0.0288,2.3,1.64,12.1176,19040.3724,6663.968,1.7125],
  [1993,11310.6242,155.288,2.8,0.0312,2.45,1.66,13.7765,21005.1388,8066.9086,1.8188],
  [1994,10976.2567,155.6095,2.8667,0.0336,2.6,1.68,15.4353,22969.9051,9469.8492,1.925],
  [1995,10641.8892,155.931,2.9333,0.036,2.75,1.7,17.0941,24934.6715,10872.7898,2.0312],
  [1996,10307.5217,156.2525,3,0.0384,2.9,1.72,18.7529,26899.4379,12275.7304,2.1375],
  [1997,9973.1541,156.574,3.0667,0.0408,3.05,1.74,20.4118,28864.2043,13678.6711,2.2438],
  [1998,9638.7866,156.8955,3.1333,0.0432,3.2,1.76,22.0706,30828.9707,15081.6117,2.35],
  [1999,9304.4191,157.217,3.2,0.0456,3.35,1.78,23.7294,32793.737,16484.5523,2.4562],
  [2000,8970.0516,157.5385,3.2667,0.048,3.5,1.8,25.3882,34758.5034,17887.4929,2.5625],
  [2001,12008.2949,245.8069,3.3333,0.0504,3.65,1.82,27.0471,36723.2698,19290.4336,2.6688],
  [2002,15046.5382,334.0752,3.4,0.0528,3.8,1.84,28.7059,38688.0362,19497.1168,2.775],
  [2003,18084.7815,422.3436,3.4667,0.0552,3.95,1.86,30.3647,40652.8026,19703.8,2.8812],
  [2004,21123.0247,510.6119,3.5333,0.0576,4.1,1.88,32.0235,42617.5689,19910.4832,2.9875],
  [2005,24161.268,598.8803,3.6,0.06,4.25,1.9,33.6824,44582.3353,20117.1664,3.0938],
  [2006,27199.5113,687.1486,4.4667,0.0778,4.4,1.9833,35.3412,46547.1017,20323.8496,3.2],
  [2007,30237.7546,775.417,5.3333,0.0957,4.55,2.0667,37,48511.8681,20530.5329,3.0556],
  [2008,33275.9979,863.6853,6.2,0.1135,4.7,2.15,33.85,50476.6345,20737.2161,2.9111],
  [2009,36314.2412,951.9537,7.0667,0.1313,4.85,2.2333,30.7,48088.2951,20943.8993,2.7667],
  [2010,39352.4844,1040.222,7.9333,0.1492,5,2.3167,27.55,45699.9557,21150.5825,2.6222],
  [2011,41107.9139,1128.4904,8.8,0.167,6.2857,2.4,24.4,43311.6163,21357.2657,2.4778],
  [2012,42863.3434,1076.3347,8.5,0.1604,7.5714,2.3222,21.25,40923.2769,21563.9489,2.3333],
  [2013,44618.7728,1024.1791,8.2,0.1539,8.8571,2.2444,18.1,38534.9375,21770.6322,2.1889],
  [2014,46374.2023,972.0235,7.9,0.1473,10.1429,2.1667,14.95,36146.5981,21977.3154,2.0444],
  [2015,48129.6317,919.8679,7.6,0.1408,11.4286,2.0889,11.8,33758.2587,22183.9986,1.9],
  [2016,49885.0612,867.7123,7.3,0.1342,12.7143,2.0111,13.6857,32665.1341,31346.9545,2.1143],
  [2017,51640.4906,815.5567,7,0.1277,14,1.9333,15.5714,31572.0096,40509.9105,2.3286],
  [2018,53395.9201,763.401,6.7,0.1211,12.3333,1.8556,17.4571,30478.885,49672.8664,2.5429],
  [2019,55151.3495,711.2454,6.4,0.1146,10.6667,1.7778,19.3429,29385.7604,58835.8223,2.7571],
  [2020,56906.779,659.0898,6.1,0.108,9,1.7,21.2286,28292.6359,67998.7783,2.9714],
  [2021,61568.6338,719.3724,6.95,0.1085,39.5,1.9,23.1143,29257.1576,77161.7342,3.1857],
  [2022,66230.4885,779.655,7.8,0.109,70,2.1,25,30221.6792,63229.7544,3.4],
  [2023,70892.3433,839.9376,8.65,0.1095,44,2.3,21.25,31186.2009,49297.7746,3.1],
  [2024,75554.1981,900.2202,9.5,0.11,18,2.5,17.5,32150.7226,35365.7948,2.8],
  [2025,118957.6736,1768.2897,11.05,0.105,19.4,2.8,17.5,50798.1417,45654.0261,3.085],
  [2026,162361.1491,2636.3593,12.6,0.1,20.8,3.1,17.5,69445.5608,55942.2573,3.37],
];

const years = rawData.map(r => r[0]);
const mineralData = {};
MINERALS.forEach((m, i) => { mineralData[m.key] = rawData.map(r => r[i + 1]); });

function fmt(v) {
  if (v >= 1e6) return (v / 1e6).toFixed(1) + 'M';
  if (v >= 1e3) return (v / 1e3).toFixed(1) + 'K';
  if (v >= 1) return v.toFixed(2);
  return v.toFixed(4);
}

// =============================================
// THEME TOGGLE
// =============================================
let currentTheme = 'dark';
let generalChart = null;
let individualCharts = [];
let currentScaleType = 'logarithmic';

function getThemeColors() {
  const isLight = currentTheme === 'light';
  return {
    gridColor: isLight ? 'rgba(0,0,0,0.06)' : 'rgba(255,255,255,0.04)',
    tickColor: isLight ? '#475569' : '#64748b',
    tooltipBg: isLight ? 'rgba(255,255,255,0.95)' : 'rgba(17,24,39,0.95)',
    tooltipTitle: isLight ? '#0f172a' : '#f1f5f9',
    tooltipBorder: isLight ? 'rgba(0,0,0,0.1)' : 'rgba(255,255,255,0.1)',
  };
}

function toggleTheme() {
  currentTheme = currentTheme === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', currentTheme === 'light' ? 'light' : '');
  document.getElementById('themeIcon').textContent = currentTheme === 'dark' ? '🌙' : '☀️';
  document.getElementById('themeLabel').textContent = currentTheme === 'dark' ? 'Oscuro' : 'Claro';
  rebuildAllCharts();
}

// =============================================
// BUILD CHARTS
// =============================================
function buildGeneralChart() {
  const tc = getThemeColors();
  const ctx = document.getElementById('generalChart').getContext('2d');
  const datasets = MINERALS.map(m => ({
    label: m.key, data: mineralData[m.key], borderColor: m.color, backgroundColor: 'transparent',
    borderWidth: 2, tension: 0.35, pointRadius: 0, pointHitRadius: 10,
    pointHoverRadius: 5, pointHoverBackgroundColor: m.color, pointHoverBorderColor: '#fff', pointHoverBorderWidth: 2,
  }));

  generalChart = new Chart(ctx, {
    type: 'line',
    data: { labels: years, datasets: datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { backgroundColor: tc.tooltipBg, titleColor: tc.tooltipTitle, bodyColor: tc.tooltipTitle, borderColor: tc.tooltipBorder, borderWidth: 1, cornerRadius: 8, padding: 12, mode: 'index', intersect: false,
          callbacks: { title: c => `Año ${c[0].label}`, label: c => `${c.dataset.label}: ${fmt(c.parsed.y)} USD/kg` }
        }
      },
      scales: {
        x: { grid: { color: tc.gridColor }, ticks: { color: tc.tickColor, font: { size: 11 }, autoSkip: true, maxRotation: 0,
          callback: function(value, index) { const y = years[index]; if (y === 1976 || y === 2026 || y % 5 === 0) return y; return ''; }
        }, border: { display: false } },
        y: { type: currentScaleType, grid: { color: tc.gridColor }, ticks: { color: tc.tickColor, font: { size: 11 }, callback: v => fmt(v) }, border: { display: false } }
      },
      interaction: { intersect: false, mode: 'index' }
    }
  });
}

function buildIndividualCharts() {
  const tc = getThemeColors();
  const grid = document.getElementById('chartGrid');
  grid.innerHTML = '';
  individualCharts = [];

  MINERALS.forEach((m, idx) => {
    const vals = mineralData[m.key];
    const maxVal = Math.max(...vals), minVal = Math.min(...vals);
    const avgVal = vals.reduce((a, b) => a + b, 0) / vals.length;
    const current = vals[vals.length - 1];

    const card = document.createElement('div');
    card.className = 'chart-card';
    card.style.animationDelay = `${idx * 0.06}s`;
    card.innerHTML = `
      <div class="chart-header">
        <div><div class="mineral-name"><span class="dot" style="background:${m.color}"></span>${m.key}</div><div class="mineral-unit">${m.unit} · 1976–2026</div></div>
        <div class="mineral-stats">
          <div class="stat"><div class="stat-label">Máx</div><div class="stat-val" style="color:${m.color}">${fmt(maxVal)}</div></div>
          <div class="stat"><div class="stat-label">Mín</div><div class="stat-val">${fmt(minVal)}</div></div>
          <div class="stat"><div class="stat-label">Prom</div><div class="stat-val">${fmt(avgVal)}</div></div>
          <div class="stat"><div class="stat-label">2026</div><div class="stat-val" style="color:#16a34a">${fmt(current)}</div></div>
        </div>
      </div>
      <div class="chart-canvas-wrap"><canvas id="chart-${idx}"></canvas></div>
    `;
    grid.appendChild(card);

    const ctx = document.getElementById(`chart-${idx}`).getContext('2d');
    const grad = ctx.createLinearGradient(0, 0, 0, 210);
    grad.addColorStop(0, m.gradient[0]); grad.addColorStop(1, m.gradient[1]);

    const chart = new Chart(ctx, {
      type: 'line',
      data: { labels: years, datasets: [{ data: vals, borderColor: m.color, backgroundColor: grad, borderWidth: 2, fill: true, tension: 0.35, pointRadius: 0, pointHitRadius: 10, pointHoverRadius: 5, pointHoverBackgroundColor: m.color, pointHoverBorderColor: '#fff', pointHoverBorderWidth: 2 }] },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { backgroundColor: tc.tooltipBg, titleColor: tc.tooltipTitle, bodyColor: tc.tooltipTitle, borderColor: m.color, borderWidth: 1, cornerRadius: 8, padding: 10, displayColors: false,
          callbacks: { title: c => `Año ${c[0].label}`, label: c => `${m.key}: ${fmt(c.parsed.y)} ${m.unit}` }
        }},
        scales: {
          x: { grid: { color: tc.gridColor }, ticks: { color: tc.tickColor, font: { size: 10 }, maxTicksLimit: 12 }, border: { display: false } },
          y: { grid: { color: tc.gridColor }, ticks: { color: tc.tickColor, font: { size: 10 }, callback: v => fmt(v) }, border: { display: false } }
        },
        interaction: { intersect: false, mode: 'index' }
      }
    });
    individualCharts.push(chart);
  });
}

function rebuildAllCharts() {
  if (generalChart) generalChart.destroy();
  individualCharts.forEach(c => c.destroy());
  buildGeneralChart();
  buildIndividualCharts();
  rebuildLegend();
}

// =============================================
// SCALE TOGGLE
// =============================================
function setScale(type) {
  currentScaleType = type;
  generalChart.options.scales.y.type = type;
  generalChart.update();
  document.getElementById('btnLinear').classList.toggle('active', type === 'linear');
  document.getElementById('btnLog').classList.toggle('active', type === 'logarithmic');
}

// =============================================
// LEGEND
// =============================================
function rebuildLegend() {
  const container = document.getElementById('generalLegend');
  container.innerHTML = '';
  MINERALS.forEach((m, i) => {
    const item = document.createElement('div');
    item.className = 'legend-item';
    item.innerHTML = `<span class="legend-dot" style="background:${m.color}"></span>${m.key}`;
    item.addEventListener('click', () => {
      const meta = generalChart.getDatasetMeta(i);
      meta.hidden = !meta.hidden;
      item.classList.toggle('hidden', meta.hidden);
      generalChart.update();
    });
    container.appendChild(item);
  });
}

// =============================================
// KPIs
// =============================================
function buildKPIs() {
  const strip = document.getElementById('kpiStrip');
  MINERALS.forEach(m => {
    const vals = mineralData[m.key];
    const cur = vals[vals.length - 1], prev = vals[vals.length - 2];
    const ch = ((cur - prev) / prev * 100).toFixed(1);
    const up = ch >= 0;
    const card = document.createElement('div');
    card.className = 'kpi-card';
    card.innerHTML = `
      <div class="kpi-label" style="color:${m.color}">${m.key}</div>
      <div class="kpi-value" style="color:${m.color}">${fmt(cur)}</div>
      <div class="kpi-change ${up ? 'up' : 'down'}">${up ? '▲' : '▼'} ${Math.abs(ch)}% vs 2025</div>
    `;
    strip.appendChild(card);
  });
}

// =============================================
// PRODUCERS TABLE
// =============================================
function buildTable() {
  const producersData = [
    { mineral: 'Oro', producers: ['China (380 t)', 'Rusia (330 t)', 'Australia (290 t)', 'Canadá (200 t)', 'EE.UU. (158 t)', 'Indonesia (140 t)', 'Ghana (130 t)', 'Sudáfrica (120 t)', 'México (120 t)', 'Uzbekistán (110 t)'] },
    { mineral: 'Plata', producers: ['México (6.120 t)', 'Perú (4.160 t)', 'China (3.570 t)', 'Chile (1.600 t)', 'Rusia (1.500 t)', 'Bolivia (1.400 t)', 'Polonia (1.300 t)', 'Australia (1.200 t)', 'Argentina (1.000 t)', 'EE.UU. (970 t)'] },
    { mineral: 'Cobre', producers: ['Chile (5.300 kt)', 'R.D. Congo (3.300 kt)', 'Perú (2.600 kt)', 'China (1.800 kt)', 'EE.UU. (1.100 kt)', 'Indonesia (1.000 kt)', 'Australia (800 kt)', 'Rusia (780 kt)', 'Kazajistán (740 kt)', 'México (700 kt)'] },
    { mineral: 'Hierro', producers: ['Australia (960 Mt)', 'Brasil (380 Mt)', 'China (340 Mt)', 'India (310 Mt)', 'Rusia (100 Mt)', 'Sudáfrica (61 Mt)', 'Ucrania (53 Mt)', 'Canadá (50 Mt)', 'EE.UU. (40 Mt)', 'Kazajistán (38 Mt)'] },
    { mineral: 'Litio', producers: ['Australia (55.000 t)', 'Chile (26.000 t)', 'China (14.000 t)', 'Argentina (9.600 t)', 'Brasil (4.900 t)', 'Zimbabwe (3.400 t)', 'Portugal (1.100 t)', 'Canadá (900 t)', 'Namibia (500 t)', 'Nigeria (300 t)'] },
    { mineral: 'Aluminio', producers: ['China (36.000 kt)', 'India (3.700 kt)', 'Rusia (3.600 kt)', 'Canadá (3.200 kt)', 'EAU (2.700 kt)', 'Australia (1.600 kt)', 'Bahréin (1.500 kt)', 'Noruega (1.400 kt)', 'EE.UU. (700 kt)', 'Islandia (680 kt)'] },
    { mineral: 'Níquel', producers: ['Indonesia (800.000 t)', 'Filipinas (370.000 t)', 'Rusia (270.000 t)', 'N. Caledonia (190.000 t)', 'Australia (150.000 t)', 'Canadá (130.000 t)', 'China (120.000 t)', 'Brasil (83.000 t)', 'Guatemala (60.000 t)', 'Cuba (50.000 t)'] },
    { mineral: 'Platino', producers: ['Sudáfrica (120.000 kg)', 'Rusia (23.000 kg)', 'Zimbabwe (16.000 kg)', 'Canadá (7.000 kg)', 'EE.UU. (4.000 kg)', 'Colombia (1.900 kg)', 'Botswana (1.200 kg)', 'Finlandia (700 kg)', 'Australia (300 kg)', 'China (200 kg)'] },
    { mineral: 'Paladio', producers: ['Rusia (92.000 kg)', 'Sudáfrica (76.000 kg)', 'Canadá (16.000 kg)', 'EE.UU. (12.000 kg)', 'Zimbabwe (11.000 kg)', 'Botswana (2.700 kg)', 'Colombia (1.400 kg)', 'Finlandia (900 kg)', 'Australia (600 kg)', 'China (500 kg)'] },
    { mineral: 'Zinc', producers: ['China (4.100 kt)', 'Australia (1.200 kt)', 'Perú (1.100 kt)', 'India (850 kt)', 'EE.UU. (700 kt)', 'México (660 kt)', 'Bolivia (500 kt)', 'Kazajistán (400 kt)', 'Suecia (340 kt)', 'Turquía (280 kt)'] },
  ];

  const tbody = document.querySelector('#producersTable tbody');
  producersData.forEach((row, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td class="rank-cell">${i + 1}</td><td class="mineral-cell" style="color:${MINERALS[i].color}">${row.mineral}</td>` +
      row.producers.map((p, j) => `<td>${j === 0 ? '<strong>' + p + '</strong>' : p}</td>`).join('');
    tbody.appendChild(tr);
  });
}

// =============================================
// INIT
// =============================================
buildKPIs();
buildGeneralChart();
rebuildLegend();
buildIndividualCharts();
buildTable();
