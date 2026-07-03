import Chart from 'chart.js/auto';
import * as XLSX from 'xlsx';
import './style.css';

// ─── ONE-TIME (NON-DAILY) ITEM DETECTION ────────────────────────────
// Keywords that indicate an item is a one-time supply, NOT consumed daily.
// These items should NOT get the ×30 projection multiplier.
const ONE_TIME_KEYWORDS = [
  'tongkat', 'kursi roda', 'kasur', 'tikar', 'tenda', 'terpal', 'tarpal',
  'selimut', 'kelambu', 'lampu', 'tabung o2',
  'alat bantu jalan', 'alat bantu dengar', 'alat bantu tangan',
  'terapi', 'rujukan', 'pengobatan lumpuh', 'pengobatan lanjut',
  'pemeriksaan mata',
];

/**
 * Check if a need/item name is a one-time (non-daily) supply.
 * Uses case-insensitive keyword matching.
 */
function isOneTimeItem(needName) {
  const lower = needName.toLowerCase();
  return ONE_TIME_KEYWORDS.some(kw => lower.includes(kw));
}

/**
 * Get the projected count for an item.
 * Daily consumables get ×30; one-time items return the raw count.
 */
function getProjection(needName, dailyCount) {
  return isOneTimeItem(needName) ? dailyCount : dailyCount * 30;
}

/** Label suffix for projection display */
function getProjectionLabel(needName) {
  return isOneTimeItem(needName) ? 'pcs' : 'pcs/30d';
}

// ─── APPLICATION STATE ──────────────────────────────────────────────
let allData = [];
let filteredData = [];
let currentPage = 1;
const pageSize = 15;

let searchQuery = '';
let isCaseSensitive = false;
let selectedDesa = 'all';
let selectedDusun = 'all';
let selectedCategories = new Set(); // multi-select from data

let needsChart = null;
let cohortChart = null;
let specificNeedsChart = null;
let desaDistributionChart = null;

// ─── INIT ───────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  initUI();
  rebuildDynamicFilters();
  applyAndRender();
});

// ─── Dynamically derive unique values from current dataset ──────────
function getUniqueDesa() {
  return [...new Set(allData.map(r => r.desa))].sort();
}
function getUniqueDusun() {
  let source = allData;
  if (selectedDesa !== 'all') source = source.filter(r => r.desa === selectedDesa);
  return [...new Set(source.map(r => r.dusun))].sort((a,b) => {
    const na = parseInt(a), nb = parseInt(b);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return a.localeCompare(b);
  });
}
function getUniqueCategories() {
  const cats = new Set();
  allData.forEach(r => r.kategori_rentan.forEach(c => cats.add(c)));
  return [...cats].sort();
}

// ─── Build sidebar filters from data ────────────────────────────────
function rebuildDynamicFilters() {
  // Desa dropdown
  const desaSel = document.getElementById('desa-select');
  const prevDesa = desaSel.value;
  desaSel.innerHTML = '<option value="all">Semua Desa</option>';
  getUniqueDesa().forEach(d => {
    const opt = document.createElement('option');
    opt.value = d; opt.textContent = d;
    desaSel.appendChild(opt);
  });
  desaSel.value = prevDesa;

  // Dusun dropdown
  rebuildDusunDropdown();

  // Categories checkboxes
  const catContainer = document.getElementById('categories-filter-list');
  catContainer.innerHTML = '';
  selectedCategories.clear();
  getUniqueCategories().forEach(cat => {
    selectedCategories.add(cat);
    const label = document.createElement('label');
    label.className = 'checkbox-container';
    const input = document.createElement('input');
    input.type = 'checkbox'; input.checked = true; input.className = 'cat-chk-item';
    input.setAttribute('data-cat', cat);
    input.addEventListener('change', e => {
      if (e.target.checked) selectedCategories.add(cat);
      else selectedCategories.delete(cat);
      applyAndRender();
    });
    const span = document.createElement('span');
    span.className = 'checkmark';
    label.appendChild(input);
    label.appendChild(span);
    label.appendChild(document.createTextNode(cat));
    catContainer.appendChild(label);
  });
}

function rebuildDusunDropdown() {
  const dusunSel = document.getElementById('dusun-select');
  const prevDusun = dusunSel.value;
  dusunSel.innerHTML = '<option value="all">Semua Dusun</option>';
  getUniqueDusun().forEach(d => {
    const opt = document.createElement('option');
    opt.value = d; opt.textContent = `Dusun ${d}`;
    dusunSel.appendChild(opt);
  });
  if ([...dusunSel.options].some(o => o.value === prevDusun)) dusunSel.value = prevDusun;
  else { dusunSel.value = 'all'; selectedDusun = 'all'; }
}

// ─── UI Wiring ──────────────────────────────────────────────────────
function initUI() {
  // Theme Switcher Initialization & Event Listener
  const themeToggle = document.getElementById('theme-toggle');
  const savedTheme = localStorage.getItem('theme') || 'dark';
  if (savedTheme === 'light') {
    document.body.classList.add('light-theme');
    themeToggle.textContent = '🌙 Mode Gelap';
  } else {
    document.body.classList.remove('light-theme');
    themeToggle.textContent = '☀️ Mode Terang';
  }

  themeToggle.addEventListener('click', () => {
    const isLight = document.body.classList.toggle('light-theme');
    if (isLight) {
      localStorage.setItem('theme', 'light');
      themeToggle.textContent = '🌙 Mode Gelap';
    } else {
      localStorage.setItem('theme', 'dark');
      themeToggle.textContent = '☀️ Mode Terang';
    }
    // Re-render components that rely on js color tokens (e.g. charts)
    renderCharts();
  });

  document.getElementById('desa-select').addEventListener('change', e => {
    selectedDesa = e.target.value;
    rebuildDusunDropdown();
    applyAndRender();
  });
  document.getElementById('dusun-select').addEventListener('change', e => {
    selectedDusun = e.target.value;
    applyAndRender();
  });
  const searchInput = document.getElementById('search-input');
  const clearBtn = document.getElementById('clear-search');
  searchInput.addEventListener('input', e => {
    searchQuery = e.target.value;
    clearBtn.style.display = searchQuery ? 'block' : 'none';
    applyAndRender();
  });
  clearBtn.addEventListener('click', () => {
    searchInput.value = ''; searchQuery = '';
    clearBtn.style.display = 'none';
    applyAndRender();
  });
  document.getElementById('case-sensitive-chk').addEventListener('change', e => {
    isCaseSensitive = e.target.checked;
    applyAndRender();
  });

  // File upload (Sidebar)
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('xlsx-file');
  if (dropZone && fileInput) {
    dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', () => { dropZone.classList.remove('dragover'); });
    dropZone.addEventListener('drop', e => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
      if (e.dataTransfer.files.length) handleUpload(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener('change', e => { if (e.target.files.length) handleUpload(e.target.files[0]); });
  }

  // File upload (Main Empty State Card)
  const mainDropZone = document.getElementById('main-drop-zone');
  const mainFileInput = document.getElementById('main-xlsx-file');
  if (mainDropZone && mainFileInput) {
    mainDropZone.addEventListener('dragover', e => { e.preventDefault(); mainDropZone.classList.add('dragover'); });
    mainDropZone.addEventListener('dragleave', () => { mainDropZone.classList.remove('dragover'); });
    mainDropZone.addEventListener('drop', e => {
      e.preventDefault();
      mainDropZone.classList.remove('dragover');
      if (e.dataTransfer.files.length) handleUpload(e.dataTransfer.files[0]);
    });
    mainFileInput.addEventListener('change', e => { if (e.target.files.length) handleUpload(e.target.files[0]); });
  }

  // Modal
  const modal = document.getElementById('details-modal');
  document.querySelector('.close-modal').addEventListener('click', () => modal.classList.remove('active'));
  window.addEventListener('click', e => { if (e.target === modal) modal.classList.remove('active'); });

  // Sidebar toggle
  document.getElementById('toggle-sidebar').addEventListener('click', () => {
    document.querySelector('.dashboard-container').classList.toggle('sidebar-collapsed');
  });

  // Charts panel toggle
  const toggleChartsBtn = document.getElementById('toggle-charts-panel');
  const splitContainer = document.getElementById('needs-split-container');
  const closeChartsBtn = document.getElementById('close-charts-panel');

  function toggleChartsPanel() {
    const isVisible = splitContainer.classList.toggle('charts-visible');
    toggleChartsBtn.classList.toggle('active', isVisible);
    toggleChartsBtn.innerHTML = isVisible ? '📊 Sembunyikan' : '📊 Charts';
    if (isVisible) renderCharts();
  }

  toggleChartsBtn.addEventListener('click', toggleChartsPanel);
  closeChartsBtn.addEventListener('click', toggleChartsPanel);

  // CSV exports & Print
  document.getElementById('print-needs-summary').addEventListener('click', () => {
    window.print();
  });
  document.getElementById('download-needs-csv').addEventListener('click', exportNeedsSummaryCSV);
  const geoBtn = document.getElementById('download-geo-csv');
  if (geoBtn) geoBtn.addEventListener('click', exportGeoCSV);
}

// ─── FILTER & RENDER PIPELINE ───────────────────────────────────────
function applyAndRender() {
  const emptyState = document.getElementById('empty-state-view');
  const contentWrapper = document.getElementById('dashboard-content-wrapper');

  if (allData.length === 0) {
    if (emptyState) emptyState.style.display = 'flex';
    if (contentWrapper) contentWrapper.style.display = 'none';
    return;
  } else {
    if (emptyState) emptyState.style.display = 'none';
    if (contentWrapper) contentWrapper.style.display = 'flex';
  }

  filteredData = allData.filter(row => {
    if (selectedDesa !== 'all' && row.desa !== selectedDesa) return false;
    if (selectedDusun !== 'all' && row.dusun !== selectedDusun) return false;
    if (!row.kategori_rentan.some(c => selectedCategories.has(c))) return false;
    if (searchQuery.trim()) {
      const q = isCaseSensitive ? searchQuery.trim() : searchQuery.trim().toLowerCase();
      const match = row.all_needs.some(n => {
        const item = isCaseSensitive ? n : n.toLowerCase();
        return item.includes(q);
      });
      if (!match) return false;
    }
    return true;
  });
  currentPage = 1;
  renderKPIs();
  renderNeedsSummaryTable();
  const geoTbody = document.querySelector('#geo-breakdown-table tbody');
  if (geoTbody) renderGeoBreakdown();
  renderCharts();
  renderIndividualTable();
}

// ─── KPI CARDS (fully dynamic from data categories) ─────────────────
function renderKPIs() {
  const grid = document.getElementById('kpi-grid');
  grid.innerHTML = '';

  // Total card
  const totalCard = makeKPICard('📊', 'Total Terdata', filteredData.length, `${new Set(filteredData.map(r=>r.hh_id).filter(Boolean)).size} KK`);
  grid.appendChild(totalCard);

  // One card per unique category in filtered data
  const catCounts = {};
  filteredData.forEach(r => r.kategori_rentan.forEach(c => { catCounts[c] = (catCounts[c]||0)+1; }));
  const catIcons = {
    'Lanjut Usia (Lansia)': '🧓', 'Bayi (<12 Bulan)': '👶', 'Balita (1-5 Tahun)': '🧸',
    'Ibu Hamil': '🤰', 'Ibu Menyusui': '🤱', 'Disabilitas': '♿', 'Penyakit Kronis': '🩺', 'Umum': '👤'
  };
  Object.entries(catCounts).sort((a,b) => b[1]-a[1]).forEach(([cat, count]) => {
    grid.appendChild(makeKPICard(catIcons[cat]||'👥', cat, count, ''));
  });
}

function makeKPICard(icon, title, value, sub) {
  const div = document.createElement('div');
  div.className = 'kpi-card';
  div.innerHTML = `
    <div class="kpi-icon">${icon}</div>
    <div class="kpi-info">
      <h3>${title}</h3>
      <p class="kpi-value">${value}</p>
      ${sub ? `<span class="kpi-subtext">${sub}</span>` : ''}
    </div>`;
  return div;
}

// ─── PRIMARY TABLE: NEEDS × CATEGORY MATRIX ─────────────────────────
function renderNeedsSummaryTable() {
  // Gather all unique categories in filtered data
  const activeCats = [];
  const catSet = new Set();
  filteredData.forEach(r => r.kategori_rentan.forEach(c => catSet.add(c)));
  // Sort: put standard ones first
  const order = ['Lanjut Usia (Lansia)', 'Bayi (<12 Bulan)', 'Balita (1-5 Tahun)', 'Ibu Hamil', 'Ibu Menyusui', 'Disabilitas', 'Penyakit Kronis', 'Umum'];
  order.forEach(c => { if (catSet.has(c)) activeCats.push(c); });
  [...catSet].forEach(c => { if (!activeCats.includes(c)) activeCats.push(c); });

  // Gather all unique desa in filtered data
  const activeDesa = [...new Set(filteredData.map(r => r.desa))].sort();

  // Build aggregation: needItem -> { total, byCat: {cat: count}, byDesa: {desa: count} }
  const agg = {};
  filteredData.forEach(row => {
    row.all_needs.forEach(need => {
      if (searchQuery.trim()) {
        const q = isCaseSensitive ? searchQuery.trim() : searchQuery.trim().toLowerCase();
        const item = isCaseSensitive ? need : need.toLowerCase();
        if (!item.includes(q)) return;
      }
      if (!agg[need]) agg[need] = { total: 0, byCat: {}, byDesa: {} };
      agg[need].total++;
      row.kategori_rentan.forEach(cat => {
        agg[need].byCat[cat] = (agg[need].byCat[cat]||0)+1;
      });
      agg[need].byDesa[row.desa] = (agg[need].byDesa[row.desa]||0)+1;
    });
  });

  // Sort by total descending
  const sorted = Object.entries(agg).sort((a,b) => b[1].total - a[1].total);

  // Find max value in any Desa cell across the whole sorted list to normalize progress bars
  let maxDesaCount = 1;
  sorted.forEach(([_, data]) => {
    activeDesa.forEach(desa => {
      const v = data.byDesa[desa] || 0;
      if (v > maxDesaCount) maxDesaCount = v;
    });
  });

  // Build thead
  const thead = document.getElementById('needs-summary-thead');
  thead.innerHTML = '';
  const hRow = document.createElement('tr');
  hRow.innerHTML = '<th>Kebutuhan</th><th>Total (1d | 30d)</th>';
  activeCats.forEach(cat => {
    const th = document.createElement('th');
    th.textContent = cat;
    th.style.fontSize = '10px';
    hRow.appendChild(th);
  });
  activeDesa.forEach(desa => {
    const th = document.createElement('th');
    th.textContent = desa;
    th.style.fontSize = '10px';
    th.style.color = '#06b6d4';
    hRow.appendChild(th);
  });
  thead.appendChild(hRow);

  // Build tbody
  const tbody = document.getElementById('needs-summary-tbody');
  tbody.innerHTML = '';

  if (sorted.length === 0) {
    tbody.innerHTML = `<tr><td colspan="${2+activeCats.length+activeDesa.length}" style="text-align:center;color:var(--text-muted);">Tidak ada data yang cocok dengan filter.</td></tr>`;
    return;
  }

  sorted.forEach(([need, data]) => {
    const tr = document.createElement('tr');
    const totalProj = getProjection(need, data.total);
    const projLabel = getProjectionLabel(need);
    const projTitle = isOneTimeItem(need) ? 'Total (Item Sekali Pakai)' : 'Proyeksi Kumulatif 30 Hari';
    let cells = `
      <td class="text-highlight">${need}${isOneTimeItem(need) ? ' <small style="opacity:0.6;font-weight:400">⚡sekali</small>' : ''}</td>
      <td>
        <div class="total-cell">
          <span class="total-daily" title="Total Kebutuhan 1 Hari">${data.total}</span>
          <span class="total-proj" title="${projTitle}">${totalProj} <small>${projLabel}</small></span>
        </div>
      </td>`;
    
    activeCats.forEach(cat => {
      const v = data.byCat[cat] || 0;
      cells += `<td style="color:${v > 0 ? 'var(--text-white)' : 'var(--text-muted)'};font-weight:${v > 0 ? '600' : '400'}">${v || '-'}</td>`;
    });
    
    activeDesa.forEach(desa => {
      const v = data.byDesa[desa] || 0;
      if (v > 0) {
        const cumulative = getProjection(need, v);
        const percentage = (v / maxDesaCount) * 100;
        cells += `
          <td>
            <div class="projection-cell">
              <div class="projection-values">
                <span class="v-daily" title="Kebutuhan 1 Hari">${v}</span>
                <span class="v-proj" title="${isOneTimeItem(need) ? 'Item Sekali Pakai' : 'Proyeksi 30 Hari'}">${cumulative}<small>${isOneTimeItem(need) ? 'pcs' : 'pcs'}</small></span>
              </div>
              <div class="projection-bar" title="Total kebutuhan relatif">
                <div class="projection-fill" style="width: ${percentage}%"></div>
              </div>
            </div>
          </td>`;
      } else {
        cells += `<td class="empty-cell">-</td>`;
      }
    });
    
    tr.innerHTML = cells;
    tbody.appendChild(tr);
  });
}

// ─── GEOGRAPHIC BREAKDOWN TABLE ─────────────────────────────────────
function renderGeoBreakdown() {
  const tbody = document.querySelector('#geo-breakdown-table tbody');
  if (!tbody) return;
  tbody.innerHTML = '';

  const agg = {};
  filteredData.forEach(row => {
    row.kategori_rentan.forEach(cat => {
      row.all_needs.forEach(need => {
        if (searchQuery.trim()) {
          const q = isCaseSensitive ? searchQuery.trim() : searchQuery.trim().toLowerCase();
          if (!(isCaseSensitive ? need : need.toLowerCase()).includes(q)) return;
        }
        const key = `${row.desa}||${row.dusun}||${cat}||${need}`;
        if (!agg[key]) agg[key] = { desa: row.desa, dusun: row.dusun, cat, need, count: 0 };
        agg[key].count++;
      });
    });
  });

  const sorted = Object.values(agg).sort((a,b) => {
    if (a.desa !== b.desa) return a.desa.localeCompare(b.desa);
    if (a.dusun !== b.dusun) return String(a.dusun).localeCompare(String(b.dusun));
    if (a.cat !== b.cat) return a.cat.localeCompare(b.cat);
    return b.count - a.count;
  });

  if (!sorted.length) {
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:var(--text-muted);">Tidak ada data.</td></tr>';
    return;
  }

  sorted.forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="text-highlight">${r.desa}</td>
      <td>Dusun ${r.dusun}</td>
      <td><span class="badge ${getBadge(r.cat)}">${r.cat}</span></td>
      <td class="text-highlight">${r.need}</td>
      <td style="font-weight:800;color:var(--text-white)">${r.count}</td>`;
    tbody.appendChild(tr);
  });
}

// ─── CHARTS ─────────────────────────────────────────────────────────
function renderCharts() {
  // Needs bar chart (top 15) - mapped to projections
  const needsFreq = {};
  filteredData.forEach(row => {
    row.all_needs.forEach(n => {
      if (searchQuery.trim()) {
        const q = isCaseSensitive ? searchQuery.trim() : searchQuery.trim().toLowerCase();
        if (!(isCaseSensitive ? n : n.toLowerCase()).includes(q)) return;
      }
      needsFreq[n] = (needsFreq[n]||0)+1;
    });
  });
  const sortedN = Object.entries(needsFreq)
    .map(([n, count]) => [n, getProjection(n, count)])
    .sort((a,b) => b[1]-a[1])
    .slice(0, 15);

  const isLight = document.body.classList.contains('light-theme');
  const textColor = isLight ? '#0f172a' : '#f8fafc';
  const mutedColor = isLight ? '#475569' : '#94a3b8';
  const gridColor = isLight ? 'rgba(15, 23, 42, 0.06)' : 'rgba(255, 255, 255, 0.05)';

  if (needsChart) needsChart.destroy();
  needsChart = new Chart(document.getElementById('needs-chart').getContext('2d'), {
    type: 'bar',
    data: {
      labels: sortedN.map(i => i[0].length > 25 ? i[0].substring(0,23)+'…' : i[0]),
      datasets: [{ label: 'Proyeksi Kebutuhan (pcs)', data: sortedN.map(i => i[1]),
        backgroundColor: 'rgba(99,102,241,0.7)', borderColor: '#6366f1', borderWidth: 1, borderRadius: 6 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: { callbacks: { title: items => sortedN[items[0].dataIndex][0] } } },
      scales: {
        x: { grid: { color: gridColor }, ticks: { color: mutedColor, font: { size: 10 } } },
        y: { grid: { display: false }, ticks: { color: textColor, font: { family: 'Outfit', size: 11 } } }
      }
    }
  });

  // Cohort doughnut
  const cohortFreq = {};
  filteredData.forEach(r => r.kategori_rentan.forEach(c => { cohortFreq[c]=(cohortFreq[c]||0)+1; }));
  if (cohortChart) cohortChart.destroy();
  cohortChart = new Chart(document.getElementById('cohort-chart').getContext('2d'), {
    type: 'doughnut',
    data: {
      labels: Object.keys(cohortFreq),
      datasets: [{ data: Object.values(cohortFreq),
        backgroundColor: ['#6366f1','#10b981','#f59e0b','#06b6d4','#ef4444','#ec4899','#8b5cf6','#3b82f6'],
        borderWidth: 1, borderColor: isLight ? '#f8fafc' : 'rgba(10,14,26,0.9)' }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'right', labels: { color: mutedColor, font: { family: 'Outfit', size: 10 }, boxWidth: 12 } } }
    }
  });

  // Desa distribution bar chart - mapped to projections
  const desaFreq = {};
  filteredData.forEach(row => {
    row.all_needs.forEach(need => {
      if (searchQuery.trim()) {
        const q = isCaseSensitive ? searchQuery.trim() : searchQuery.trim().toLowerCase();
        if (!(isCaseSensitive ? need : need.toLowerCase()).includes(q)) return;
      }
      const proj = getProjection(need, 1);
      desaFreq[row.desa] = (desaFreq[row.desa] || 0) + proj;
    });
  });
  const sortedDesa = Object.entries(desaFreq).sort((a, b) => b[1] - a[1]);

  if (desaDistributionChart) desaDistributionChart.destroy();
  const desaCtx = document.getElementById('desa-distribution-chart');
  if (desaCtx) {
    const desaColors = ['#06b6d4', '#0891b2', '#0e7490', '#0369a1', '#0284c7', '#3b82f6', '#6366f1', '#8b5cf6'];
    desaDistributionChart = new Chart(desaCtx.getContext('2d'), {
      type: 'bar',
      data: {
        labels: sortedDesa.map(i => i[0]),
        datasets: [{
          label: 'Total Proyeksi Kebutuhan (pcs)',
          data: sortedDesa.map(i => i[1]),
          backgroundColor: sortedDesa.map((_, idx) => desaColors[idx % desaColors.length]),
          borderRadius: 6,
          borderWidth: 0
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { grid: { display: false }, ticks: { color: textColor, font: { family: 'Outfit', size: 10 } } },
          y: { grid: { color: gridColor }, ticks: { color: mutedColor, font: { size: 10 } } }
        }
      }
    });
  }

  // Specific needs chart (needs_specific field) - unfiltered, complete items, scrollable
  const specificFreq = {};
  allData.forEach(row => {
    (row.needs_specific || []).forEach(n => {
      if (n && n.trim()) {
        const key = n.trim();
        specificFreq[key] = (specificFreq[key] || 0) + 1;
      }
    });
  });
  const sortedSpecific = Object.entries(specificFreq)
    .map(([key, count]) => [key, getProjection(key, count)])
    .sort((a, b) => b[1] - a[1]);

  if (specificNeedsChart) specificNeedsChart.destroy();
  const container = document.getElementById('specific-needs-chart-container');
  const specCtx = document.getElementById('specific-needs-chart');

  if (specCtx) {
    if (sortedSpecific.length === 0) {
      if (container) container.style.height = '320px';
      specCtx.parentElement.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:var(--text-muted);font-style:italic">Tidak ada data kebutuhan khusus.</div>';
    } else {
      if (container) {
        container.style.height = `${Math.max(320, sortedSpecific.length * 28)}px`;
      }
      specificNeedsChart = new Chart(specCtx.getContext('2d'), {
        type: 'bar',
        data: {
          labels: sortedSpecific.map(i => i[0]),
          datasets: [{
            label: 'Proyeksi Kebutuhan (pcs)',
            data: sortedSpecific.map(i => i[1]),
            backgroundColor: sortedSpecific.map((_, idx) => {
              const colors = ['#10b981', '#059669', '#047857', '#0d9488', '#14b8a6', '#06b6d4', '#0891b2', '#0e7490', '#0284c7', '#0369a1',
                              '#6366f1', '#4f46e5', '#7c3aed', '#8b5cf6', '#a855f7', '#c084fc', '#ec4899', '#f43f5e', '#ef4444', '#f59e0b'];
              return colors[idx % colors.length];
            }),
            borderRadius: 6,
            borderWidth: 0
          }]
        },
        options: {
          responsive: true, maintainAspectRatio: false, indexAxis: 'y',
          plugins: {
            legend: { display: false },
            tooltip: { callbacks: { title: items => sortedSpecific[items[0].dataIndex][0] } }
          },
          scales: {
            x: { grid: { color: gridColor }, ticks: { color: mutedColor, font: { size: 10 } } },
            y: { 
              grid: { display: false }, 
              ticks: { 
                color: textColor, 
                autoSkip: false,
                font: { family: 'Outfit', size: 11, weight: '500' } 
              } 
            }
          }
        }
      });
    }
  }
}

// ─── INDIVIDUAL TABLE ───────────────────────────────────────────────
function renderIndividualTable() {
  const tbody = document.querySelector('#individual-table tbody');
  tbody.innerHTML = '';
  const total = filteredData.length;
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  if (currentPage > totalPages) currentPage = totalPages;
  const start = (currentPage-1)*pageSize;
  const end = Math.min(start+pageSize, total);
  document.getElementById('table-pagination-summary').textContent = total ? `${start+1}–${end} dari ${total}` : '0 records';

  if (!total) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--text-muted);">Tidak ada data.</td></tr>';
    document.getElementById('table-pagination-controls').innerHTML = '';
    return;
  }

  filteredData.slice(start, end).forEach(row => {
    const tr = document.createElement('tr');
    let ageText = row.umur !== null ? `${row.umur} th` : '-';
    if (row.umur_bulan !== null && row.umur_bulan < 12) ageText = `${row.umur_bulan} bl`;
    else if (row.umur_bulan !== null && row.umur_bulan <= 60) ageText = `${Math.floor(row.umur_bulan/12)} th ${row.umur_bulan%12} bl`;

    const catHTML = row.kategori_rentan.map(c => `<span class="badge ${getBadge(c)}">${c}</span>`).join(' ');
    const needsHTML = row.all_needs.map(n => {
      let cls = 'tag';
      if (searchQuery.trim()) {
        const q = isCaseSensitive ? searchQuery.trim() : searchQuery.trim().toLowerCase();
        if ((isCaseSensitive ? n : n.toLowerCase()).includes(q)) cls = 'tag tag-highlight';
      }
      return `<span class="${cls}">${n}</span>`;
    }).join(' ');

    const actionTd = document.createElement('td');
    const btn = document.createElement('button');
    btn.className = 'btn btn-primary'; btn.innerHTML = '👁️'; btn.style.padding = '4px 8px'; btn.style.fontSize = '11px';
    btn.addEventListener('click', () => showModal(row));
    actionTd.appendChild(btn);

    tr.innerHTML = `
      <td style="font-family:var(--font-display);font-weight:500">${row.hh_id||'N/A'}</td>
      <td class="text-highlight">${row.nama_rentan||'N/A'}</td>
      <td>${row.nama_kk||'N/A'}</td>
      <td>${ageText}</td>
      <td>${catHTML}</td>
      <td>${row.desa} / Dusun ${row.dusun}</td>
      <td><div class="tags-container">${needsHTML}</div></td>`;
    tr.appendChild(actionTd);
    tbody.appendChild(tr);
  });

  // Pagination
  const pag = document.getElementById('table-pagination-controls');
  pag.innerHTML = '';
  if (totalPages <= 1) return;
  const mkBtn = (label, page, disabled) => {
    const b = document.createElement('button'); b.className = `page-btn ${page===currentPage?'active':''}`;
    b.innerHTML = label; b.disabled = disabled;
    b.addEventListener('click', () => { currentPage = page; renderIndividualTable(); });
    return b;
  };
  pag.appendChild(mkBtn('«', currentPage-1, currentPage===1));
  for (let i = Math.max(1,currentPage-2); i <= Math.min(totalPages,currentPage+2); i++) pag.appendChild(mkBtn(i, i, false));
  pag.appendChild(mkBtn('»', currentPage+1, currentPage===totalPages));
}

// ─── DETAIL MODAL ───────────────────────────────────────────────────
function showModal(row) {
  document.getElementById('modal-rentan-name').textContent = row.nama_rentan;
  document.getElementById('modal-individual-id').textContent = `Row: ${row.row_idx} | Surveyor: ${row.surveyor}`;
  let ageText = row.umur !== null ? `${row.umur} Tahun` : '-';
  if (row.umur_bulan !== null) {
    if (row.umur_bulan < 12) ageText = `${row.umur_bulan} Bulan`;
    else ageText = `${Math.floor(row.umur_bulan/12)} Tahun ${row.umur_bulan%12} Bulan`;
  }
  document.getElementById('modal-age').textContent = ageText;
  document.getElementById('modal-gender').textContent = row.gender;
  document.getElementById('modal-kk-name').textContent = row.nama_kk;
  document.getElementById('modal-nik').textContent = row.nik_kk || '-';
  document.getElementById('modal-phone').textContent = row.phone || '-';
  document.getElementById('modal-kecamatan').textContent = row.kecamatan;
  document.getElementById('modal-desa').textContent = row.desa;
  document.getElementById('modal-dusun').textContent = `Dusun ${row.dusun}`;
  document.getElementById('modal-asal').textContent = row.asal_rt_rw || '-';
  document.getElementById('modal-categories').innerHTML = row.kategori_rentan.map(c => `<span class="badge ${getBadge(c)}">${c}</span>`).join(' ');
  document.getElementById('modal-detail-usia').textContent = row.detail_usia_penyakit || '-';
  document.getElementById('modal-all-needs').innerHTML = row.all_needs.length ? row.all_needs.map(n => `<span class="tag">${n}</span>`).join(' ') : '<span style="color:var(--text-muted)">Tidak ada</span>';
  const notesSec = document.getElementById('modal-notes-section');
  if (row.notes) { notesSec.style.display = 'block'; document.getElementById('modal-notes').textContent = row.notes; }
  else notesSec.style.display = 'none';
  document.getElementById('modal-surveyor').textContent = `Surveyor: ${row.surveyor}`;
  document.getElementById('modal-timestamp').textContent = `Timestamp: ${row.timestamp||'N/A'}`;
  document.getElementById('details-modal').classList.add('active');
}

// ─── CSV EXPORTS ────────────────────────────────────────────────────
function exportNeedsSummaryCSV() {
  const activeCats = [];
  const catSet = new Set();
  filteredData.forEach(r => r.kategori_rentan.forEach(c => catSet.add(c)));
  const order = ['Lanjut Usia (Lansia)', 'Bayi (<12 Bulan)', 'Balita (1-5 Tahun)', 'Ibu Hamil', 'Ibu Menyusui', 'Disabilitas', 'Penyakit Kronis', 'Umum'];
  order.forEach(c => { if (catSet.has(c)) activeCats.push(c); });
  [...catSet].forEach(c => { if (!activeCats.includes(c)) activeCats.push(c); });

  const activeDesa = [...new Set(filteredData.map(r => r.desa))].sort();

  const agg = {};
  filteredData.forEach(row => {
    row.all_needs.forEach(need => {
      if (searchQuery.trim()) {
        const q = isCaseSensitive ? searchQuery.trim() : searchQuery.trim().toLowerCase();
        const item = isCaseSensitive ? need : need.toLowerCase();
        if (!item.includes(q)) return;
      }
      if (!agg[need]) agg[need] = { total: 0, byCat: {}, byDesa: {} };
      agg[need].total++;
      row.kategori_rentan.forEach(c => { agg[need].byCat[c] = (agg[need].byCat[c]||0)+1; });
      agg[need].byDesa[row.desa] = (agg[need].byDesa[row.desa]||0)+1;
    });
  });

  let csv = '\ufeff'; // UTF-8 BOM
  csv += '"Kebutuhan","Total Harian","Total Proyeksi 30 Hari"';
  activeCats.forEach(c => { csv += `,"${c}"`; });
  activeDesa.forEach(d => { csv += `,"${d} (Harian)","${d} (30 Hari Proyeksi)"`; });
  csv += '\n';

  Object.entries(agg).sort((a,b) => b[1].total - a[1].total).forEach(([need, data]) => {
    const total30d = getProjection(need, data.total);
    csv += `"${need.replace(/"/g, '""')}",${data.total},${total30d}`;
    activeCats.forEach(c => {
      csv += `,${data.byCat[c] || 0}`;
    });
    activeDesa.forEach(d => {
      const v = data.byDesa[d] || 0;
      csv += `,${v},${getProjection(need, v)}`;
    });
    csv += '\n';
  });

  downloadCSV(csv, 'needs_summary.csv');
}

function exportGeoCSV() {
  const agg = {};
  filteredData.forEach(row => {
    row.kategori_rentan.forEach(cat => {
      row.all_needs.forEach(need => {
        if (searchQuery.trim()) {
          const q = isCaseSensitive ? searchQuery.trim() : searchQuery.trim().toLowerCase();
          const item = isCaseSensitive ? need : need.toLowerCase();
          if (!item.includes(q)) return;
        }
        const key = `${row.desa}||${row.dusun}||${cat}||${need}`;
        if (!agg[key]) agg[key] = { desa: row.desa, dusun: row.dusun, cat, need, count: 0 };
        agg[key].count++;
      });
    });
  });

  let csv = '\ufeff'; // UTF-8 BOM
  csv += '"Desa","Dusun","Kategori","Kebutuhan","Jumlah Harian","Jumlah Proyeksi 30 Hari"\n';
  Object.values(agg).sort((a,b) => {
    if (a.desa !== b.desa) return a.desa.localeCompare(b.desa);
    if (a.dusun !== b.dusun) return String(a.dusun).localeCompare(String(b.dusun));
    if (a.cat !== b.cat) return a.cat.localeCompare(b.cat);
    return b.count - a.count;
  }).forEach(r => {
    csv += `"${r.desa.replace(/"/g, '""')}","Dusun ${String(r.dusun).replace(/"/g, '""')}","${r.cat.replace(/"/g, '""')}","${r.need.replace(/"/g, '""')}",${r.count},${getProjection(r.need, r.count)}\n`;
  });
  downloadCSV(csv, 'geo_breakdown.csv');
}

function downloadCSV(csvContent, filename) {
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// ─── HELPERS ────────────────────────────────────────────────────────
function getBadge(cat) {
  if (cat.includes('Lansia')) return 'badge-warning';
  if (cat.includes('Bayi')) return 'badge-info';
  if (cat.includes('Balita')) return 'badge-primary';
  if (cat.includes('Hamil')) return 'badge-warning';
  if (cat.includes('Menyusui')) return 'badge-info';
  if (cat.includes('Disabilitas') || cat.includes('Kronis')) return 'badge-warning';
  return 'badge-primary';
}

// ─── CLIENT-SIDE XLSX UPLOAD & PARSE ────────────────────────────────
function handleUpload(file) {
  const status = document.getElementById('file-status');
  const mainStatus = document.getElementById('main-file-status');
  const processingText = '⚡ Memproses data…';
  
  if (status) {
    status.innerHTML = processingText; status.style.color = 'var(--color-warning)';
  }
  if (mainStatus) {
    mainStatus.innerHTML = processingText; mainStatus.style.color = 'var(--color-warning)';
  }

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const structured = {};
      let globalRowOffset = 1;
      let totalSheetsProcessed = 0;

      wb.SheetNames.forEach(sheetName => {
        const sheet = wb.Sheets[sheetName];
        const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
        if (rawRows.length < 2) return;
        totalSheetsProcessed++;

        for (let i = 1; i < rawRows.length; i++) {
          const rd = rawRows[i]; if (!rd) continue;
          const rowDict = {}; let has = false;
          for (let c = 0; c < rd.length; c++) {
            if (rd[c] !== null && rd[c] !== undefined && String(rd[c]).trim()) {
              rowDict[colLetter(c+1)] = rd[c]; has = true;
            }
          }
          if (has) {
            structured[globalRowOffset + i] = rowDict;
          }
        }
        globalRowOffset += rawRows.length;
      });

      if (totalSheetsProcessed === 0) throw new Error('No valid sheets with data found');

      // Reset filters to ensure all rows are read and displayed without exception
      selectedDesa = 'all';
      selectedDusun = 'all';
      selectedCategories.clear();
      searchQuery = '';

      const searchInput = document.getElementById('search-input');
      if (searchInput) searchInput.value = '';
      const clearBtn = document.getElementById('clear-search');
      if (clearBtn) clearBtn.style.display = 'none';

      const desaSel = document.getElementById('desa-select');
      if (desaSel) desaSel.value = 'all';
      const dusunSel = document.getElementById('dusun-select');
      if (dusunSel) dusunSel.value = 'all';

      allData = clientParse(structured);
      const successText = `✅ ${file.name} (${allData.length} records)`;
      if (status) {
        status.innerHTML = successText; status.style.color = 'var(--color-success)';
      }
      if (mainStatus) {
        mainStatus.innerHTML = successText; mainStatus.style.color = 'var(--color-success)';
      }
      rebuildDynamicFilters();
      applyAndRender();
    } catch (err) {
      console.error(err);
      const errorText = `❌ ${err.message}`;
      if (status) {
        status.innerHTML = errorText; status.style.color = 'var(--color-danger)';
      }
      if (mainStatus) {
        mainStatus.innerHTML = errorText; mainStatus.style.color = 'var(--color-danger)';
      }
    }
  };
  reader.readAsArrayBuffer(file);
}

function colLetter(n) {
  let s = '';
  while (n > 0) { const t = (n-1)%26; s = String.fromCharCode(65+t)+s; n = (n-t-1)/26; }
  return s;
}

// ─── CLIENT-SIDE HEURISTIC PARSER ───────────────────────────────────
const SURVEYOR_NAMES = ['Agelia Magi','7210031309120007','DEPRIN','Mirawati','HERLAMBANG P. PRATAMA','MARIA TALANTAN','Marta Tamolo','Melani','Aqila Ramadani'];

function clientParse(rows) {
  const records = [];
  for (const rIdx of Object.keys(rows).map(Number).sort((a,b)=>a-b)) {
    const row = rows[rIdx];
    const rec = { row_idx:rIdx, timestamp:null, dusun:'Lainnya', desa:'Lainnya', kecamatan:'Lainnya',
      village_id:null, hh_id:null, nama_kk:null, phone:null, asal_rt_rw:null, nik_kk:null,
      surveyor:'Tidak Diketahui', umur:null, gender:'Tidak Diketahui', kategori_rentan:[],
      detail_usia_penyakit:null, needs_dropdown:[], needs_specific:[], notes:null, nama_rentan:null, umur_bulan:null };

    const ua = {};
    for (const [c,v] of Object.entries(row)) { if (v!=null && String(v).trim()) ua[c]=String(v).trim(); }

    // Timestamp
    for (const c of Object.keys(ua)) { const s=ua[c]; if((/^\d+(\.\d+)?$/.test(s)&&+s>40000&&+s<50000)||(s.includes('/')&&s.includes(':'))) { rec.timestamp=s; delete ua[c]; break; } }
    // Desa/Kec
    for (const c of Object.keys(ua)) {
      const s = ua[c];
      const mapped = mapDesa(s);
      if (mapped) {
        rec.desa = mapped;
        delete ua[c];
      } else if (['Nokilalaki','Palolo'].includes(s)) {
        rec.kecamatan = s;
        delete ua[c];
      }
    }
    // Surveyor
    for (const c of Object.keys(ua)) { if(SURVEYOR_NAMES.includes(ua[c])){rec.surveyor=ua[c];delete ua[c];} }
    // Gender
    for (const c of Object.keys(ua)) { if(['Laki-laki','Perempuan'].includes(ua[c])){rec.gender=ua[c];delete ua[c];} }
    // KR IDs
    const kr=[]; for(const c of Object.keys(ua)){const s=ua[c];if(s.startsWith('KR-')&&/^KR-\d+$/.test(s))kr.push({c,s});}
    for(const{c,s}of kr){const n=+s.split('-')[1];if([1,60,90,200,417,581,661,754,839,864,976,986].includes(n))rec.village_id=s;else rec.hh_id=s;delete ua[c];}
    // NIK
    for(const c of Object.keys(ua)){const s=ua[c];if(/^\d+$/.test(s)&&s.length>=15){rec.nik_kk=s;delete ua[c];}}
    // Phone
    for(const c of Object.keys(ua)){const s=ua[c];if(/^08\d+$/.test(s)||(/^\d+$/.test(s)&&s.length>=9&&s.length<=13)||s.includes('E10')){rec.phone=s;delete ua[c];}}
    // Dusun (C)
    if(ua.C&&/^\d+(\.0)?$/.test(ua.C)){rec.dusun=String(parseInt(parseFloat(ua.C)));delete ua.C;}
    // Umur (L)
    if(ua.L&&/^\d+(\.0)?$/.test(ua.L)){rec.umur=parseFloat(ua.L);delete ua.L;}
    // Dropdown
    const dd=['Selimut, Sembako, Kelambu','Susu Formula, Popok, Obat obantan balita','Gangguan perilaku'];
    for(const c of Object.keys(ua)){if(dd.includes(ua[c])){rec.needs_dropdown=ua[c].split(',').map(s=>s.trim());delete ua[c];}}
    // Category
    const co=['Bayi / Balita','Lanjut Usia (Lansia)','Ibu Hamil','Ibu Menyusui','Disabilitas (Fisik / Sensorik / Mental)','Penyakit Kronis'];
    for(const c of Object.keys(ua)){const s=ua[c];if(co.some(o=>s===o||s.startsWith(o))){rec.kategori_rentan=s.split(',').map(x=>x.trim());delete ua[c];}}
    // Details vs needs
    const medWords=['bulan','tahun','hari','kandungan','menyusui','penyakit','stroke','diabetes','asma','hipertensi','jantung','gula','katarak','gatal','syaraf','disabilitas','tuna','kolestrol','urat'];
    const needWords=['susu','obat','popok','pempers','selimut','kelambu','sembako','terpal','kasur','bantuan','alat','mandi','telon','biskuit','kursi roda','tongkat'];
    for(const c of Object.keys(ua)){const s=ua[c],lo=s.toLowerCase();if(medWords.some(w=>lo.includes(w))){if(needWords.some(w=>lo.includes(w)))rec.needs_specific=s.split(/,|\bdan\b|&/).map(x=>x.trim()).filter(Boolean);else rec.detail_usia_penyakit=s;delete ua[c];}}
    for(const c of Object.keys(ua)){const s=ua[c],lo=s.toLowerCase();if(needWords.some(w=>lo.includes(w))){rec.needs_specific.push(...s.split(/,|\bdan\b|&/).map(x=>x.trim()).filter(Boolean));delete ua[c];}else if(s.length>25){rec.notes=s;delete ua[c];}else if(['rt','rw','dusun','asal','desa','dila','mando','bose'].some(w=>lo.includes(w))){rec.asal_rt_rw=s;delete ua[c];}}
    // Names
    for(const c of Object.keys(ua)){const cl=cleanN(ua[c]);if(cl){if(['G','H','I'].includes(c)){if(!rec.nama_kk){rec.nama_kk=cl;delete ua[c];}}else{if(!rec.nama_rentan){rec.nama_rentan=cl;delete ua[c];}}}}
    for(const c of Object.keys(ua)){const cl=cleanN(ua[c]);if(cl){if(!rec.nama_rentan)rec.nama_rentan=cl;else if(!rec.nama_kk)rec.nama_kk=cl;}}
    if(!rec.nama_kk)rec.nama_kk='Tidak Diketahui';
    if(!rec.nama_rentan)rec.nama_rentan=rec.nama_kk;

    // Age months
    rec.umur_bulan=parseMonths(rec.detail_usia_penyakit,rec.umur);

    // Category inference
    if(!rec.kategori_rentan.length){
      if(rec.umur!==null){if(rec.umur>=60)rec.kategori_rentan.push('Lanjut Usia (Lansia)');else if(rec.umur_bulan!==null){if(rec.umur_bulan<12)rec.kategori_rentan.push('Bayi');else if(rec.umur_bulan<=60)rec.kategori_rentan.push('Balita');}}
      const d=(rec.detail_usia_penyakit||'').toLowerCase();
      if(d.includes('hamil')||d.includes('kandungan'))rec.kategori_rentan.push('Ibu Hamil');
      if(d.includes('menyusui'))rec.kategori_rentan.push('Ibu Menyusui');
      if(d.includes('disabilitas'))rec.kategori_rentan.push('Disabilitas');
      if(['kronis','stroke','diabetes','asma'].some(w=>d.includes(w)))rec.kategori_rentan.push('Penyakit Kronis');
    }
    // Normalize category names
    const ku=[];
    for(const k of rec.kategori_rentan){const lo=k.toLowerCase().trim();
      if(lo.includes('lansia')||lo.includes('lanjut usia'))ku.push('Lanjut Usia (Lansia)');
      else if(lo.includes('hamil')||lo.includes('bumil'))ku.push('Ibu Hamil');
      else if(lo.includes('menyusui')||lo.includes('busui'))ku.push('Ibu Menyusui');
      else if(lo.includes('disabilitas'))ku.push('Disabilitas');
      else if(lo.includes('kronis')||lo.includes('penyakit')||['stroke','diabetes','asma','jantung','hipertensi','tensi','gula','paru','bronkitis','tumor','komplikasi','saraf'].some(i=>lo.includes(i)))ku.push('Penyakit Kronis');
      else if(lo.includes('bayi')||lo.includes('balita')){if(rec.umur_bulan!==null){ku.push(rec.umur_bulan<12?'Bayi (<12 Bulan)':'Balita (1-5 Tahun)');}else ku.push(rec.umur!==null&&rec.umur<1?'Bayi (<12 Bulan)':'Balita (1-5 Tahun)');}
      else ku.push(k);
    }
    if(rec.umur!==null){if(rec.umur>=60&&!ku.includes('Lanjut Usia (Lansia)'))ku.push('Lanjut Usia (Lansia)');else if(rec.umur_bulan!==null){if(rec.umur_bulan<12&&!ku.includes('Bayi (<12 Bulan)'))ku.push('Bayi (<12 Bulan)');else if(rec.umur_bulan<=60&&!ku.includes('Balita (1-5 Tahun)'))ku.push('Balita (1-5 Tahun)');}}
    rec.kategori_rentan=[...new Set(ku)];
    if(!rec.kategori_rentan.length)rec.kategori_rentan=['Umum'];

    // Build all_needs (raw, no normalization)
    const allN=[],seen=new Set();
    for(const n of[...rec.needs_dropdown,...rec.needs_specific]){const cl=n.trim();if(!cl)continue;const disp=cl[0].toUpperCase()+cl.slice(1);const key=disp.toLowerCase();if(!seen.has(key)){seen.add(key);allN.push(disp);}}
    rec.all_needs=allN;

    records.push(rec);
  }
  return records;
}

function cleanN(s){if(!s)return null;s=s.trim();if(/^\d+$/.test(s)||s.includes('KR-')||s.length>25||['layak huni','sakit','sehat','revisi','kk','ktp','alamat','tidak ada','belum ada'].some(w=>s.toLowerCase().includes(w)))return null;return s;}

function parseMonths(d,ay){
  if(!d){return ay!=null?Math.floor(ay*12):null;}
  const lo=String(d).toLowerCase();
  let m=lo.match(/(\d+)\s*tahun\s*(\d+)\s*bulan/);if(m)return+m[1]*12+ +m[2];
  m=lo.match(/(\d+)\s*bulan/);if(m){let y=0;const my=lo.match(/(\d+)\s*tahun/);if(my)y=+my[1];return y*12+ +m[1];}
  m=lo.match(/(\d+)\s*tahun/);if(m)return+m[1]*12;
  m=lo.match(/(\d+)\s*hari/);if(m)return Math.round(+m[1]/30*10)/10;
  return ay!=null?Math.floor(ay*12):null;
}

function mapDesa(valStr) {
  if (!valStr) return null;
  const valLower = String(valStr).toLowerCase().trim();
  if (valLower.includes('bulili')) return 'Bulili';
  if (valLower.includes('kadidia')) return 'Kadidia';
  if (valLower.includes('kamarora b') || valLower.includes('kamarora_b')) return 'Kamarora B';
  if (valLower.includes('kamarora a') || valLower.includes('kamarora_a') || valLower.includes('kamarora')) return 'Kamarora A';
  if (valLower.includes('lemban') || valLower.includes('tongoa')) return 'Lembantongoa';
  if (valLower.includes('sopu')) return 'Sopu';
  if (valLower.includes('uwenuni')) return 'Uwenuni';
  return null;
}
