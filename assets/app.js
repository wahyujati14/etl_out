/**
 * DataFlow ETL Studio — Outgoing Manufacturing Edition
 * Interactive Controller: Dashboard, Staging Anomaly Filter, Star Schema, & Guided SQL Exercises
 */

const state = {
  snapshot: null,
  transform: null,
  activeView: 'dashboard',
  stagingSheet: 'detail', // 'summary' | 'detail'
  stagingFilter: 'ALL',
  stagingSearch: '',
  warehouseTab: 'fact_outgoing_part',
  solvedExercises: new Set(),
};

const EXERCISES = {
  ex1_reconcile: {
    title: 'Latihan 1: Rekonsiliasi Sheet SUMMARY vs DETAIL',
    expectedRows: 6,
    sql: `-- LATIHAN 1: Temukan 6 Invoice yang jumlah baris DETAIL-nya lebih banyak dari SUMMARY (akibat Split PO)
SELECT 
    invoice_no,
    customer_code AS customer,
    summary_qty_part AS part_di_summary,
    raw_detail_rows AS baris_di_detail,
    split_extra_rows AS selisih_baris_split,
    pic_name AS pic
FROM dim_invoice
WHERE split_extra_rows > 0
ORDER BY split_extra_rows DESC;`
  },
  ex2_fanout: {
    title: 'Latihan 2: Buktikan Jebakan Double-Counting (61.300 pcs)',
    expectedRows: 26,
    sql: `-- LATIHAN 2: Tampilkan 26 Part yang terpecah (po_split_lines > 1) dan selisih overcount QTY_SCAN yang berhasil dicegah ETL
SELECT 
    invoice_no,
    customer_code AS customer,
    part_no,
    part_name_clean AS part_name,
    po_split_lines AS jml_baris_pecahan,
    qty_order_total AS total_order_asli,
    qty_scan_raw_sum AS scan_mentah_salah,
    qty_scan_clean AS scan_bersih_etl,
    overcount_saved AS pcs_terselamatkan
FROM fact_outgoing_part
WHERE po_split_lines > 1
ORDER BY overcount_saved DESC;`
  },
  ex3_lot_trace: {
    title: 'Latihan 3: Traceability Lot 1NF & Sub-Lot Bersarang (AISIN)',
    expectedRows: 10,
    sql: `-- LATIHAN 3: Cari Lot di tabel fact_outgoing_lot (hasil explode 1NF) yang memiliki Sub-Lot di dalam tanda kurung
SELECT 
    invoice_no,
    customer_code,
    part_no,
    part_name_clean,
    lot_full_code AS kode_lot_lengkap,
    primary_lot AS lot_utama,
    nested_sub_lot AS sub_lot_bersarang,
    lot_qty
FROM fact_outgoing_lot
WHERE nested_sub_lot IS NOT NULL AND nested_sub_lot != '-'
ORDER BY lot_qty DESC;`
  },
  ex4_window_rank: {
    title: 'Latihan 4: Window Function RANK() Part Terbesar per Customer',
    expectedRows: 18,
    sql: `-- LATIHAN 4: Gunakan CTE & Window Function RANK() untuk mencari Top 2 Part dengan Qty terbesar di tiap Customer
WITH ranked_parts AS (
    SELECT 
        customer_code,
        invoice_no,
        part_no,
        part_name_clean,
        qty_scan_clean,
        RANK() OVER (
            PARTITION BY customer_code 
            ORDER BY qty_scan_clean DESC
        ) AS peringkat
    FROM fact_outgoing_part
)
SELECT *
FROM ranked_parts
WHERE peringkat <= 2
ORDER BY customer_code ASC, peringkat ASC;`
  }
};

function formatNum(num) {
  return Number(num || 0).toLocaleString('id-ID');
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str === null || str === undefined ? '' : String(str);
  return div.innerHTML;
}

function showBanner(msg, isError = false) {
  const el = document.getElementById('statusBanner');
  const txt = document.getElementById('statusBannerText');
  if (!el || !txt) return;
  txt.textContent = msg;
  el.classList.toggle('is-error', isError);
  el.hidden = false;
}

function switchView(viewName) {
  state.activeView = viewName;
  document.querySelectorAll('.view-tab-btn').forEach((btn) => {
    btn.classList.toggle('is-active', btn.getAttribute('data-view') === viewName);
  });
  document.querySelectorAll('.view-section').forEach((sec) => {
    sec.hidden = sec.getAttribute('data-view-section') !== viewName;
  });
  if (viewName === 'dashboard' && state.snapshot) {
    setTimeout(() => renderCharts(state.snapshot.charts), 30);
  }
}

async function apiPost(action, formData = new FormData()) {
  formData.append('action', action);
  const res = await fetch(`api.php?action=${encodeURIComponent(action)}`, {
    method: 'POST',
    body: formData,
  });
  const data = await res.json();
  if (!data.ok) {
    throw new Error(data.error || 'Terjadi kesalahan pada server ETL.');
  }
  return data;
}

async function loadInitialStatus() {
  try {
    const res = await fetch('api.php?action=status');
    const data = await res.json();
    if (data.ok) {
      state.snapshot = data.snapshot;
      state.transform = data.transform;
      renderAll();
    }
  } catch (err) {
    showBanner(err.message, true);
  }
}

async function runExtractXlsx(isUpload = false) {
  try {
    const fd = new FormData();
    if (isUpload) {
      const input = document.getElementById('xlsxUploadInput');
      if (!input || !input.files.length) return;
      fd.append('xlsx_file', input.files[0]);
    }
    const data = await apiPost('extract', fd);
    state.snapshot = data.snapshot;
    state.transform = data.transform;
    showBanner(data.message);
    renderAll();
    switchView('pipeline');
  } catch (err) {
    showBanner(err.message, true);
  }
}

async function runEtlPipeline() {
  try {
    const fd = new FormData();
    ['dedup_split_po', 'explode_lots', 'clean_part_names', 'reconcile_status'].forEach((k) => {
      const cb = document.getElementById(`rule_${k}`);
      fd.append(k, cb && cb.checked ? '1' : '0');
    });
    const data = await apiPost('run_etl', fd);
    state.snapshot = data.snapshot;
    state.transform = data.transform;
    showBanner(data.message);
    renderAll();
  } catch (err) {
    showBanner(err.message, true);
  }
}

/* ==========================================================================
   Render UI Elements
   ========================================================================== */
function renderAll() {
  if (!state.snapshot) return;
  const snap = state.snapshot;
  const kpi = snap.kpi;

  document.getElementById('dbEngineLabel').textContent = snap.engine.label;

  // Navigation pills
  document.getElementById('navPillStaging').textContent = `${snap.counts.staging_summary}+${snap.counts.staging_detail}`;
  document.getElementById('navPillWarehouse').textContent = `${snap.counts.fact_part} Part / ${snap.counts.fact_lot} Lot`;

  // Reconciliation vs Box
  document.getElementById('vsRawScan').textContent = `${formatNum(kpi.total_qty_scan_raw)} pcs`;
  document.getElementById('vsCleanScan').textContent = `${formatNum(kpi.total_qty_scan_clean)} pcs`;
  document.getElementById('vsOvercountSaved').textContent = `-${formatNum(kpi.total_overcount_saved)} pcs`;

  // KPI Grid
  document.getElementById('kpiCleanQty').textContent = `${formatNum(kpi.total_qty_scan_clean)} pcs`;
  document.getElementById('kpiOvercount').textContent = `${formatNum(kpi.total_overcount_saved)} pcs`;
  document.getElementById('kpiUniqueParts').textContent = `${formatNum(snap.counts.fact_part)} Item Part`;
  document.getElementById('kpiExplodedLots').textContent = `${formatNum(snap.counts.fact_lot)} Baris Lot`;
  document.getElementById('kpiTotalBags').textContent = `${formatNum(kpi.total_bags)} Bag (${snap.counts.dim_invoice} Invoice)`;

  // Charts & Logs
  renderCharts(snap.charts);
  renderLogsTable(snap.logs);

  // Staging & Transform
  renderStagingArea();
  renderTransformDiffs();

  // Warehouse
  document.getElementById('whCountFactPart').textContent = `${snap.counts.fact_part} baris`;
  document.getElementById('whCountFactLot').textContent = `${snap.counts.fact_lot} baris`;
  document.getElementById('whCountDimInvoice').textContent = `${snap.counts.dim_invoice} baris`;
  document.getElementById('whCountDimPart').textContent = `${snap.counts.dim_part} baris`;
  renderWarehouseTable(state.warehouseTab);
}

function setStagingSheet(sheet) {
  state.stagingSheet = sheet;
  state.stagingFilter = 'ALL';
  document.getElementById('btnSheetDetail').classList.toggle('btn-primary', sheet === 'detail');
  document.getElementById('btnSheetSummary').classList.toggle('btn-primary', sheet === 'summary');
  renderStagingArea();
}

function setStagingFilter(filterCode) {
  state.stagingFilter = filterCode;
  document.querySelectorAll('[data-stg-filter]').forEach((btn) => {
    btn.classList.toggle('btn-primary', btn.getAttribute('data-stg-filter') === filterCode);
  });
  renderStagingArea();
}

function onStagingSearch(val) {
  state.stagingSearch = val.trim().toLowerCase();
  renderStagingArea();
}

function renderStagingArea() {
  if (!state.snapshot) return;
  const thead = document.getElementById('stagingTableHead');
  const tbody = document.getElementById('stagingTableBody');
  const caption = document.getElementById('stagingTableCaption');
  if (!thead || !tbody) return;

  if (state.stagingSheet === 'summary') {
    const rows = (state.snapshot.tables.staging_summary || []).filter((r) => {
      if (state.stagingSearch) {
        const hay = `${r.invoice_no} ${r.delivery_no} ${r.customer} ${r.pic}`.toLowerCase();
        if (!hay.includes(state.stagingSearch)) return false;
      }
      if (state.stagingFilter !== 'ALL') {
        return (r.anomalies || []).some((a) => a.code === state.stagingFilter || (state.stagingFilter === 'ANOMALY_ONLY' && r.anomalies.length > 0));
      }
      return true;
    });

    caption.textContent = `Sheet 1: SUMMARY (${rows.length} dari 20 Baris Ditampilkan)`;
    thead.innerHTML = `
      <tr>
        <th>No &amp; Invoice No</th>
        <th>Delivery No &amp; Tgl</th>
        <th>Customer</th>
        <th>Summary QTY PART</th>
        <th>Baris di Sheet DETAIL</th>
        <th>Jml Bag &amp; PIC</th>
        <th>Registered</th>
        <th>Deteksi Anomali</th>
      </tr>
    `;
    tbody.innerHTML = rows.map((r) => {
      const badges = (r.anomalies || []).length === 0
        ? `<span class="badge badge-success">Klop (100% Match)</span>`
        : r.anomalies.map((a) => `<span class="badge badge-${a.severity}">${escapeHtml(a.label)}</span>`).join(' ');
      return `
        <tr>
          <th scope="row">#${escapeHtml(r.row_num)} <code>${escapeHtml(r.invoice_no)}</code></th>
          <td><code>${escapeHtml(r.delivery_no)}</code><br><small>${escapeHtml(r.delivery_date)}</small></td>
          <td><strong>${escapeHtml(r.customer)}</strong></td>
          <td><strong>${escapeHtml(r.qty_part)}</strong> part</td>
          <td><strong>${escapeHtml(r.detail_row_count)}</strong> baris</td>
          <td>${escapeHtml(r.jml_bag)} bag • ${escapeHtml(r.pic)}</td>
          <td>${escapeHtml(r.registered_date)} ${escapeHtml(r.registered_time)}</td>
          <td>${badges}</td>
        </tr>
      `;
    }).join('');
    return;
  }

  // Sheet DETAIL
  const allDetail = state.snapshot.tables.staging_detail || [];
  const filtered = allDetail.filter((r) => {
    if (state.stagingSearch) {
      const hay = `${r.invoice_no} ${r.customer} ${r.part_no} ${r.part_name} ${r.lot_no_raw} ${r.pic}`.toLowerCase();
      if (!hay.includes(state.stagingSearch)) return false;
    }
    if (state.stagingFilter === 'ALL') return true;
    if (state.stagingFilter === 'ANOMALY_ONLY') return (r.anomalies || []).length > 0;
    return (r.anomalies || []).some((a) => a.code === state.stagingFilter);
  });

  caption.textContent = `Sheet 2: DETAIL (${filtered.length} dari ${allDetail.length} Baris Ditampilkan)`;
  thead.innerHTML = `
    <tr>
      <th>Baris &amp; Invoice</th>
      <th>Customer</th>
      <th>Part No &amp; Part Name Mentah</th>
      <th>QTY ORDER</th>
      <th>QTY SCAN (Raw)</th>
      <th>LOT NO (Raw 1 Sel)</th>
      <th>PIC &amp; Status</th>
      <th>Deteksi Anomali Kualitas Data</th>
    </tr>
  `;

  tbody.innerHTML = filtered.map((r) => {
    const badges = (r.anomalies || []).length === 0
      ? `<span class="badge badge-success">Normal</span>`
      : r.anomalies.map((a) => `<span class="badge badge-${a.severity}">${escapeHtml(a.label)}</span>`).join(' ');
    return `
      <tr>
        <th scope="row">#${escapeHtml(r.row_num)} <code>${escapeHtml(r.invoice_no)}</code></th>
        <td><strong>${escapeHtml(r.customer)}</strong></td>
        <td><code>${escapeHtml(r.part_no)}</code><br>"${escapeHtml(r.part_name)}"</td>
        <td><strong>${formatNum(r.qty_order)}</strong></td>
        <td><strong>${formatNum(r.qty_scan)}</strong></td>
        <td><small>${escapeHtml(r.lot_no_raw || '(KOSONG)')}</small></td>
        <td>${escapeHtml(r.pic)}<br><span class="badge ${r.status === 'OPEN' ? 'badge-critical' : 'badge-info'}">${escapeHtml(r.status)}</span></td>
        <td>${badges}</td>
      </tr>
    `;
  }).join('');
}

function renderTransformDiffs() {
  const tbody = document.getElementById('transformDiffBody');
  const summaryBadges = document.getElementById('transformMetricsBadges');
  if (!tbody || !state.transform) return;

  const { metrics, diffs } = state.transform;
  if (summaryBadges) {
    summaryBadges.innerHTML = `
      <span class="badge badge-critical">${metrics.split_rows_merged} Baris Split PO Digabung (-${formatNum(metrics.overcount_prevented)} pcs Overcount)</span>
      <span class="badge badge-info">${metrics.exploded_lot_rows} Lot Di-Explode ke 1NF</span>
      <span class="badge badge-warning">${metrics.part_names_cleaned} Part Name Dirapikan</span>
      <span class="badge badge-success">${metrics.status_reconciled} Status OPEN Direkonsiliasi</span>
    `;
  }

  tbody.innerHTML = diffs.map((d, idx) => {
    let badgeCls = 'badge-info';
    if (d.action === 'MERGED_SPLIT_PO') badgeCls = 'badge-critical';
    if (d.action === 'STATUS_FIXED') badgeCls = 'badge-warning';

    return `
      <tr>
        <th scope="row">${escapeHtml(d.row_nums)} • <code>${escapeHtml(d.part_no)}</code></th>
        <td><code>${escapeHtml(d.invoice_no)}</code> (${escapeHtml(d.customer)})</td>
        <td><span class="badge ${badgeCls}">${escapeHtml(d.action)}</span></td>
        <td>${escapeHtml(d.summary)}</td>
        <td>
          <button type="button" class="btn btn-sm" onclick="openDiffModal(${idx})">
            Inspeksi JSON Diff
          </button>
        </td>
      </tr>
    `;
  }).join('');
}

function renderWarehouseTable(tabName) {
  state.warehouseTab = tabName;
  document.querySelectorAll('[data-wh-tab]').forEach((btn) => {
    btn.classList.toggle('btn-primary', btn.getAttribute('data-wh-tab') === tabName);
  });

  if (!state.snapshot) return;
  const rows = state.snapshot.tables[tabName] || [];
  const thead = document.getElementById('warehouseTableHead');
  const tbody = document.getElementById('warehouseTableBody');
  const tableEl = document.getElementById('warehouseTable');
  const captionEl = document.getElementById('warehouseTableCaption');

  const titles = {
    fact_outgoing_part: 'fact_outgoing_part — 270 Baris Unik Invoice + Part (Bebas Double-Counting 61.300 pcs)',
    fact_outgoing_lot: 'fact_outgoing_lot — 345 Baris Traceability Lot Individual (Hasil Normalisasi 1NF)',
    dim_invoice: 'dim_invoice — 20 Master Invoice & Rekonsiliasi Summary vs Detail + Lead Time',
    dim_part: 'dim_part — Master Katalog Spring & Kategori Hasil Standarisasi Nama',
    dim_customer: 'dim_customer — Master Pelanggan Manufaktur & Total Bag',
  };
  if (captionEl) captionEl.textContent = titles[tabName] || tabName;

  if (rows.length === 0) {
    thead.innerHTML = `<tr><th>Info</th></tr>`;
    tbody.innerHTML = `<tr><th scope="row">Kosong</th><td>Tidak ada data di tabel ini.</td></tr>`;
    return;
  }

  const cols = Object.keys(rows[0]);
  cols.forEach((c, i) => tableEl.style.setProperty(`--label-${i + 1}`, `"${c}"`));

  thead.innerHTML = `<tr>${cols.map((c) => `<th>${escapeHtml(c)}</th>`).join('')}</tr>`;
  tbody.innerHTML = rows.map((r) => `
    <tr>${cols.map((c, idx) => {
      let val = r[c];
      if (['qty_order_total', 'qty_scan_raw_sum', 'qty_scan_clean', 'overcount_saved', 'lot_qty', 'jml_bag'].includes(c)) {
        val = formatNum(val);
      }
      if (c === 'overcount_saved' && Number(r[c]) > 0) {
        return `<td><span class="badge badge-critical">+${formatNum(r[c])} pcs dicegah</span></td>`;
      }
      return idx === 0
        ? `<th scope="row">${escapeHtml(val)}</th>`
        : `<td>${escapeHtml(val)}</td>`;
    }).join('')}</tr>
  `).join('');
}

/* ==========================================================================
   Canvas Analytics Charts
   ========================================================================== */
function renderCharts(charts) {
  if (!charts) return;
  drawHorizontalBarChart(
    'chartCustomerCanvas',
    (charts.by_customer || []).map((c) => ({
      label: `${c.customer_code} (${c.unique_parts} part)`,
      value: Number(c.clean_qty || 0),
      extra: Number(c.overcount_qty || 0) > 0 ? ` [Overcount dicegah: ${formatNum(c.overcount_qty)}]` : '',
    })),
    '#2563eb'
  );

  drawHorizontalBarChart(
    'chartCategoryCanvas',
    (charts.by_category || []).map((c) => ({
      label: `${c.part_category} (${c.part_lines})`,
      value: Number(c.total_qty || 0),
      extra: '',
    })),
    '#0d9488'
  );
}

function drawHorizontalBarChart(canvasId, items, barColor) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const dpr = window.devicePixelRatio || 1;
  const cssWidth = canvas.clientWidth || 500;
  const cssHeight = 240;
  canvas.width = cssWidth * dpr;
  canvas.height = cssHeight * dpr;
  ctx.scale(dpr, dpr);

  const isDark = document.documentElement.getAttribute('data-theme') === 'dark' ||
    (!document.documentElement.getAttribute('data-theme') && window.matchMedia('(prefers-color-scheme: dark)').matches);

  const textColor = isDark ? '#f8fafc' : '#0f172a';
  const mutedColor = isDark ? '#94a3b8' : '#475569';
  const trackColor = isDark ? '#1e293b' : '#e2e8f0';

  ctx.clearRect(0, 0, cssWidth, cssHeight);
  if (!items || items.length === 0) return;

  const maxVal = Math.max(...items.map((i) => i.value), 1);
  const padLeft = 165;
  const padRight = 155;
  const usableWidth = Math.max(80, cssWidth - padLeft - padRight);
  const rowHeight = Math.min(30, Math.floor((cssHeight - 14) / items.length));

  items.forEach((item, idx) => {
    const y = 10 + idx * rowHeight;
    const barH = Math.max(13, rowHeight - 10);
    const barW = Math.round((item.value / maxVal) * usableWidth);

    ctx.fillStyle = textColor;
    ctx.font = '600 11.5px Inter, sans-serif';
    ctx.fillText(item.label.slice(0, 24), 6, y + barH - 2);

    ctx.fillStyle = trackColor;
    ctx.beginPath();
    ctx.roundRect(padLeft, y, usableWidth, barH, 5);
    ctx.fill();

    ctx.fillStyle = barColor;
    ctx.beginPath();
    ctx.roundRect(padLeft, y, Math.max(6, barW), barH, 5);
    ctx.fill();

    ctx.fillStyle = mutedColor;
    ctx.font = '600 11px Inter, sans-serif';
    ctx.fillText(`${formatNum(item.value)} pcs`, padLeft + usableWidth + 8, y + barH - 2);
  });
}

function renderLogsTable(logs) {
  const tbody = document.getElementById('logsTableBody');
  if (!tbody) return;
  tbody.innerHTML = (logs || []).map((l) => `
    <tr>
      <th scope="row"><code>${escapeHtml(l.batch_id)}</code></th>
      <td>${escapeHtml(l.file_name)}</td>
      <td>${escapeHtml(l.summary_rows)} Summary / ${escapeHtml(l.detail_raw_rows)} Detail</td>
      <td><strong>${escapeHtml(l.fact_part_rows)} Part</strong> / <strong>${escapeHtml(l.fact_lot_rows)} Lot</strong></td>
      <td><span class="badge badge-critical">-${formatNum(l.overcount_prevented)} pcs Overcount</span></td>
      <td>${escapeHtml(l.duration_ms)} ms (${escapeHtml(l.executed_at)})</td>
    </tr>
  `).join('');
}

/* ==========================================================================
   Guided Exercises & SQL Playground
   ========================================================================== */
async function runExercise(exKey) {
  const ex = EXERCISES[exKey];
  if (!ex) return;
  const editor = document.getElementById('sqlQueryInput');
  if (editor) {
    editor.value = ex.sql;
  }
  await executeSqlLab(exKey);
}

async function executeSqlLab(activeExKey = null) {
  const editor = document.getElementById('sqlQueryInput');
  const thead = document.getElementById('sqlResultHead');
  const tbody = document.getElementById('sqlResultBody');
  const meta = document.getElementById('sqlResultMeta');
  if (!editor || !thead || !tbody) return;

  try {
    const fd = new FormData();
    fd.append('sql', editor.value);
    const data = await apiPost('sql_query', fd);

    if (meta) {
      meta.textContent = `${data.row_count} baris dikembalikan dalam ${data.elapsed_ms} ms${data.engine_note || ''}`;
    }

    if (activeExKey && EXERCISES[activeExKey]) {
      const ex = EXERCISES[activeExKey];
      if (data.row_count === ex.expectedRows) {
        state.solvedExercises.add(activeExKey);
        const card = document.getElementById(`card_${activeExKey}`);
        const badge = document.getElementById(`badge_${activeExKey}`);
        if (card) card.classList.add('is-solved');
        if (badge) {
          badge.className = 'badge badge-success';
          badge.textContent = `✓ Selesai (${data.row_count} baris tepat)`;
        }
        document.getElementById('exerciseProgressBadge').textContent = `${state.solvedExercises.size} / 4 Latihan Selesai`;
        showBanner(`Mantap! ${ex.title} berhasil diselesaikan (${data.row_count} baris terverifikasi).`);
      }
    }

    const cols = data.columns || [];
    thead.innerHTML = `<tr>${cols.map((c) => `<th>${escapeHtml(c)}</th>`).join('')}</tr>`;
    tbody.innerHTML = (data.rows || []).map((r) => `
      <tr>${cols.map((c, idx) => {
        const v = r[c];
        return idx === 0 ? `<th scope="row">${escapeHtml(v)}</th>` : `<td>${escapeHtml(v)}</td>`;
      }).join('')}</tr>
    `).join('');
  } catch (err) {
    showBanner(err.message, true);
  }
}

/* ==========================================================================
   Modal Inspector & Light-Dismiss Fallback
   ========================================================================== */
function openDiffModal(idx) {
  if (!state.transform || !state.transform.diffs[idx]) return;
  const d = state.transform.diffs[idx];
  const dialog = document.getElementById('rowInspectorDialog');
  if (!dialog) return;

  document.getElementById('dialogRowTitle').textContent = `Inspeksi Transformasi ${d.row_nums} — Part ${d.part_no} (${d.customer})`;
  document.getElementById('dialogRowSummary').textContent = d.summary;
  document.getElementById('dialogBeforeJson').textContent = JSON.stringify(d.before, null, 2);
  document.getElementById('dialogAfterJson').textContent = JSON.stringify(d.after, null, 2);
  dialog.showModal();
}

function setupDialogLightDismissFallbacks() {
  document.querySelectorAll('dialog').forEach((dialog) => {
    if (!('closedBy' in HTMLDialogElement.prototype)) {
      dialog.addEventListener('click', (event) => {
        if (event.target !== dialog) return;
        const rect = dialog.getBoundingClientRect();
        const isDialogContent = (
          rect.top <= event.clientY &&
          event.clientY <= rect.top + rect.height &&
          rect.left <= event.clientX &&
          event.clientX <= rect.left + rect.width
        );
        if (isDialogContent) return;
        dialog.close();
      });
    }
  });
}

function toggleColorTheme() {
  const html = document.documentElement;
  const current = html.getAttribute('data-theme');
  const sysDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
  const next = current ? (current === 'dark' ? 'light' : 'dark') : (sysDark ? 'light' : 'dark');
  html.setAttribute('data-theme', next);
  if (state.snapshot) renderCharts(state.snapshot.charts);
}

window.addEventListener('DOMContentLoaded', async () => {
  setupDialogLightDismissFallbacks();
  window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
    if (state.snapshot) renderCharts(state.snapshot.charts);
  });
  window.addEventListener('resize', () => {
    if (state.snapshot && state.activeView === 'dashboard') {
      renderCharts(state.snapshot.charts);
    }
  });

  await loadInitialStatus();
  const editor = document.getElementById('sqlQueryInput');
  if (editor) {
    editor.value = EXERCISES.ex1_reconcile.sql;
    await executeSqlLab();
  }
});
