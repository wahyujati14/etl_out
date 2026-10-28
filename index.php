<?php
declare(strict_types=1);
?>
<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <!-- MANDATORY: Declare support for both light and dark themes to help prevent FOUC -->
  <meta name="color-scheme" content="light dark">
  <title>DataFlow ETL Studio — Outgoing Delivery &amp; Latihan Data Engineer</title>
  <link rel="stylesheet" href="assets/style.css">
</head>
<body>

  <!-- Sticky Header & View Navigation Bar -->
  <header class="app-header">
    <div class="header-top">
      <div class="brand-group">
        <div class="brand-logo" aria-hidden="true">ETL</div>
        <div class="brand-title">
          <h1>DataFlow ETL Studio • Outgoing Edition</h1>
          <p>Pipeline &amp; Latihan Data Engineer: <code>OUTGOING_2026-10-28_sd_2026-10-28.xlsx</code></p>
        </div>
      </div>

      <div class="header-actions">
        <div class="db-badge" title="Engine Database Aktif">
          <span class="db-dot" aria-hidden="true"></span>
          <span id="dbEngineLabel">Memuat Engine...</span>
        </div>

        <label class="btn btn-sm" style="cursor:pointer;">
          Upload File .XLSX Baru
          <input type="file" id="xlsxUploadInput" accept=".xlsx" style="display:none" onchange="runExtractXlsx(true)">
        </label>

        <button type="button" class="btn btn-sm btn-primary" onclick="runEtlPipeline()">
          Jalankan Ulang Pipeline ETL
        </button>

        <button type="button" class="btn btn-sm" onclick="toggleColorTheme()">
          Tema Terang / Gelap
        </button>
      </div>
    </div>

    <!-- Clean View Switcher Tabs -->
    <nav class="view-nav-bar" aria-label="Menu Tampilan Utama">
      <div class="view-nav-inner">
        <button type="button" class="view-tab-btn is-active" data-view="dashboard" onclick="switchView('dashboard')">
          1. Dashboard &amp; Rekonsiliasi
        </button>
        <button type="button" class="view-tab-btn" data-view="pipeline" onclick="switchView('pipeline')">
          2. Staging &amp; Transformasi ETL
          <span class="view-tab-pill" id="navPillStaging">20+298</span>
        </button>
        <button type="button" class="view-tab-btn" data-view="warehouse" onclick="switchView('warehouse')">
          3. Star Schema &amp; Lot 1NF
          <span class="view-tab-pill" id="navPillWarehouse">270 Part / 345 Lot</span>
        </button>
        <button type="button" class="view-tab-btn" data-view="exercises" onclick="switchView('exercises')">
          4. Latihan Praktik &amp; SQL Lab
          <span class="view-tab-pill" id="exerciseProgressBadge">0 / 4 Latihan Selesai</span>
        </button>
      </div>
    </nav>
  </header>

  <main class="app-container">

    <!-- Status Notification Banner -->
    <div id="statusBanner" class="status-banner" role="status" aria-live="polite" hidden>
      <span id="statusBannerText">Siap.</span>
      <button type="button" class="btn btn-sm" onclick="document.getElementById('statusBanner').hidden = true">Tutup</button>
    </div>

    <!-- ======================================================================
         VIEW 1: DASHBOARD & REKONSILIASI OUTGOING
         ====================================================================== -->
    <div class="view-section" data-view-section="dashboard">

      <!-- Reconciliation Alert Banner -->
      <section class="recon-banner" aria-labelledby="reconHeading">
        <div class="recon-info">
          <h2 id="reconHeading">Rekonsiliasi Otomatis: Mengapa Data Excel Ini Tidak Boleh Di-SUM Mentah?</h2>
          <p>
            Di sheet <code>DETAIL</code> (298 baris), terdapat <strong>28 baris pecahan PO (Split Order)</strong> milik <strong>EPSON</strong> &amp; <strong>MURAMOTO</strong>
            yang mengulang angka total <code>QTY SCAN</code> di setiap baris pecahan. Pipeline ETL otomatis menggabungkan baris pecahan tersebut menjadi
            <strong>270 Part Unik</strong> (100% klop dengan sheet <code>SUMMARY</code>) dan mencegah kelebihan hitung sebesar <strong>61.300 pcs</strong>!
          </p>
        </div>

        <div class="recon-vs-box">
          <div class="vs-card is-danger">
            <small>Raw SUM(QTY_SCAN) Salah</small>
            <strong id="vsRawScan">1.353.559 pcs</strong>
            <span style="font-size:0.72rem;">Terjebak Double-Counting</span>
          </div>
          <div class="vs-card is-success">
            <small>Clean ETL Qty (Klop)</small>
            <strong id="vsCleanScan">1.292.259 pcs</strong>
            <span style="font-size:0.72rem;">= SUM(QTY_ORDER)</span>
          </div>
          <div class="vs-card">
            <small>Overcount Dicegah</small>
            <strong id="vsOvercountSaved">-61.300 pcs</strong>
            <span style="font-size:0.72rem;">Dari 28 Baris Split PO</span>
          </div>
        </div>
      </section>

      <!-- Executive KPI Grid -->
      <section class="panel" aria-labelledby="kpiTitle">
        <div class="panel-header">
          <div>
            <h3 id="kpiTitle">Ringkasan Eksekutif Pengiriman Outgoing (Periode 2026-10-28)</h3>
            <p>Metrik bersih dari Star Schema Data Warehouse setelah deduplikasi Split PO dan normalisasi Lot 1NF.</p>
          </div>
          <div style="display:flex; gap:0.5rem;">
            <button type="button" class="btn btn-sm btn-primary" onclick="switchView('exercises')">
              Mulai Latihan SQL Data Engineer →
            </button>
          </div>
        </div>

        <div class="kpi-grid">
          <div class="kpi-card">
            <div class="kpi-label">Total Qty Terkirim (Valid)</div>
            <div class="kpi-value" id="kpiCleanQty">0 pcs</div>
            <div class="kpi-sub">100% klop antara Order, Scan &amp; Lot</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-label">Selisih Fan-Out Terselamatkan</div>
            <div class="kpi-value" id="kpiOvercount" style="color:var(--status-danger)">0 pcs</div>
            <div class="kpi-sub">Dieliminasi dari 28 baris duplikat scan</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-label">Total Part Unik (Fact Part)</div>
            <div class="kpi-value" id="kpiUniqueParts">0 Item</div>
            <div class="kpi-sub">Dari 298 baris Detail → 270 Part Summary</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-label">Total Lot Terurai (1NF)</div>
            <div class="kpi-value" id="kpiExplodedLots">0 Lot</div>
            <div class="kpi-sub">Dipecah dari string multi-lot di kolom I</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-label">Total Kemasan &amp; Invoice</div>
            <div class="kpi-value" id="kpiTotalBags">0 Bag</div>
            <div class="kpi-sub">8 Customer Manufaktur (Status CLOSED)</div>
          </div>
        </div>

        <div class="charts-grid">
          <div class="chart-card">
            <h4>Volume Pengiriman Bersih per Customer (<code>fact_outgoing_part</code>)</h4>
            <canvas id="chartCustomerCanvas" width="520" height="240" aria-label="Grafik volume pengiriman per customer"></canvas>
          </div>
          <div class="chart-card">
            <h4>Volume Pengiriman per Kategori Produk Spring (<code>dim_part</code>)</h4>
            <canvas id="chartCategoryCanvas" width="520" height="240" aria-label="Grafik volume pengiriman per kategori spring"></canvas>
          </div>
        </div>
      </section>

      <!-- Audit Log History -->
      <section class="panel heavy-section-deferred" aria-labelledby="logTitle">
        <div class="panel-header">
          <div>
            <h3 id="logTitle">Riwayat Eksekusi Pipeline ETL (<code>etl_outgoing_logs</code>)</h3>
            <p>Catatan observability setiap proses Extract, Transform, dan Load yang dijalankan pada file Outgoing.</p>
          </div>
        </div>
        <div tabindex="0" role="region" aria-labelledby="logsCaption" class="table-wrapper">
          <table style="--label-2:'File Sumber'; --label-3:'Baris Mentah'; --label-4:'Baris Star Schema'; --label-5:'Overcount Dicegah'; --label-6:'Durasi';">
            <caption id="logsCaption">Audit Log Eksekusi Pipeline Outgoing</caption>
            <thead>
              <tr>
                <th>Batch ID</th>
                <th>File Sumber</th>
                <th>Baris Mentah (Staging)</th>
                <th>Hasil Star Schema</th>
                <th>Overcount Dicegah</th>
                <th>Durasi &amp; Waktu</th>
              </tr>
            </thead>
            <tbody id="logsTableBody"></tbody>
          </table>
        </div>
      </section>

    </div>

    <!-- ======================================================================
         VIEW 2: PIPELINE ETL INTERAKTIF (STAGING & TRANSFORMASI)
         ====================================================================== -->
    <div class="view-section" data-view-section="pipeline" hidden>

      <section class="panel" aria-labelledby="stgTitle">
        <div class="panel-header">
          <div>
            <h3 id="stgTitle">Tahap 1: EXTRACT — Eksplorasi Staging Area &amp; Filter Anomali</h3>
            <p>Pilih sheet dan filter anomali di bawah ini untuk melihat langsung baris-baris bermasalah di file Excel asli.</p>
          </div>
          <div style="display:flex; gap:0.45rem;">
            <button type="button" id="btnSheetDetail" class="btn btn-sm btn-primary" onclick="setStagingSheet('detail')">
              Sheet 2: DETAIL (298 Baris)
            </button>
            <button type="button" id="btnSheetSummary" class="btn btn-sm" onclick="setStagingSheet('summary')">
              Sheet 1: SUMMARY (20 Baris)
            </button>
          </div>
        </div>

        <div class="filter-bar">
          <div class="filter-pills" role="group" aria-label="Filter Anomali Staging">
            <button type="button" class="btn btn-sm btn-primary" data-stg-filter="ALL" onclick="setStagingFilter('ALL')">
              Semua Baris
            </button>
            <button type="button" class="btn btn-sm" data-stg-filter="ANOMALY_ONLY" onclick="setStagingFilter('ANOMALY_ONLY')">
              Hanya Baris Beranomali
            </button>
            <button type="button" class="btn btn-sm" data-stg-filter="SPLIT_PO_FANOUT" onclick="setStagingFilter('SPLIT_PO_FANOUT')">
              ⚠️ Split PO (Double-Count Trap)
            </button>
            <button type="button" class="btn btn-sm" data-stg-filter="STATUS_OPEN_MISMATCH" onclick="setStagingFilter('STATUS_OPEN_MISMATCH')">
              ⚠️ Status OPEN (Baris #56)
            </button>
            <button type="button" class="btn btn-sm" data-stg-filter="NESTED_LOT_FORMAT" onclick="setStagingFilter('NESTED_LOT_FORMAT')">
              🔍 Lot Bersarang (AISIN)
            </button>
            <button type="button" class="btn btn-sm" data-stg-filter="MESSY_PART_NAME" onclick="setStagingFilter('MESSY_PART_NAME')">
              🧹 Nama Part Kotor
            </button>
          </div>

          <input
            type="search"
            class="search-input"
            placeholder="Cari Invoice, Customer, Part No, Lot..."
            aria-label="Cari di tabel staging"
            oninput="onStagingSearch(this.value)"
          >
        </div>

        <div tabindex="0" role="region" aria-labelledby="stagingTableCaption" class="table-wrapper">
          <table style="--label-2:'Customer'; --label-3:'Part'; --label-4:'QTY ORDER'; --label-5:'QTY SCAN'; --label-6:'LOT NO'; --label-7:'PIC & Status'; --label-8:'Anomali';">
            <caption id="stagingTableCaption">Sheet 2: DETAIL</caption>
            <thead id="stagingTableHead"></thead>
            <tbody id="stagingTableBody"></tbody>
          </table>
        </div>
      </section>

      <section class="panel" aria-labelledby="trfTitle">
        <div class="panel-header">
          <div>
            <h3 id="trfTitle">Tahap 2: TRANSFORM — Konfigurasi Aturan Pembersihan &amp; Log Perubahan</h3>
            <p>Coba matikan centang <em>"Gabungkan Split PO"</em> lalu klik <strong>Terapkan Transformasi</strong> untuk melihat apa yang terjadi jika tidak di-deduplikasi!</p>
          </div>
          <button type="button" class="btn btn-primary" onclick="runEtlPipeline()">
            Terapkan Transformasi &amp; Load ke Warehouse
          </button>
        </div>

        <div class="rules-grid">
          <label class="rule-item">
            <input type="checkbox" id="rule_dedup_split_po" checked>
            <div>
              <strong>Gabungkan Split PO per (Invoice + Part No)</strong>
              <span>Jumlahkan <code>QTY ORDER</code> yang terpecah dan cegah <code>QTY SCAN</code> terhitung ganda (hemat 61.300 pcs).</span>
            </div>
          </label>
          <label class="rule-item">
            <input type="checkbox" id="rule_explode_lots" checked>
            <div>
              <strong>Normalisasi 1NF Kolom LOT NO (Explode)</strong>
              <span>Uraikan string multi-lot &amp; sub-lot bersarang menjadi 345 baris relasional di <code>fact_outgoing_lot</code>.</span>
            </div>
          </label>
          <label class="rule-item">
            <input type="checkbox" id="rule_clean_part_names" checked>
            <div>
              <strong>Standarisasi Nama Part &amp; Klasifikasi Kategori</strong>
              <span>Perbaiki <code>"Compressionspring  7.77"</code>, <code>"SPRING(FG)"</code>, dan hapus spasi ekor berlebih.</span>
            </div>
          </label>
          <label class="rule-item">
            <input type="checkbox" id="rule_reconcile_status" checked>
            <div>
              <strong>Rekonsiliasi Status OPEN vs SUMMARY</strong>
              <span>Selaraskan baris #56 yang masih berstatus <code>OPEN</code> agar mengikuti status <code>CLOSED</code> di Summary.</span>
            </div>
          </label>
        </div>

        <div style="display:flex; flex-wrap:wrap; align-items:center; justify-content:space-between; gap:0.6rem;">
          <strong>Daftar Baris yang Mengalami Transformasi:</strong>
          <div id="transformMetricsBadges" style="display:flex; flex-wrap:wrap; gap:0.4rem;"></div>
        </div>

        <div tabindex="0" role="region" aria-labelledby="trfDiffCaption" class="table-wrapper">
          <table style="--label-2:'Invoice & Customer'; --label-3:'Tipe Aksi'; --label-4:'Ringkasan Perubahan'; --label-5:'Inspeksi';">
            <caption id="trfDiffCaption">Log Transformasi Baris (Klik "Inspeksi JSON Diff" untuk Melihat Detail Sebelum vs Sesudah)</caption>
            <thead>
              <tr>
                <th>Baris Asal &amp; Part No</th>
                <th>Invoice &amp; Customer</th>
                <th>Tipe Aksi</th>
                <th>Ringkasan Perubahan (Before → After)</th>
                <th>Detail JSON</th>
              </tr>
            </thead>
            <tbody id="transformDiffBody"></tbody>
          </table>
        </div>
      </section>

    </div>

    <!-- ======================================================================
         VIEW 3: STAR SCHEMA & LOT 1NF WAREHOUSE EXPLORER
         ====================================================================== -->
    <div class="view-section" data-view-section="warehouse" hidden>

      <section class="panel" aria-labelledby="whTitle">
        <div class="panel-header">
          <div>
            <h3 id="whTitle">Tahap 3: LOAD — Star Schema Data Warehouse &amp; Tabel Traceability Lot</h3>
            <p>Dari 2 sheet Excel yang tidak ternormalisasi, pipeline membentuk 2 Tabel Fakta dan 3 Tabel Dimensi yang bersih.</p>
          </div>
          <div style="display:flex; flex-wrap:wrap; gap:0.4rem;">
            <button type="button" class="btn btn-sm btn-primary" data-wh-tab="fact_outgoing_part" onclick="renderWarehouseTable('fact_outgoing_part')">
              fact_outgoing_part (270)
            </button>
            <button type="button" class="btn btn-sm" data-wh-tab="fact_outgoing_lot" onclick="renderWarehouseTable('fact_outgoing_lot')">
              fact_outgoing_lot (345 Lot 1NF)
            </button>
            <button type="button" class="btn btn-sm" data-wh-tab="dim_invoice" onclick="renderWarehouseTable('dim_invoice')">
              dim_invoice (20)
            </button>
            <button type="button" class="btn btn-sm" data-wh-tab="dim_part" onclick="renderWarehouseTable('dim_part')">
              dim_part
            </button>
            <button type="button" class="btn btn-sm" data-wh-tab="dim_customer" onclick="renderWarehouseTable('dim_customer')">
              dim_customer (8)
            </button>
          </div>
        </div>

        <div class="schema-diagram" aria-label="Arsitektur Star Schema Outgoing">
          <div class="schema-box is-fact">
            <h4><span>FACT_OUTGOING_PART</span> <span class="badge badge-success" id="whCountFactPart">270 baris</span></h4>
            <ul class="schema-cols">
              <li><span>🔑 fact_part_id</span> <span>PK</span></li>
              <li><span>🔗 invoice_key, customer_key</span> <span>FK</span></li>
              <li><span>🔗 part_key, pic_key</span> <span>FK</span></li>
              <li><span>po_split_lines</span> <span>INT (1..3)</span></li>
              <li><span>qty_order_total</span> <span>INT</span></li>
              <li><span>qty_scan_clean</span> <span>INT (Dedup)</span></li>
              <li><span>overcount_saved</span> <span>INT</span></li>
            </ul>
          </div>

          <div class="schema-box is-fact">
            <h4><span>FACT_OUTGOING_LOT (1NF)</span> <span class="badge badge-info" id="whCountFactLot">345 baris</span></h4>
            <ul class="schema-cols">
              <li><span>🔑 fact_lot_id</span> <span>PK</span></li>
              <li><span>invoice_no, part_no</span> <span>VARCHAR</span></li>
              <li><span>lot_full_code</span> <span>VARCHAR</span></li>
              <li><span>primary_lot</span> <span>VARCHAR</span></li>
              <li><span>nested_sub_lot</span> <span>VARCHAR NULL</span></li>
              <li><span>lot_qty</span> <span>INT</span></li>
            </ul>
          </div>

          <div class="schema-box">
            <h4><span>DIM_INVOICE</span> <span class="badge badge-info" id="whCountDimInvoice">20 baris</span></h4>
            <ul class="schema-cols">
              <li><span>🔑 invoice_key</span> <span>PK</span></li>
              <li><span>invoice_no, delivery_no</span> <span>UNIQUE</span></li>
              <li><span>registered_date, lead_time_days</span> <span>INT</span></li>
              <li><span>summary_qty_part</span> <span>INT</span></li>
              <li><span>raw_detail_rows, split_extra_rows</span> <span>INT</span></li>
            </ul>
          </div>

          <div class="schema-box">
            <h4><span>DIM_PART</span> <span class="badge badge-info" id="whCountDimPart">0 baris</span></h4>
            <ul class="schema-cols">
              <li><span>🔑 part_key</span> <span>PK</span></li>
              <li><span>part_no</span> <span>UNIQUE</span></li>
              <li><span>part_name_clean</span> <span>VARCHAR</span></li>
              <li><span>part_category</span> <span>VARCHAR</span></li>
              <li><span>was_cleaned</span> <span>TINYINT</span></li>
            </ul>
          </div>
        </div>

        <div tabindex="0" role="region" aria-labelledby="warehouseTableCaption" class="table-wrapper">
          <table id="warehouseTable">
            <caption id="warehouseTableCaption">fact_outgoing_part</caption>
            <thead id="warehouseTableHead"></thead>
            <tbody id="warehouseTableBody"></tbody>
          </table>
        </div>
      </section>

    </div>

    <!-- ======================================================================
         VIEW 4: LATIHAN PRAKTIK DATA ENGINEER & SQL LAB
         ====================================================================== -->
    <div class="view-section" data-view-section="exercises" hidden>

      <section class="panel" aria-labelledby="exHeading">
        <div class="panel-header">
          <div>
            <h3 id="exHeading">Latihan Praktik Data Engineer (Studi Kasus Outgoing Delivery)</h3>
            <p>Selesaikan 4 tantangan analisis &amp; audit data di bawah ini. Klik <strong>"Muat &amp; Verifikasi Query"</strong> pada tiap kartu atau modifikasi query-nya di SQL Lab!</p>
          </div>
        </div>

        <div class="exercises-grid">
          <div class="exercise-card" id="card_ex1_reconcile">
            <div>
              <span class="badge badge-warning" id="badge_ex1_reconcile">Tantangan #1 • Rekonsiliasi</span>
              <h4>Temukan 6 Invoice yang Selisih Baris Summary vs Detail</h4>
              <p>Gunakan tabel <code>dim_invoice</code> untuk mencari invoice mana saja yang memiliki <code>split_extra_rows &gt; 0</code>.</p>
            </div>
            <button type="button" class="btn btn-sm btn-primary" onclick="runExercise('ex1_reconcile')">
              Muat &amp; Verifikasi Query #1
            </button>
          </div>

          <div class="exercise-card" id="card_ex2_fanout">
            <div>
              <span class="badge badge-critical" id="badge_ex2_fanout">Tantangan #2 • Anti Double-Count</span>
              <h4>Audit 26 Part yang Terpecah (Hemat 61.300 pcs)</h4>
              <p>Tampilkan part dengan <code>po_split_lines &gt; 1</code> di <code>fact_outgoing_part</code> dan urutkan dari <code>overcount_saved</code> terbesar.</p>
            </div>
            <button type="button" class="btn btn-sm btn-primary" onclick="runExercise('ex2_fanout')">
              Muat &amp; Verifikasi Query #2
            </button>
          </div>

          <div class="exercise-card" id="card_ex3_lot_trace">
            <div>
              <span class="badge badge-info" id="badge_ex3_lot_trace">Tantangan #3 • Normalisasi 1NF</span>
              <h4>Lacak 10 Lot Bersarang (Sub-Lot di Dalam Kurung)</h4>
              <p>Query tabel <code>fact_outgoing_lot</code> untuk menampilkan pengiriman yang memiliki <code>nested_sub_lot != '-'</code>.</p>
            </div>
            <button type="button" class="btn btn-sm btn-primary" onclick="runExercise('ex3_lot_trace')">
              Muat &amp; Verifikasi Query #3
            </button>
          </div>

          <div class="exercise-card" id="card_ex4_window_rank">
            <div>
              <span class="badge badge-info" id="badge_ex4_window_rank">Tantangan #4 • CTE &amp; Window Function</span>
              <h4>Top 2 Part Terbanyak per Customer dengan <code>RANK() OVER</code></h4>
              <p>Gunakan <code>WITH</code> (CTE) dan <code>RANK() OVER (PARTITION BY customer_code ...)</code> untuk meranking part per pelanggan.</p>
            </div>
            <button type="button" class="btn btn-sm btn-primary" onclick="runExercise('ex4_window_rank')">
              Muat &amp; Verifikasi Query #4
            </button>
          </div>
        </div>

        <div>
          <div style="display:flex; justify-content:space-between; align-items:center; margin-block-end:0.4rem;">
            <label for="sqlQueryInput" style="font-size:0.82rem; font-weight:700;">
              SQL Lab Editor (Tabel tersedia: <code>fact_outgoing_part</code>, <code>fact_outgoing_lot</code>, <code>dim_invoice</code>, <code>dim_part</code>, <code>dim_customer</code>, <code>raw_outgoing_detail</code>):
            </label>
            <button type="button" class="btn btn-sm btn-accent" onclick="executeSqlLab()">
              ▶ Jalankan Query SQL Bebas
            </button>
          </div>
          <textarea id="sqlQueryInput" class="sql-editor" spellcheck="false"></textarea>
        </div>

        <div style="display:flex; justify-content:space-between; align-items:center;">
          <strong>Hasil Eksekusi Query SQL:</strong>
          <span id="sqlResultMeta" style="font-size:0.8rem; color:var(--text-muted);">Siap</span>
        </div>

        <div tabindex="0" role="region" aria-labelledby="sqlResCaption" class="table-wrapper">
          <table>
            <caption id="sqlResCaption">Hasil Query SQL Lab</caption>
            <thead id="sqlResultHead"></thead>
            <tbody id="sqlResultBody"></tbody>
          </table>
        </div>
      </section>

    </div>

  </main>

  <!-- MODAL: JSON Diff Inspector (Light-dismiss via closedby="any") -->
  <dialog id="rowInspectorDialog" closedby="any" aria-labelledby="dialogRowTitle">
    <div class="dialog-header">
      <h2 id="dialogRowTitle">Inspeksi Transformasi Baris</h2>
      <form method="dialog">
        <button type="submit" class="btn btn-sm">Tutup (Esc)</button>
      </form>
    </div>
    <p id="dialogRowSummary" style="font-size:0.84rem; color:var(--text-muted); margin-block-start:0;"></p>
    <div class="diff-grid">
      <div>
        <strong style="font-size:0.8rem; display:block; margin-block-end:0.35rem;">SEBELUM (Raw Excel Detail):</strong>
        <pre id="dialogBeforeJson" class="diff-box"></pre>
      </div>
      <div>
        <strong style="font-size:0.8rem; display:block; margin-block-end:0.35rem;">SESUDAH (Clean Star Schema &amp; 1NF Lots):</strong>
        <pre id="dialogAfterJson" class="diff-box"></pre>
      </div>
    </div>
  </dialog>

  <script src="assets/app.js"></script>
</body>
</html>
