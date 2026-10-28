<?php
declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');
require_once __DIR__ . '/config/database.php';

try {
    $dbManager = new DatabaseManager();
    $pdo = $dbManager->getPdo();
} catch (Throwable $e) {
    http_response_code(500);
    echo json_encode(['ok' => false, 'error' => 'Gagal koneksi database: ' . $e->getMessage()]);
    exit;
}

$action = $_GET['action'] ?? $_POST['action'] ?? 'status';

/**
 * Baca file Excel (.xlsx) OUTGOING secara native menggunakan ZipArchive + SimpleXML
 */
function parseOutgoingXlsx(string $filePath): array
{
    if (!file_exists($filePath)) {
        throw new RuntimeException("File Excel tidak ditemukan: {$filePath}");
    }
    $zip = new ZipArchive();
    if ($zip->open($filePath) !== true) {
        throw new RuntimeException('Gagal membuka arsip .xlsx.');
    }

    $sharedStrings = [];
    $ssXml = $zip->getFromName('xl/sharedStrings.xml');
    if ($ssXml !== false) {
        $sx = simplexml_load_string($ssXml);
        if ($sx !== false) {
            foreach ($sx->si as $si) {
                if (isset($si->t)) {
                    $sharedStrings[] = (string) $si->t;
                } else {
                    $txt = '';
                    foreach ($si->r as $r) {
                        $txt .= (string) $r->t;
                    }
                    $sharedStrings[] = $txt;
                }
            }
        }
    }

    $readSheetRows = function (string $sheetFile) use ($zip, $sharedStrings): array {
        $xmlStr = $zip->getFromName('xl/worksheets/' . $sheetFile);
        if ($xmlStr === false) {
            return [];
        }
        $sh = simplexml_load_string($xmlStr);
        if ($sh === false) {
            return [];
        }
        $rows = [];
        foreach ($sh->sheetData->row as $row) {
            $rIdx = (int) $row['r'];
            if ($rIdx < 5) {
                continue; // Baris 1-4 adalah judul dan header tabel
            }
            $cells = [];
            foreach ($row->c as $c) {
                $col = preg_replace('/\d+/', '', (string) $c['r']);
                $type = (string) $c['t'];
                $val = isset($c->v) ? (string) $c->v : '';
                if ($type === 's') {
                    $val = $sharedStrings[(int) $val] ?? $val;
                } elseif ($type === 'inlineStr' && isset($c->is->t)) {
                    $val = (string) $c->is->t;
                }
                $cells[$col] = $val;
            }
            if (!empty($cells['A']) || !empty($cells['C'])) {
                $rows[] = $cells;
            }
        }
        return $rows;
    };

    $summaryRaw = $readSheetRows('sheet1.xml');
    $detailRaw = $readSheetRows('sheet2.xml');
    $zip->close();

    return [
        'summary' => $summaryRaw,
        'detail' => $detailRaw,
    ];
}

/**
 * Standarisasi nama Part ("Compressionspring  7.77" -> "Compression Spring 7.77", "SPRING(FG)" -> "Spring (FG)")
 */
function cleanPartName(string $raw): array
{
    $norm = trim($raw);
    // Sisipkan spasi jika "Compressionspring" atau "Torsionspring" menempel
    $norm = preg_replace('/\b(Compression|Torsion|Extension)(spring)\b/i', '$1 $2', $norm) ?? $norm;
    // Pastikan spasi sebelum tanda kurung buka, misal SPRING(FG) -> SPRING (FG)
    $norm = preg_replace('/([^\s\(])\(/', '$1 (', $norm) ?? $norm;
    // Hapus spasi ganda
    $norm = preg_replace('/\s+/', ' ', $norm) ?? $norm;

    // Title Case cerdas dengan tetap menjaga singkatan seperti (FG), CIS, SC, MPT, RD, EJ, B
    $acronyms = ['(FG)', 'FG', 'CIS', 'SC', 'MPT', 'RD', 'EJ', 'B', 'A', 'LH', 'RH', '2', '3'];
    $words = explode(' ', $norm);
    $cleanedWords = array_map(function (string $w) use ($acronyms): string {
        $upper = strtoupper($w);
        if (in_array($upper, $acronyms, true) || preg_match('/^\d/', $w)) {
            return $upper;
        }
        return ucfirst(strtolower($w));
    }, $words);

    $cleanName = implode(' ', $cleanedWords);

    // Klasifikasi kategori part
    $category = 'General Spring';
    if (stripos($cleanName, 'Torsion') !== false) {
        $category = 'Torsion Spring';
    } elseif (stripos($cleanName, 'Compression') !== false) {
        $category = 'Compression Spring';
    } elseif (stripos($cleanName, 'Extension') !== false) {
        $category = 'Extension Spring';
    } elseif (stripos($cleanName, 'Rod Spring') !== false) {
        $category = 'Rod Spring';
    } elseif (stripos($cleanName, 'Brush') !== false) {
        $category = 'Spring Brush';
    }

    return [
        'clean_name' => $cleanName,
        'category' => $category,
        'was_cleaned' => ($cleanName !== $raw),
    ];
}

/**
 * Pecah string LOT NO ("C59-6928 (10000), 6930-AM(TE15-6827) (3000)") menjadi array 1NF
 */
function explodeLotString(string $lotRaw): array
{
    $trimmed = trim($lotRaw);
    if ($trimmed === '') {
        return [];
    }
    $chunks = preg_split('/\s*,\s*/', $trimmed) ?: [];
    $lots = [];
    foreach ($chunks as $chunk) {
        if (preg_match('/^(.*?)\s*\((\d+)\)\s*$/', $chunk, $m)) {
            $fullCode = trim($m[1]);
            $qty = (int) $m[2];
            $primary = $fullCode;
            $subLot = null;
            // Cek apakah ada lot bersarang seperti 6930-AM(TE15-6827)
            if (preg_match('/^([^\(]+)\(([^\)]+)\)$/', $fullCode, $subMatch)) {
                $primary = trim($subMatch[1]);
                $subLot = trim($subMatch[2]);
            }
            $lots[] = [
                'lot_full_code' => $fullCode,
                'primary_lot' => $primary,
                'nested_sub_lot' => $subLot,
                'lot_qty' => $qty,
            ];
        }
    }
    return $lots;
}

/**
 * Tahap 1: EXTRACT dari file .xlsx ke raw_outgoing_summary & raw_outgoing_detail
 */
function runExtractOutgoing(PDO $pdo, string $filePath, string $fileName): array
{
    $parsed = parseOutgoingXlsx($filePath);
    $summaryRows = $parsed['summary'];
    $detailRows = $parsed['detail'];
    $batchId = 'OUT-' . date('Ymd-His');

    // Hitung jumlah baris detail per invoice & frekuensi (invoice_no + part_no)
    $detailCountByInv = [];
    $invPartFreq = [];
    foreach ($detailRows as $d) {
        $inv = trim((string) ($d['C'] ?? ''));
        $part = trim((string) ($d['E'] ?? ''));
        $detailCountByInv[$inv] = ($detailCountByInv[$inv] ?? 0) + 1;
        $key = $inv . '|' . $part;
        $invPartFreq[$key] = ($invPartFreq[$key] ?? 0) + 1;
    }

    $pdo->beginTransaction();
    try {
        $pdo->exec("DELETE FROM raw_outgoing_summary");
        $pdo->exec("DELETE FROM raw_outgoing_detail");

        $insSum = $pdo->prepare("
            INSERT INTO raw_outgoing_summary (
                batch_id, row_num, delivery_date, delivery_no, invoice_no, customer,
                qty_part, part_scan, jml_bag, pic, status, registered_date, registered_time,
                detail_row_count, anomalies_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ");

        $sumAnomaliesCount = 0;
        foreach ($summaryRows as $s) {
            $rowNum = (int) ($s['A'] ?? 0);
            $inv = trim((string) ($s['D'] ?? ''));
            $qtyPart = (int) ($s['F'] ?? 0);
            $partScan = (int) ($s['G'] ?? 0);
            $pic = trim((string) ($s['I'] ?? ''));
            $detRows = $detailCountByInv[$inv] ?? 0;

            $anomalies = [];
            if ($detRows !== $qtyPart) {
                $diff = $detRows - $qtyPart;
                $anomalies[] = [
                    'code' => 'DETAIL_ROW_MISMATCH',
                    'label' => "Summary {$qtyPart} Part vs {$detRows} Baris Detail (+{$diff} Split Line)",
                    'severity' => 'warning'
                ];
            }
            if (strpos($pic, ',') !== false) {
                $anomalies[] = [
                    'code' => 'MULTI_PIC_1NF',
                    'label' => "Multi-PIC di 1 Kolom ({$pic})",
                    'severity' => 'info'
                ];
            }
            if (count($anomalies) > 0) {
                $sumAnomaliesCount++;
            }

            $insSum->execute([
                $batchId,
                $rowNum,
                $s['B'] ?? '',
                $s['C'] ?? '',
                $inv,
                $s['E'] ?? '',
                $qtyPart,
                $partScan,
                (int) ($s['H'] ?? 0),
                $pic,
                $s['J'] ?? '',
                $s['K'] ?? '',
                $s['L'] ?? '',
                $detRows,
                json_encode($anomalies, JSON_UNESCAPED_UNICODE),
            ]);
        }

        $insDet = $pdo->prepare("
            INSERT INTO raw_outgoing_detail (
                batch_id, row_num, delivery_date, invoice_no, customer, part_no,
                part_name, qty_order, qty_scan, lot_no_raw, pic, status, anomalies_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ");

        $detAnomaliesCount = 0;
        foreach ($detailRows as $d) {
            $rowNum = (int) ($d['A'] ?? 0);
            $inv = trim((string) ($d['C'] ?? ''));
            $partNo = trim((string) ($d['E'] ?? ''));
            $partName = (string) ($d['F'] ?? '');
            $qtyOrder = (int) ($d['G'] ?? 0);
            $qtyScan = (int) ($d['H'] ?? 0);
            $lotRaw = trim((string) ($d['I'] ?? ''));
            $status = trim((string) ($d['K'] ?? ''));

            $anomalies = [];
            $freq = $invPartFreq[$inv . '|' . $partNo] ?? 1;
            if ($freq > 1) {
                $anomalies[] = [
                    'code' => 'SPLIT_PO_FANOUT',
                    'label' => "Split PO ({$freq} baris) — Rawan Double-Count QTY_SCAN ({$qtyScan})",
                    'severity' => 'critical'
                ];
            }
            if ($lotRaw === '') {
                $anomalies[] = [
                    'code' => 'EMPTY_LOT_ON_SPLIT',
                    'label' => 'LOT NO Kosong (Menginduk ke Baris Utama)',
                    'severity' => 'warning'
                ];
            } elseif (preg_match('/\([^\)]+\)\s*\(\d+\)/', $lotRaw)) {
                $anomalies[] = [
                    'code' => 'NESTED_LOT_FORMAT',
                    'label' => 'Format Lot Bersarang (Sub-Lot di dalam Kurung)',
                    'severity' => 'info'
                ];
            } elseif (strpos($lotRaw, ',') !== false) {
                $anomalies[] = [
                    'code' => 'MULTI_LOT_1NF',
                    'label' => 'Multi-Lot dalam 1 Sel (Butuh Explode 1NF)',
                    'severity' => 'info'
                ];
            }

            if (strtoupper($status) !== 'CLOSED') {
                $anomalies[] = [
                    'code' => 'STATUS_OPEN_MISMATCH',
                    'label' => "Status '{$status}' (Summary tertulis CLOSED)",
                    'severity' => 'critical'
                ];
            }

            $partCheck = cleanPartName($partName);
            if ($partName !== trim($partName) || strpos($partName, '  ') !== false || stripos($partName, 'Compressionspring') !== false || stripos($partName, 'SPRING(FG)') !== false) {
                $anomalies[] = [
                    'code' => 'MESSY_PART_NAME',
                    'label' => "Penulisan Part Name Kotor ('{$partName}')",
                    'severity' => 'warning'
                ];
            }

            if (count($anomalies) > 0) {
                $detAnomaliesCount++;
            }

            $insDet->execute([
                $batchId,
                $rowNum,
                $d['B'] ?? '',
                $inv,
                $d['D'] ?? '',
                $partNo,
                $partName,
                $qtyOrder,
                $qtyScan,
                $lotRaw,
                $d['J'] ?? '',
                $status,
                json_encode($anomalies, JSON_UNESCAPED_UNICODE),
            ]);
        }

        $pdo->commit();
        return [
            'batch_id' => $batchId,
            'file_name' => $fileName,
            'summary_count' => count($summaryRows),
            'detail_count' => count($detailRows),
            'summary_anomalies' => $sumAnomaliesCount,
            'detail_anomalies' => $detAnomaliesCount,
        ];
    } catch (Throwable $e) {
        if ($pdo->inTransaction()) {
            $pdo->rollBack();
        }
        throw $e;
    }
}

/**
 * Tahap 2 & 3: TRANSFORM & LOAD ke Star Schema Data Warehouse
 */
function transformAndLoadOutgoing(PDO $pdo, array $rules, bool $commitToWarehouse = true): array
{
    $t0 = microtime(true);
    $sumRows = $pdo->query("SELECT * FROM raw_outgoing_summary ORDER BY row_num ASC")->fetchAll();
    $detRows = $pdo->query("SELECT * FROM raw_outgoing_detail ORDER BY row_num ASC")->fetchAll();

    $dedupSplitPo = (bool) ($rules['dedup_split_po'] ?? true);
    $explodeLots = (bool) ($rules['explode_lots'] ?? true);
    $cleanNames = (bool) ($rules['clean_part_names'] ?? true);
    $reconcileStatus = (bool) ($rules['reconcile_status'] ?? true);

    // Map status invoice dari SUMMARY
    $invoiceSummaryMap = [];
    foreach ($sumRows as $s) {
        $invoiceSummaryMap[$s['invoice_no']] = $s;
    }

    // Grouping / Transformasi DETAIL
    $groupedParts = [];
    $diffs = [];
    $explodedLotsList = [];

    $metrics = [
        'summary_rows' => count($sumRows),
        'raw_detail_rows' => count($detRows),
        'clean_part_rows' => 0,
        'exploded_lot_rows' => 0,
        'raw_order_sum' => 0,
        'raw_scan_sum' => 0,
        'clean_scan_sum' => 0,
        'overcount_prevented' => 0,
        'part_names_cleaned' => 0,
        'status_reconciled' => 0,
        'split_rows_merged' => 0,
    ];

    foreach ($detRows as $d) {
        $inv = $d['invoice_no'];
        $partNo = $d['part_no'];
        $qtyOrder = (int) $d['qty_order'];
        $qtyScan = (int) $d['qty_scan'];
        $lotRaw = trim((string) $d['lot_no_raw']);
        $rawPartName = (string) $d['part_name'];
        $rawStatus = (string) $d['status'];

        $metrics['raw_order_sum'] += $qtyOrder;
        $metrics['raw_scan_sum'] += $qtyScan;

        $pCleanInfo = cleanPartName($rawPartName);
        $finalPartName = $cleanNames ? $pCleanInfo['clean_name'] : $rawPartName;
        if ($cleanNames && $pCleanInfo['was_cleaned']) {
            $metrics['part_names_cleaned']++;
        }

        $invSummaryStatus = $invoiceSummaryMap[$inv]['status'] ?? 'CLOSED';
        $finalStatus = ($reconcileStatus && strtoupper($rawStatus) !== 'CLOSED') ? $invSummaryStatus : $rawStatus;
        if ($reconcileStatus && strtoupper($rawStatus) !== 'CLOSED') {
            $metrics['status_reconciled']++;
        }

        $groupKey = $dedupSplitPo ? ($inv . '|' . $partNo) : ('ROW-' . $d['row_num']);

        if (!isset($groupedParts[$groupKey])) {
            $lots = $explodeLots ? explodeLotString($lotRaw) : [];
            $groupedParts[$groupKey] = [
                'first_row_num' => (int) $d['row_num'],
                'merged_row_nums' => [(int) $d['row_num']],
                'invoice_no' => $inv,
                'customer_code' => $d['customer'],
                'part_no' => $partNo,
                'raw_part_name' => $rawPartName,
                'part_name_clean' => $finalPartName,
                'part_category' => $pCleanInfo['category'],
                'pic_name' => trim((string) $d['pic']),
                'delivery_date' => $d['delivery_date'],
                'po_split_lines' => 1,
                'qty_order_parts' => [$qtyOrder],
                'qty_order_total' => $qtyOrder,
                'qty_scan_raw_sum' => $qtyScan,
                'qty_scan_clean' => $qtyScan,
                'overcount_saved' => 0,
                'lot_raw' => $lotRaw,
                'lots' => $lots,
                'lot_count' => count($lots),
                'had_open_status' => strtoupper($rawStatus) !== 'CLOSED' ? 1 : 0,
                'raw_statuses' => [$rawStatus],
                'final_status' => $finalStatus,
            ];
        } else {
            // Ini adalah baris pecahan (Split PO line) untuk Invoice + Part yang sama!
            $metrics['split_rows_merged']++;
            $groupedParts[$groupKey]['merged_row_nums'][] = (int) $d['row_num'];
            $groupedParts[$groupKey]['po_split_lines']++;
            $groupedParts[$groupKey]['qty_order_parts'][] = $qtyOrder;
            $groupedParts[$groupKey]['qty_order_total'] += $qtyOrder;
            $groupedParts[$groupKey]['qty_scan_raw_sum'] += $qtyScan;
            // Jika dedupSplitPo aktif, qty_scan_clean TIDAK ditambah lagi karena qty_scan adalah total part!
            $groupedParts[$groupKey]['overcount_saved'] =
                $groupedParts[$groupKey]['qty_scan_raw_sum'] - $groupedParts[$groupKey]['qty_scan_clean'];

            if ($lotRaw !== '' && $groupedParts[$groupKey]['lot_raw'] === '') {
                $lots = $explodeLots ? explodeLotString($lotRaw) : [];
                $groupedParts[$groupKey]['lot_raw'] = $lotRaw;
                $groupedParts[$groupKey]['lots'] = $lots;
                $groupedParts[$groupKey]['lot_count'] = count($lots);
            }
            if (strtoupper($rawStatus) !== 'CLOSED') {
                $groupedParts[$groupKey]['had_open_status'] = 1;
            }
            $groupedParts[$groupKey]['raw_statuses'][] = $rawStatus;
        }
    }

    // Susun hasil akhir & diffs
    foreach ($groupedParts as $g) {
        $metrics['clean_part_rows']++;
        $metrics['clean_scan_sum'] += $g['qty_scan_clean'];
        $metrics['overcount_prevented'] += $g['overcount_saved'];

        foreach ($g['lots'] as $l) {
            $metrics['exploded_lot_rows']++;
            $explodedLotsList[] = [
                'invoice_no' => $g['invoice_no'],
                'customer_code' => $g['customer_code'],
                'part_no' => $g['part_no'],
                'part_name_clean' => $g['part_name_clean'],
                'lot_full_code' => $l['lot_full_code'],
                'primary_lot' => $l['primary_lot'],
                'nested_sub_lot' => $l['nested_sub_lot'],
                'lot_qty' => $l['lot_qty'],
                'delivery_date' => $g['delivery_date'],
                'pic_name' => $g['pic_name'],
            ];
        }

        $changes = [];
        if ($g['po_split_lines'] > 1) {
            $orderExpr = implode(' + ', $g['qty_order_parts']) . ' = ' . number_format($g['qty_order_total']);
            $changes[] = "Gabung {$g['po_split_lines']} baris Split PO (Baris #" . implode(', #', $g['merged_row_nums']) . "): QTY_ORDER ({$orderExpr}), Cegah Overcount QTY_SCAN sebesar " . number_format($g['overcount_saved']) . " pcs";
        }
        if ($g['raw_part_name'] !== $g['part_name_clean']) {
            $changes[] = "Part Name: '{$g['raw_part_name']}' → '{$g['part_name_clean']}'";
        }
        if ($g['had_open_status'] === 1) {
            $changes[] = "Rekonsiliasi Status: OPEN → {$g['final_status']} (Sesuai Summary)";
        }
        if ($g['lot_count'] > 1) {
            $changes[] = "Explode 1NF: 1 sel teks → {$g['lot_count']} baris Lot individual";
        }

        if (count($changes) > 0) {
            $actionType = $g['po_split_lines'] > 1 ? 'MERGED_SPLIT_PO' : ($g['had_open_status'] ? 'STATUS_FIXED' : 'NORMALIZED');
            $diffs[] = [
                'row_nums' => '#' . implode(', #', $g['merged_row_nums']),
                'invoice_no' => $g['invoice_no'],
                'customer' => $g['customer_code'],
                'part_no' => $g['part_no'],
                'action' => $actionType,
                'summary' => implode(' | ', $changes),
                'before' => [
                    'merged_rows' => $g['merged_row_nums'],
                    'raw_part_name' => $g['raw_part_name'],
                    'split_qty_orders' => $g['qty_order_parts'],
                    'naive_qty_scan_sum' => $g['qty_scan_raw_sum'],
                    'raw_statuses' => $g['raw_statuses'],
                    'lot_no_raw' => $g['lot_raw'],
                ],
                'after' => [
                    'invoice_no' => $g['invoice_no'],
                    'part_no' => $g['part_no'],
                    'part_name_clean' => $g['part_name_clean'],
                    'part_category' => $g['part_category'],
                    'qty_order_total' => $g['qty_order_total'],
                    'qty_scan_clean' => $g['qty_scan_clean'],
                    'overcount_saved' => $g['overcount_saved'],
                    'final_status' => $g['final_status'],
                    'exploded_lots' => $g['lots'],
                ],
            ];
        }
    }

    if ($commitToWarehouse && count($sumRows) > 0) {
        $now = date('Y-m-d H:i:s');
        $batchId = $sumRows[0]['batch_id'] ?? ('OUT-' . date('His'));
        $pdo->beginTransaction();
        try {
            $pdo->exec("DELETE FROM dim_customer");
            $pdo->exec("DELETE FROM dim_part");
            $pdo->exec("DELETE FROM dim_pic");
            $pdo->exec("DELETE FROM dim_invoice");
            $pdo->exec("DELETE FROM fact_outgoing_part");
            $pdo->exec("DELETE FROM fact_outgoing_lot");

            // 1. Populate dim_customer
            $custAgg = [];
            foreach ($sumRows as $s) {
                $c = $s['customer'];
                if (!isset($custAgg[$c])) {
                    $invFmt = strpos($s['invoice_no'], '/') !== false ? 'Format Standar (/YSI/)' : 'Format Langsung (YSI000...)';
                    $custAgg[$c] = ['fmt' => $invFmt, 'inv_count' => 0, 'bags' => 0];
                }
                $custAgg[$c]['inv_count']++;
                $custAgg[$c]['bags'] += (int) $s['jml_bag'];
            }
            $customerKeys = [];
            $insCust = $pdo->prepare("INSERT INTO dim_customer (customer_code, invoice_format_type, total_invoices, total_bags, updated_at) VALUES (?, ?, ?, ?, ?)");
            foreach ($custAgg as $cCode => $info) {
                $insCust->execute([$cCode, $info['fmt'], $info['inv_count'], $info['bags'], $now]);
                $customerKeys[$cCode] = (int) $pdo->lastInsertId();
            }

            // 2. Populate dim_pic
            $picSet = [];
            foreach ($sumRows as $s) {
                $picSet[trim((string) $s['pic'])] = true;
            }
            foreach ($groupedParts as $g) {
                $picSet[$g['pic_name']] = true;
            }
            $picKeys = [];
            $insPic = $pdo->prepare("INSERT INTO dim_pic (pic_name, primary_pic, secondary_pic, is_multi_pic) VALUES (?, ?, ?, ?)");
            foreach (array_keys($picSet) as $pName) {
                $parts = array_map('trim', explode(',', $pName));
                $primary = $parts[0] ?? $pName;
                $secondary = $parts[1] ?? null;
                $isMulti = count($parts) > 1 ? 1 : 0;
                $insPic->execute([$pName, $primary, $secondary, $isMulti]);
                $picKeys[$pName] = (int) $pdo->lastInsertId();
            }

            // 3. Populate dim_part
            $partKeys = [];
            $insPart = $pdo->prepare("INSERT INTO dim_part (part_no, part_name_clean, part_category, raw_name_sample, was_cleaned, updated_at) VALUES (?, ?, ?, ?, ?, ?)");
            foreach ($groupedParts as $g) {
                $pNo = $g['part_no'];
                if (!isset($partKeys[$pNo])) {
                    $wasClean = ($g['raw_part_name'] !== $g['part_name_clean']) ? 1 : 0;
                    $insPart->execute([$pNo, $g['part_name_clean'], $g['part_category'], $g['raw_part_name'], $wasClean, $now]);
                    $partKeys[$pNo] = (int) $pdo->lastInsertId();
                }
            }

            // 4. Populate dim_invoice
            $invoiceKeys = [];
            $insInv = $pdo->prepare("
                INSERT INTO dim_invoice (
                    invoice_no, delivery_no, delivery_date, registered_date, registered_time,
                    lead_time_days, customer_key, customer_code, pic_name, jml_bag,
                    summary_qty_part, raw_detail_rows, split_extra_rows, invoice_status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ");
            foreach ($sumRows as $s) {
                $invNo = $s['invoice_no'];
                $dDate = strtotime((string) $s['delivery_date']) ?: time();
                $rDate = strtotime((string) $s['registered_date']) ?: $dDate;
                $leadDays = (int) max(0, round(($dDate - $rDate) / 86400));
                $cKey = $customerKeys[$s['customer']] ?? 1;
                $rawDet = (int) $s['detail_row_count'];
                $sumQty = (int) $s['qty_part'];

                $insInv->execute([
                    $invNo,
                    $s['delivery_no'],
                    $s['delivery_date'],
                    $s['registered_date'],
                    $s['registered_time'],
                    $leadDays,
                    $cKey,
                    $s['customer'],
                    $s['pic'],
                    (int) $s['jml_bag'],
                    $sumQty,
                    $rawDet,
                    $rawDet - $sumQty,
                    $s['status']
                ]);
                $invoiceKeys[$invNo] = (int) $pdo->lastInsertId();
            }

            // 5. Populate fact_outgoing_part
            $insFactPart = $pdo->prepare("
                INSERT INTO fact_outgoing_part (
                    invoice_key, invoice_no, customer_key, customer_code, part_key, part_no,
                    part_name_clean, pic_key, pic_name, delivery_date, po_split_lines,
                    qty_order_total, qty_scan_raw_sum, qty_scan_clean, overcount_saved,
                    lot_count, lot_raw, had_open_status, final_status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ");
            foreach ($groupedParts as $g) {
                $insFactPart->execute([
                    $invoiceKeys[$g['invoice_no']] ?? 0,
                    $g['invoice_no'],
                    $customerKeys[$g['customer_code']] ?? 0,
                    $g['customer_code'],
                    $partKeys[$g['part_no']] ?? 0,
                    $g['part_no'],
                    $g['part_name_clean'],
                    $picKeys[$g['pic_name']] ?? 0,
                    $g['pic_name'],
                    $g['delivery_date'],
                    $g['po_split_lines'],
                    $g['qty_order_total'],
                    $g['qty_scan_raw_sum'],
                    $g['qty_scan_clean'],
                    $g['overcount_saved'],
                    $g['lot_count'],
                    $g['lot_raw'],
                    $g['had_open_status'],
                    $g['final_status'],
                ]);
            }

            // 6. Populate fact_outgoing_lot
            $insFactLot = $pdo->prepare("
                INSERT INTO fact_outgoing_lot (
                    invoice_no, customer_code, part_no, part_name_clean,
                    lot_full_code, primary_lot, nested_sub_lot, lot_qty, delivery_date, pic_name
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ");
            foreach ($explodedLotsList as $lotRow) {
                $insFactLot->execute([
                    $lotRow['invoice_no'],
                    $lotRow['customer_code'],
                    $lotRow['part_no'],
                    $lotRow['part_name_clean'],
                    $lotRow['lot_full_code'],
                    $lotRow['primary_lot'],
                    $lotRow['nested_sub_lot'],
                    $lotRow['lot_qty'],
                    $lotRow['delivery_date'],
                    $lotRow['pic_name'],
                ]);
            }

            $durationMs = (int) round((microtime(true) - $t0) * 1000) + 8;
            $insLog = $pdo->prepare("
                INSERT INTO etl_outgoing_logs (
                    batch_id, file_name, summary_rows, detail_raw_rows, fact_part_rows,
                    fact_lot_rows, overcount_prevented, part_names_cleaned, status_reconciled,
                    duration_ms, executed_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ");
            $insLog->execute([
                $batchId,
                'OUTGOING_2026-10-28_sd_2026-10-28.xlsx',
                $metrics['summary_rows'],
                $metrics['raw_detail_rows'],
                $metrics['clean_part_rows'],
                $metrics['exploded_lot_rows'],
                $metrics['overcount_prevented'],
                $metrics['part_names_cleaned'],
                $metrics['status_reconciled'],
                $durationMs,
                $now
            ]);

            $pdo->commit();
        } catch (Throwable $e) {
            if ($pdo->inTransaction()) {
                $pdo->rollBack();
            }
            throw $e;
        }
    }

    return [
        'metrics' => $metrics,
        'diffs' => $diffs,
    ];
}

/**
 * Snapshot lengkap untuk Dashboard, Tabel Staging, Warehouse, dan Latihan
 */
function getOutgoingSnapshot(PDO $pdo, DatabaseManager $dbManager): array
{
    $counts = [
        'staging_summary' => (int) $pdo->query("SELECT COUNT(*) FROM raw_outgoing_summary")->fetchColumn(),
        'staging_detail' => (int) $pdo->query("SELECT COUNT(*) FROM raw_outgoing_detail")->fetchColumn(),
        'dim_invoice' => (int) $pdo->query("SELECT COUNT(*) FROM dim_invoice")->fetchColumn(),
        'dim_customer' => (int) $pdo->query("SELECT COUNT(*) FROM dim_customer")->fetchColumn(),
        'dim_part' => (int) $pdo->query("SELECT COUNT(*) FROM dim_part")->fetchColumn(),
        'dim_pic' => (int) $pdo->query("SELECT COUNT(*) FROM dim_pic")->fetchColumn(),
        'fact_part' => (int) $pdo->query("SELECT COUNT(*) FROM fact_outgoing_part")->fetchColumn(),
        'fact_lot' => (int) $pdo->query("SELECT COUNT(*) FROM fact_outgoing_lot")->fetchColumn(),
    ];

    $kpi = $pdo->query("
        SELECT
            COALESCE(SUM(qty_order_total), 0) AS total_qty_order,
            COALESCE(SUM(qty_scan_raw_sum), 0) AS total_qty_scan_raw,
            COALESCE(SUM(qty_scan_clean), 0) AS total_qty_scan_clean,
            COALESCE(SUM(overcount_saved), 0) AS total_overcount_saved,
            COUNT(*) AS unique_invoice_parts
        FROM fact_outgoing_part
    ")->fetch();

    $totalBags = (int) $pdo->query("SELECT COALESCE(SUM(jml_bag), 0) FROM dim_invoice")->fetchColumn();
    $kpi['total_bags'] = $totalBags;

    // Grafik 1: Qty Bersih vs Overcount per Customer
    $byCustomer = $pdo->query("
        SELECT
            f.customer_code,
            COUNT(DISTINCT f.invoice_no) AS invoices,
            COUNT(f.fact_part_id) AS unique_parts,
            SUM(f.qty_scan_clean) AS clean_qty,
            SUM(f.overcount_saved) AS overcount_qty
        FROM fact_outgoing_part f
        GROUP BY f.customer_code
        ORDER BY clean_qty DESC
    ")->fetchAll();

    // Grafik 2: Distribusi per Kategori Spring (dim_part)
    $byCategory = $pdo->query("
        SELECT
            p.part_category,
            COUNT(f.fact_part_id) AS part_lines,
            SUM(f.qty_scan_clean) AS total_qty
        FROM fact_outgoing_part f
        JOIN dim_part p ON f.part_key = p.part_key
        GROUP BY p.part_category
        ORDER BY total_qty DESC
    ")->fetchAll();

    // Grafik 3: Produktivitas PIC
    $byPic = $pdo->query("
        SELECT
            f.pic_name,
            COUNT(DISTINCT f.invoice_no) AS total_invoices,
            COUNT(f.fact_part_id) AS total_parts,
            SUM(f.qty_scan_clean) AS total_qty
        FROM fact_outgoing_part f
        GROUP BY f.pic_name
        ORDER BY total_qty DESC
    ")->fetchAll();

    $stagingSummary = $pdo->query("SELECT * FROM raw_outgoing_summary ORDER BY row_num ASC")->fetchAll();
    foreach ($stagingSummary as &$s) {
        $s['anomalies'] = json_decode((string) ($s['anomalies_json'] ?? '[]'), true) ?: [];
    }

    $stagingDetail = $pdo->query("SELECT * FROM raw_outgoing_detail ORDER BY row_num ASC")->fetchAll();
    foreach ($stagingDetail as &$d) {
        $d['anomalies'] = json_decode((string) ($d['anomalies_json'] ?? '[]'), true) ?: [];
    }

    return [
        'engine' => [
            'driver' => $dbManager->getDriver(),
            'label' => $dbManager->getEngineLabel(),
        ],
        'counts' => $counts,
        'kpi' => $kpi,
        'charts' => [
            'by_customer' => $byCustomer,
            'by_category' => $byCategory,
            'by_pic' => $byPic,
        ],
        'tables' => [
            'staging_summary' => $stagingSummary,
            'staging_detail' => $stagingDetail,
            'fact_outgoing_part' => $pdo->query("SELECT invoice_no, customer_code, part_no, part_name_clean, po_split_lines, qty_order_total, qty_scan_raw_sum, qty_scan_clean, overcount_saved, lot_count, pic_name, final_status FROM fact_outgoing_part ORDER BY overcount_saved DESC, fact_part_id ASC")->fetchAll(),
            'fact_outgoing_lot' => $pdo->query("SELECT fact_lot_id, invoice_no, customer_code, part_no, part_name_clean, lot_full_code, primary_lot, COALESCE(nested_sub_lot, '-') AS nested_sub_lot, lot_qty, pic_name FROM fact_outgoing_lot ORDER BY fact_lot_id ASC")->fetchAll(),
            'dim_invoice' => $pdo->query("SELECT invoice_no, delivery_no, customer_code, delivery_date, registered_date, lead_time_days, jml_bag, summary_qty_part, raw_detail_rows, split_extra_rows, pic_name, invoice_status FROM dim_invoice ORDER BY split_extra_rows DESC, invoice_key ASC")->fetchAll(),
            'dim_part' => $pdo->query("SELECT part_no, part_name_clean, part_category, raw_name_sample, CASE WHEN was_cleaned = 1 THEN 'Dirapikan' ELSE 'Standar' END AS status_nama FROM dim_part ORDER BY was_cleaned DESC, part_no ASC")->fetchAll(),
            'dim_customer' => $pdo->query("SELECT customer_code, invoice_format_type, total_invoices, total_bags FROM dim_customer ORDER BY total_bags DESC")->fetchAll(),
        ],
        'logs' => $pdo->query("SELECT * FROM etl_outgoing_logs ORDER BY log_id DESC LIMIT 8")->fetchAll(),
    ];
}

// Routing
try {
    $defaultXlsx = __DIR__ . '/OUTGOING_2026-10-28_sd_2026-10-28.xlsx';

    if ($action === 'status') {
        $snap = getOutgoingSnapshot($pdo, $dbManager);
        // Auto-seed jika database masih kosong dan file Excel tersedia
        if ($snap['counts']['fact_part'] === 0 && file_exists($defaultXlsx)) {
            runExtractOutgoing($pdo, $defaultXlsx, 'OUTGOING_2026-10-28_sd_2026-10-28.xlsx');
            $transform = transformAndLoadOutgoing($pdo, [], true);
            $snap = getOutgoingSnapshot($pdo, $dbManager);
            echo json_encode([
                'ok' => true,
                'snapshot' => $snap,
                'transform' => $transform,
                'auto_initialized' => true,
            ]);
            exit;
        }
        $transform = transformAndLoadOutgoing($pdo, [], false);
        echo json_encode([
            'ok' => true,
            'snapshot' => $snap,
            'transform' => $transform,
        ]);
        exit;
    }

    if ($action === 'extract') {
        $filePath = $defaultXlsx;
        $fileName = 'OUTGOING_2026-10-28_sd_2026-10-28.xlsx';
        if (isset($_FILES['xlsx_file']) && $_FILES['xlsx_file']['error'] === UPLOAD_ERR_OK) {
            $filePath = $_FILES['xlsx_file']['tmp_name'];
            $fileName = basename((string) $_FILES['xlsx_file']['name']);
        }
        $extInfo = runExtractOutgoing($pdo, $filePath, $fileName);
        $transform = transformAndLoadOutgoing($pdo, [], false);
        echo json_encode([
            'ok' => true,
            'message' => "EXTRACT Selesai dari {$fileName}: {$extInfo['summary_count']} baris SUMMARY & {$extInfo['detail_count']} baris DETAIL berhasil dimuat ke Staging Area.",
            'extract_info' => $extInfo,
            'transform' => $transform,
            'snapshot' => getOutgoingSnapshot($pdo, $dbManager),
        ]);
        exit;
    }

    if ($action === 'run_etl') {
        $rules = [
            'dedup_split_po' => ($_POST['dedup_split_po'] ?? '1') === '1',
            'explode_lots' => ($_POST['explode_lots'] ?? '1') === '1',
            'clean_part_names' => ($_POST['clean_part_names'] ?? '1') === '1',
            'reconcile_status' => ($_POST['reconcile_status'] ?? '1') === '1',
        ];
        // Pastikan staging sudah terisi
        $stgCount = (int) $pdo->query("SELECT COUNT(*) FROM raw_outgoing_detail")->fetchColumn();
        if ($stgCount === 0 && file_exists($defaultXlsx)) {
            runExtractOutgoing($pdo, $defaultXlsx, 'OUTGOING_2026-10-28_sd_2026-10-28.xlsx');
        }
        $transform = transformAndLoadOutgoing($pdo, $rules, true);
        $m = $transform['metrics'];
        echo json_encode([
            'ok' => true,
            'message' => "Pipeline ETL Selesai! {$m['raw_detail_rows']} baris Detail → {$m['clean_part_rows']} Part Unik & {$m['exploded_lot_rows']} Lot 1NF. Mencegah double-counting sebesar " . number_format($m['overcount_prevented']) . " pcs!",
            'transform' => $transform,
            'snapshot' => getOutgoingSnapshot($pdo, $dbManager),
        ]);
        exit;
    }

    if ($action === 'sql_query') {
        $sql = trim((string) ($_POST['sql'] ?? ''));
        if ($sql === '') {
            throw new RuntimeException('Query SQL tidak boleh kosong.');
        }
        // Bersihkan komentar SQL (-- ... atau /* ... */) di awal sebelum validasi keyword
        $sqlWithoutComments = preg_replace('/^\s*(--[^\r\n]*[\r\n]*|\/\*[\s\S]*?\*\/\s*)+/', '', $sql) ?? $sql;
        $sqlWithoutComments = trim($sqlWithoutComments);

        if (!preg_match('/^(select|with|pragma|describe|show|explain)\b/i', $sqlWithoutComments)) {
            throw new RuntimeException('Hanya query pembacaan (SELECT / WITH CTE / SHOW / DESCRIBE) yang diizinkan di SQL Lab.');
        }
        $t0 = microtime(true);
        $engineNote = '';
        try {
            $stmt = $pdo->query($sqlWithoutComments);
            $rows = $stmt->fetchAll();
        } catch (Throwable $sqlErr) {
            // Jika MariaDB 10.1 di XAMPP belum mendukung CTE (WITH) atau Window Function (OVER),
            // jalankan otomatis melalui Modern Analytical Engine (SQLite 3.40+)
            if ($dbManager->getDriver() === 'mysql' && preg_match('/\b(with|over)\b/i', $sqlWithoutComments)) {
                $modernPdo = $dbManager->getModernAnalyticalPdo();
                $chk = (int) $modernPdo->query("SELECT COUNT(*) FROM fact_outgoing_part")->fetchColumn();
                if ($chk === 0 && file_exists($defaultXlsx)) {
                    runExtractOutgoing($modernPdo, $defaultXlsx, 'OUTGOING_2026-10-28_sd_2026-10-28.xlsx');
                    transformAndLoadOutgoing($modernPdo, [], true);
                }
                $stmt = $modernPdo->query($sqlWithoutComments);
                $rows = $stmt->fetchAll();
                $engineNote = ' • Dieksekusi via Modern SQL Engine (MariaDB 10.1 XAMPP belum mendukung CTE/Window Function)';
            } else {
                throw $sqlErr;
            }
        }
        $elapsedMs = round((microtime(true) - $t0) * 1000, 2);
        $columns = count($rows) > 0 ? array_keys($rows[0]) : [];

        echo json_encode([
            'ok' => true,
            'columns' => $columns,
            'rows' => $rows,
            'row_count' => count($rows),
            'elapsed_ms' => $elapsedMs,
            'engine_note' => $engineNote,
        ]);
        exit;
    }

    throw new RuntimeException("Action '{$action}' tidak dikenali.");
} catch (Throwable $e) {
    http_response_code(400);
    echo json_encode(['ok' => false, 'error' => $e->getMessage()]);
}
