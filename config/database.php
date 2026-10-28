<?php
declare(strict_types=1);

/**
 * DatabaseManager — Mengelola koneksi Data Warehouse (MySQL XAMPP dengan auto-fallback ke SQLite).
 * Mendukung Staging Area (SUMMARY & DETAIL) dan Star Schema Warehouse untuk data OUTGOING.
 */
class DatabaseManager
{
    private ?PDO $pdo = null;
    private string $driver = 'mysql';
    private string $engineLabel = 'MySQL (XAMPP)';
    private string $dbName = 'etl_studio_db';

    public function __construct()
    {
        $this->connect();
        $this->initializeSchema();
    }

    private function connect(): void
    {
        $host = getenv('ETL_DB_HOST') ?: '127.0.0.1';
        $port = getenv('ETL_DB_PORT') ?: '3306';
        $user = getenv('ETL_DB_USER') ?: 'root';
        $pass = getenv('ETL_DB_PASS') !== false ? getenv('ETL_DB_PASS') : '';

        try {
            $dsnNoDb = "mysql:host={$host};port={$port};charset=utf8mb4";
            $tempPdo = new PDO($dsnNoDb, $user, $pass, [
                PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
                PDO::ATTR_TIMEOUT => 2,
            ]);
            $tempPdo->exec("CREATE DATABASE IF NOT EXISTS `{$this->dbName}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci");
            $tempPdo = null;

            $dsn = "mysql:host={$host};port={$port};dbname={$this->dbName};charset=utf8mb4";
            $this->pdo = new PDO($dsn, $user, $pass, [
                PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
                PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
            ]);
            $this->driver = 'mysql';
            $this->engineLabel = "MySQL XAMPP ({$this->dbName})";
        } catch (Throwable $e) {
            $storageDir = dirname(__DIR__) . DIRECTORY_SEPARATOR . 'storage';
            if (!is_dir($storageDir)) {
                mkdir($storageDir, 0777, true);
            }
            $sqlitePath = $storageDir . DIRECTORY_SEPARATOR . 'etl_outgoing.sqlite';
            $this->pdo = new PDO("sqlite:{$sqlitePath}", null, null, [
                PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
                PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
            ]);
            $this->driver = 'sqlite';
            $this->engineLabel = 'SQLite Warehouse (Auto-Fallback)';
        }
    }

    public function getPdo(): PDO
    {
        return $this->pdo;
    }

    public function getDriver(): string
    {
        return $this->driver;
    }

    public function getEngineLabel(): string
    {
        return $this->engineLabel;
    }

    /**
     * Sediakan engine SQLite 3.40+ pendamping untuk mengeksekusi query SQL Modern
     * (CTE WITH & Window Functions OVER) jika MariaDB di XAMPP masih versi 10.1 (< 10.2).
     */
    public function getModernAnalyticalPdo(): PDO
    {
        if ($this->driver === 'sqlite') {
            return $this->pdo;
        }
        $storageDir = dirname(__DIR__) . DIRECTORY_SEPARATOR . 'storage';
        if (!is_dir($storageDir)) {
            mkdir($storageDir, 0777, true);
        }
        $sqlitePath = $storageDir . DIRECTORY_SEPARATOR . 'etl_outgoing.sqlite';
        $sqlitePdo = new PDO("sqlite:{$sqlitePath}", null, null, [
            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        ]);

        // Pastikan skema di SQLite pendamping juga terinisialisasi
        $prevPdo = $this->pdo;
        $prevDriver = $this->driver;
        $this->pdo = $sqlitePdo;
        $this->driver = 'sqlite';
        $this->initializeSchema();
        $this->pdo = $prevPdo;
        $this->driver = $prevDriver;

        return $sqlitePdo;
    }

    private function initializeSchema(): void
    {
        $autoInc = $this->driver === 'mysql'
            ? 'INT AUTO_INCREMENT PRIMARY KEY'
            : 'INTEGER PRIMARY KEY AUTOINCREMENT';

        // 1. STAGING: Sheet SUMMARY
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS raw_outgoing_summary (
                staging_sum_id {$autoInc},
                batch_id VARCHAR(64) NOT NULL,
                row_num INT NOT NULL,
                delivery_date VARCHAR(32),
                delivery_no VARCHAR(64),
                invoice_no VARCHAR(64),
                customer VARCHAR(100),
                qty_part INT,
                part_scan INT,
                jml_bag INT,
                pic VARCHAR(120),
                status VARCHAR(32),
                registered_date VARCHAR(32),
                registered_time VARCHAR(32),
                detail_row_count INT DEFAULT 0,
                anomalies_json TEXT
            )
        ");

        // 2. STAGING: Sheet DETAIL
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS raw_outgoing_detail (
                staging_det_id {$autoInc},
                batch_id VARCHAR(64) NOT NULL,
                row_num INT NOT NULL,
                delivery_date VARCHAR(32),
                invoice_no VARCHAR(64),
                customer VARCHAR(100),
                part_no VARCHAR(64),
                part_name VARCHAR(150),
                qty_order INT,
                qty_scan INT,
                lot_no_raw TEXT,
                pic VARCHAR(120),
                status VARCHAR(32),
                anomalies_json TEXT
            )
        ");

        // 3. STAR SCHEMA: dim_customer
        // Migrasi otomatis jika tabel dim_customer lama (versi e-commerce) masih ada tanpa kolom customer_code
        try {
            $this->pdo->query("SELECT customer_code FROM dim_customer LIMIT 1");
        } catch (Throwable $e) {
            $this->pdo->exec("DROP TABLE IF EXISTS dim_customer");
        }

        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS dim_customer (
                customer_key {$autoInc},
                customer_code VARCHAR(64) NOT NULL UNIQUE,
                invoice_format_type VARCHAR(64) NOT NULL,
                total_invoices INT DEFAULT 0,
                total_bags INT DEFAULT 0,
                updated_at VARCHAR(32)
            )
        ");

        // 4. STAR SCHEMA: dim_part
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS dim_part (
                part_key {$autoInc},
                part_no VARCHAR(64) NOT NULL UNIQUE,
                part_name_clean VARCHAR(150) NOT NULL,
                part_category VARCHAR(64) NOT NULL,
                raw_name_sample VARCHAR(150),
                was_cleaned INT DEFAULT 0,
                updated_at VARCHAR(32)
            )
        ");

        // 5. STAR SCHEMA: dim_pic
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS dim_pic (
                pic_key {$autoInc},
                pic_name VARCHAR(120) NOT NULL UNIQUE,
                primary_pic VARCHAR(64) NOT NULL,
                secondary_pic VARCHAR(64),
                is_multi_pic INT DEFAULT 0
            )
        ");

        // 6. STAR SCHEMA: dim_invoice
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS dim_invoice (
                invoice_key {$autoInc},
                invoice_no VARCHAR(64) NOT NULL UNIQUE,
                delivery_no VARCHAR(64) NOT NULL,
                delivery_date VARCHAR(32) NOT NULL,
                registered_date VARCHAR(32) NOT NULL,
                registered_time VARCHAR(32) NOT NULL,
                lead_time_days INT NOT NULL,
                customer_key INT NOT NULL,
                customer_code VARCHAR(64) NOT NULL,
                pic_name VARCHAR(120) NOT NULL,
                jml_bag INT NOT NULL,
                summary_qty_part INT NOT NULL,
                raw_detail_rows INT NOT NULL,
                split_extra_rows INT NOT NULL,
                invoice_status VARCHAR(32) NOT NULL
            )
        ");

        // 7. STAR SCHEMA: fact_outgoing_part (270 Baris Unik Invoice + Part, Bebas Double-Counting!)
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS fact_outgoing_part (
                fact_part_id {$autoInc},
                invoice_key INT NOT NULL,
                invoice_no VARCHAR(64) NOT NULL,
                customer_key INT NOT NULL,
                customer_code VARCHAR(64) NOT NULL,
                part_key INT NOT NULL,
                part_no VARCHAR(64) NOT NULL,
                part_name_clean VARCHAR(150) NOT NULL,
                pic_key INT NOT NULL,
                pic_name VARCHAR(120) NOT NULL,
                delivery_date VARCHAR(32) NOT NULL,
                po_split_lines INT NOT NULL,
                qty_order_total INT NOT NULL,
                qty_scan_raw_sum INT NOT NULL,
                qty_scan_clean INT NOT NULL,
                overcount_saved INT NOT NULL,
                lot_count INT NOT NULL,
                lot_raw TEXT,
                had_open_status INT NOT NULL DEFAULT 0,
                final_status VARCHAR(32) NOT NULL
            )
        ");

        // 8. STAR SCHEMA: fact_outgoing_lot (345 Baris Hasil Explode 1NF dari kolom LOT NO)
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS fact_outgoing_lot (
                fact_lot_id {$autoInc},
                invoice_no VARCHAR(64) NOT NULL,
                customer_code VARCHAR(64) NOT NULL,
                part_no VARCHAR(64) NOT NULL,
                part_name_clean VARCHAR(150) NOT NULL,
                lot_full_code VARCHAR(100) NOT NULL,
                primary_lot VARCHAR(64) NOT NULL,
                nested_sub_lot VARCHAR(64),
                lot_qty INT NOT NULL,
                delivery_date VARCHAR(32) NOT NULL,
                pic_name VARCHAR(120) NOT NULL
            )
        ");

        // 9. Log Eksekusi Pipeline ETL
        $this->pdo->exec("
            CREATE TABLE IF NOT EXISTS etl_outgoing_logs (
                log_id {$autoInc},
                batch_id VARCHAR(64) NOT NULL,
                file_name VARCHAR(150) NOT NULL,
                summary_rows INT NOT NULL,
                detail_raw_rows INT NOT NULL,
                fact_part_rows INT NOT NULL,
                fact_lot_rows INT NOT NULL,
                overcount_prevented INT NOT NULL,
                part_names_cleaned INT NOT NULL,
                status_reconciled INT NOT NULL,
                duration_ms INT NOT NULL,
                executed_at VARCHAR(32) NOT NULL
            )
        ");
    }
}
