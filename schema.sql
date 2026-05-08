CREATE TABLE IF NOT EXISTS course_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    department TEXT NOT NULL,
    employee_name TEXT NOT NULL,
    target_year INTEGER NOT NULL,
    period TEXT NOT NULL,
    record_label TEXT DEFAULT '',
    tableau REAL DEFAULT 0,
    rpa REAL DEFAULT 0,
    db_engineer REAL DEFAULT 0,
    pro REAL DEFAULT 0,
    memo TEXT DEFAULT '',
    created_by TEXT DEFAULT '',
    created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_course_records_person_period
ON course_records(department, employee_name, target_year, period);

CREATE TABLE IF NOT EXISTS deleted_course_records (
    deleted_id INTEGER PRIMARY KEY AUTOINCREMENT,
    original_id INTEGER NOT NULL,
    department TEXT NOT NULL,
    employee_name TEXT NOT NULL,
    target_year INTEGER NOT NULL,
    period TEXT NOT NULL,
    record_label TEXT DEFAULT '',
    tableau REAL DEFAULT 0,
    rpa REAL DEFAULT 0,
    db_engineer REAL DEFAULT 0,
    pro REAL DEFAULT 0,
    memo TEXT DEFAULT '',
    created_by TEXT DEFAULT '',
    created_at TEXT NOT NULL,
    deleted_by TEXT DEFAULT '',
    delete_comment TEXT NOT NULL,
    deleted_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_deleted_course_records_original_id
ON deleted_course_records(original_id);

CREATE TABLE IF NOT EXISTS audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    action TEXT NOT NULL,
    record_id INTEGER,
    department TEXT,
    employee_name TEXT,
    actor TEXT,
    created_at TEXT NOT NULL,
    FOREIGN KEY(record_id) REFERENCES course_records(id)
);
