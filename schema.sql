CREATE TABLE IF NOT EXISTS course_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    department TEXT NOT NULL,
    employee_name TEXT NOT NULL,
    target_year INTEGER NOT NULL,
    period TEXT NOT NULL,
    record_label TEXT DEFAULT '',
    tableau TEXT DEFAULT '',
    rpa TEXT DEFAULT '',
    db_engineer TEXT DEFAULT '',
    pro TEXT DEFAULT '',
    memo TEXT DEFAULT '',
    created_by TEXT DEFAULT '',
    created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_course_records_person_period
ON course_records(department, employee_name, target_year, period);

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
