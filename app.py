from __future__ import annotations

import io
import os
import sqlite3
import threading
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

APP_TITLE = "社内コース入力・集計システム"
COURSE_COLUMNS = ["Tableau", "RPA", "DBエンジニア", "プロ"]
COURSE_TO_DB = {
    "Tableau": "tableau",
    "RPA": "rpa",
    "DBエンジニア": "db_engineer",
    "プロ": "pro",
}
DB_TO_COURSE = {v: k for k, v in COURSE_TO_DB.items()}

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.getenv("DATA_DIR", BASE_DIR / "data"))
DB_PATH = Path(os.getenv("DB_PATH", DATA_DIR / "course_entries.db"))
EXCEL_PATH = Path(os.getenv("EXCEL_PATH", DATA_DIR / "course_entries.xlsx"))
EXPORT_LOCK = threading.Lock()

PERIOD_OPTIONS = ["上期", "下期", "1Q", "2Q", "3Q", "4Q", "通年"]


# -----------------------------
# Database
# -----------------------------
def get_conn() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=10000")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db() -> None:
    with get_conn() as conn:
        conn.execute(
            """
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
            )
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_course_records_person_period
            ON course_records(department, employee_name, target_year, period)
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS audit_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                action TEXT NOT NULL,
                record_id INTEGER,
                department TEXT,
                employee_name TEXT,
                actor TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(record_id) REFERENCES course_records(id)
            )
            """
        )


def insert_records(
    department: str,
    employee_name: str,
    rows: list[dict[str, Any]],
) -> int:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    actor = employee_name

    values = []
    for row in rows:
        values.append(
            (
                department,
                employee_name,
                int(row.get("target_year", datetime.now().year)),
                clean_text(row.get("period", "")) or PERIOD_OPTIONS[0],
                row.get("record_label", ""),
                row.get("tableau", ""),
                row.get("rpa", ""),
                row.get("db_engineer", ""),
                row.get("pro", ""),
                row.get("memo", ""),
                actor,
                now,
            )
        )

    if not values:
        return 0

    with get_conn() as conn:
        conn.execute("BEGIN IMMEDIATE")
        conn.executemany(
            """
            INSERT INTO course_records (
                department,
                employee_name,
                target_year,
                period,
                record_label,
                tableau,
                rpa,
                db_engineer,
                pro,
                memo,
                created_by,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            values,
        )
        # sqlite3.executemany() can return None for lastrowid depending on Python/SQLite.
        inserted_ids = [
            row["id"]
            for row in conn.execute(
                "SELECT id FROM course_records ORDER BY id DESC LIMIT ?",
                (len(values),),
            ).fetchall()
        ]
        for record_id in inserted_ids:
            conn.execute(
                """
                INSERT INTO audit_logs(action, record_id, department, employee_name, actor, created_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                ("insert", record_id, department, employee_name, actor, now),
            )
        conn.commit()
    return len(values)


def fetch_records() -> pd.DataFrame:
    with get_conn() as conn:
        rows = conn.execute(
            """
            SELECT
                id,
                department,
                employee_name,
                target_year,
                period,
                tableau,
                rpa,
                db_engineer,
                pro,
                created_by,
                created_at
            FROM course_records
            ORDER BY created_at DESC, id DESC
            """
        ).fetchall()

    data = [dict(row) for row in rows]
    if not data:
        return pd.DataFrame(
            columns=[
                "ID",
                "部署",
                "名前",
                "年",
                "時期",
                "Tableau",
                "RPA",
                "DBエンジニア",
                "プロ",
                "登録者",
                "登録日時",
            ]
        )

    df = pd.DataFrame(data)
    df = df.rename(
        columns={
            "id": "ID",
            "department": "部署",
            "employee_name": "名前",
            "target_year": "年",
            "period": "時期",
            "tableau": "Tableau",
            "rpa": "RPA",
            "db_engineer": "DBエンジニア",
            "pro": "プロ",
            "created_by": "登録者",
            "created_at": "登録日時",
        }
    )
    return df[
        [
            "ID",
            "部署",
            "名前",
            "年",
            "時期",
            "Tableau",
            "RPA",
            "DBエンジニア",
            "プロ",
            "登録者",
            "登録日時",
        ]
    ]


def fetch_detail_records() -> pd.DataFrame:
    aggregate_df = fetch_records()
    detail_rows: list[dict[str, Any]] = []
    for _, row in aggregate_df.iterrows():
        for course in COURSE_COLUMNS:
            value = clean_text(row.get(course, ""))
            if value:
                detail_rows.append(
                    {
                        "ID": row["ID"],
                        "部署": row["部署"],
                        "名前": row["名前"],
                        "年": row["年"],
                        "時期": row["時期"],
                        "コース": course,
                        "入力された情報": value,
                        "登録日時": row["登録日時"],
                    }
                )
    return pd.DataFrame(
        detail_rows,
        columns=["ID", "部署", "名前", "年", "時期", "コース", "入力された情報", "登録日時"],
    )


# -----------------------------
# Excel export
# -----------------------------
def sync_excel_from_sqlite() -> Path:
    """Recreate the Excel workbook from SQLite data and save it to EXCEL_PATH."""
    with EXPORT_LOCK:
        EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
        aggregate_df = fetch_records()
        detail_df = fetch_detail_records()

        wb = Workbook()
        ws = wb.active
        ws.title = "集計テーブル"
        write_dataframe(ws, aggregate_df)
        style_sheet_as_table(ws, "AggregateTable")

        detail_ws = wb.create_sheet("コース別明細")
        write_dataframe(detail_ws, detail_df)
        style_sheet_as_table(detail_ws, "DetailTable")

        summary_ws = wb.create_sheet("サマリー")
        write_summary(summary_ws, aggregate_df, detail_df)

        wb.save(EXCEL_PATH)
        return EXCEL_PATH


def write_dataframe(ws, df: pd.DataFrame) -> None:
    headers = list(df.columns)
    ws.append(headers)
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))


def style_sheet_as_table(ws, table_name: str) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    max_row = max(ws.max_row, 1)
    max_col = max(ws.max_column, 1)

    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"

    widths = {
        "A": 8,
        "B": 18,
        "C": 18,
        "D": 8,
        "E": 10,
        "F": 26,
        "G": 26,
        "H": 28,
        "I": 26,
        "J": 16,
        "K": 20,
    }
    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = widths.get(col_letter, 18)

    if max_row >= 2:
        ref = f"A1:{get_column_letter(max_col)}{max_row}"
        table = Table(displayName=table_name, ref=ref)
        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        table.tableStyleInfo = style
        ws.add_table(table)
    else:
        ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}1"


def write_summary(ws, aggregate_df: pd.DataFrame, detail_df: pd.DataFrame) -> None:
    ws["A1"] = "サマリー"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A3"] = "総登録件数"
    ws["B3"] = int(len(aggregate_df))
    ws["A4"] = "コース別明細件数"
    ws["B4"] = int(len(detail_df))
    ws["A5"] = "最終Excel更新日時"
    ws["B5"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ws["A7"] = "コース"
    ws["B7"] = "入力あり件数"
    row_no = 8
    for course in COURSE_COLUMNS:
        count = 0
        if not aggregate_df.empty:
            count = aggregate_df[course].fillna("").astype(str).str.strip().ne("").sum()
        ws.cell(row=row_no, column=1, value=course)
        ws.cell(row=row_no, column=2, value=int(count))
        row_no += 1

    ws["D7"] = "部署"
    ws["E7"] = "登録件数"
    dept_row = 8
    if not aggregate_df.empty:
        dept_counts = aggregate_df.groupby("部署", dropna=False).size().reset_index(name="登録件数")
        for _, row in dept_counts.iterrows():
            ws.cell(row=dept_row, column=4, value=row["部署"])
            ws.cell(row=dept_row, column=5, value=int(row["登録件数"]))
            dept_row += 1

    for cell in ws[7]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")

    thin = Side(style="thin", color="D9E2F3")
    for row in ws.iter_rows(min_row=1, max_row=max(ws.max_row, 10), max_col=5):
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for col, width in {"A": 20, "B": 18, "D": 24, "E": 14}.items():
        ws.column_dimensions[col].width = width


# -----------------------------
# Streamlit UI helpers
# -----------------------------
def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip()


def make_default_input_df(default_year: int, default_period: str) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"年": default_year, "時期": default_period, "Tableau": "", "RPA": "", "DBエンジニア": "", "プロ": ""},
            {"年": default_year, "時期": default_period, "Tableau": "", "RPA": "", "DBエンジニア": "", "プロ": ""},
            {"年": default_year, "時期": default_period, "Tableau": "", "RPA": "", "DBエンジニア": "", "プロ": ""},
        ]
    )


def normalize_input_rows(input_df: pd.DataFrame) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for idx, row in input_df.iterrows():
        values = {COURSE_TO_DB[course]: clean_text(row.get(course, "")) for course in COURSE_COLUMNS}
        target_year = int(row.get("年", datetime.now().year) or datetime.now().year)
        period = clean_text(row.get("時期", "")) or PERIOD_OPTIONS[0]

        has_any_course_value = any(values.values())
        if not has_any_course_value:
            # コース入力がない行は登録しない。
            continue

        rows.append(
            {
                "target_year": target_year,
                "period": period,
                "record_label": f"明細{idx + 1}",
                **values,
                "memo": "",
            }
        )
    return rows


def apply_filters(df: pd.DataFrame, department: str, employee_name: str, target_year: str, period: str) -> pd.DataFrame:
    result = df.copy()
    if department:
        result = result[result["部署"].astype(str).str.contains(department, case=False, na=False)]
    if employee_name:
        result = result[result["名前"].astype(str).str.contains(employee_name, case=False, na=False)]
    if target_year:
        result = result[result["年"].astype(str) == target_year]
    if period:
        result = result[result["時期"].astype(str) == period]
    return result


def get_excel_bytes() -> bytes | None:
    if not EXCEL_PATH.exists():
        return None
    return EXCEL_PATH.read_bytes()


def page_input() -> None:
    st.subheader("入力")
    st.caption("1人につき複数件を登録できます。各行が1件、年・時期・各コースをテーブル内で入力します。")

    with st.form("entry_form", clear_on_submit=False):
        col1, col2 = st.columns(2)
        with col1:
            department = st.text_input("部署", placeholder="例：IT / 人事 / 生産")
        with col2:
            employee_name = st.text_input("名前", placeholder="例：山田 太郎")

        st.markdown("#### コース別入力テーブル")
        edited_df = st.data_editor(
            make_default_input_df(datetime.now().year, PERIOD_OPTIONS[0]),
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "年": st.column_config.NumberColumn(
                    "年",
                    min_value=2000,
                    max_value=2100,
                    step=1,
                    format="%d",
                ),
                "時期": st.column_config.SelectboxColumn(
                    "時期",
                    options=PERIOD_OPTIONS,
                ),
                "Tableau": st.column_config.TextColumn("Tableau"),
                "RPA": st.column_config.TextColumn("RPA"),
                "DBエンジニア": st.column_config.TextColumn("DBエンジニア"),
                "プロ": st.column_config.TextColumn("プロ"),
            },
        )

        submitted = st.form_submit_button("確定してSQLiteへ登録し、Excelへ保存", type="primary")

    if submitted:
        department = department.strip()
        employee_name = employee_name.strip()
        if not department or not employee_name:
            st.error("部署と名前を入力してください。")
            return

        rows = normalize_input_rows(edited_df)
        if not rows:
            st.warning("登録対象の行がありません。Tableau / RPA / DBエンジニア / プロ のいずれかに入力してください。")
            return

        saved_count = insert_records(department, employee_name, rows)
        excel_path = sync_excel_from_sqlite()
        st.success(f"{saved_count}件を登録しました。Excelにも保存しました: {excel_path}")
        st.info("同じ人でも、次回の確定で別件として追加登録されます。")


def page_summary() -> None:
    st.subheader("集計テーブル")
    aggregate_df = fetch_records()

    with st.expander("絞り込み", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            department = st.text_input("部署", key="filter_department")
        with col2:
            employee_name = st.text_input("名前", key="filter_name")
        with col3:
            year_options = [""] + sorted([str(x) for x in aggregate_df["年"].dropna().unique()], reverse=True) if not aggregate_df.empty else [""]
            target_year = st.selectbox("年", year_options, key="filter_year")
        with col4:
            period = st.selectbox("時期", [""] + PERIOD_OPTIONS, key="filter_period")

    filtered_df = apply_filters(aggregate_df, department, employee_name, target_year, period)
    detail_df = fetch_detail_records()

    c1, c2, c3 = st.columns(3)
    c1.metric("表示中の登録件数", len(filtered_df))
    c2.metric("全登録件数", len(aggregate_df))
    c3.metric("全コース別明細件数", len(detail_df))

    st.dataframe(filtered_df, use_container_width=True, hide_index=True)

    st.markdown("#### Excel")
    if st.button("Excelを再作成", type="secondary"):
        path = sync_excel_from_sqlite()
        st.success(f"Excelを再作成しました: {path}")

    excel_bytes = get_excel_bytes()
    if excel_bytes:
        st.download_button(
            "Excelをダウンロード",
            data=excel_bytes,
            file_name="course_entries.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("まだExcelファイルが作成されていません。入力画面で確定するか、Excelを再作成してください。")


def page_detail() -> None:
    st.subheader("コース別明細")
    st.caption("Excelの「コース別明細」シートと同じ形式です。1コース1行で確認できます。")
    detail_df = fetch_detail_records()
    st.dataframe(detail_df, use_container_width=True, hide_index=True)


def page_admin() -> None:
    st.subheader("データ管理")
    st.write("保存先")
    st.code(f"SQLite: {DB_PATH}\nExcel:  {EXCEL_PATH}")

    aggregate_df = fetch_records()
    if not aggregate_df.empty:
        st.write("直近10件")
        st.dataframe(aggregate_df.head(10), use_container_width=True, hide_index=True)
    else:
        st.info("まだ登録データがありません。")

    with st.expander("危険操作: 全データ削除"):
        st.warning("検証環境用です。本番運用では権限管理を追加してください。")
        confirm = st.text_input("削除する場合は DELETE と入力")
        if st.button("SQLiteとExcelを初期化", type="secondary", disabled=(confirm != "DELETE")):
            with get_conn() as conn:
                conn.execute("BEGIN IMMEDIATE")
                conn.execute("DELETE FROM audit_logs")
                conn.execute("DELETE FROM course_records")
                conn.commit()
            sync_excel_from_sqlite()
            st.success("全データを削除し、Excelを初期化しました。")


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")
    init_db()

    st.title(APP_TITLE)
    st.caption("Streamlit + SQLite + Excel保存版")

    menu = st.sidebar.radio(
        "メニュー",
        ["入力", "集計テーブル", "コース別明細", "データ管理"],
        index=0,
    )

    if menu == "入力":
        page_input()
    elif menu == "集計テーブル":
        page_summary()
    elif menu == "コース別明細":
        page_detail()
    else:
        page_admin()


if __name__ == "__main__":
    main()
