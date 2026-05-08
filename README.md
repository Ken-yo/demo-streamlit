# 社内コース入力・集計システム

Streamlit + SQLite + Excel保存版です。

複数人がブラウザからアクセスし、部署・名前・年・時期を指定して、以下のコース項目に入力できます。

- Tableau
- RPA
- DBエンジニア
- プロ

## 集計単位

**1人複数件**です。

入力画面では、1行が1件です。  
同じ人が何度「確定」しても、既存データの上書きではなく、別件として追加登録されます。

## 保存先

登録時に以下へ保存します。

1. SQLite
   - `data/course_entries.db`
2. Excel
   - `data/course_entries.xlsx`

Excelは登録のたびにSQLiteの内容から再作成されます。

## Excelシート構成

### 集計テーブル

1行が1件です。

| ID | 部署 | 名前 | 年 | 時期 | 件名/No | Tableau | RPA | DBエンジニア | プロ | メモ | 登録者 | 登録日時 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|

### コース別明細

1コース1行です。コース別に集計・フィルターしやすい形式です。

| ID | 部署 | 名前 | 年 | 時期 | 件名/No | コース | 入力された情報 | メモ | 登録日時 |
|---|---|---|---|---|---|---|---|---|---|

### サマリー

- 総登録件数
- コース別入力あり件数
- 部署別登録件数
- 最終Excel更新日時

## セットアップ

```bash
cd streamlit_sqlite_excel_course_system
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
```

macOS / Linux:

```bash
source .venv/bin/activate
```

依存ライブラリをインストールします。

```bash
pip install -r requirements.txt
```

## 起動

```bash
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

起動PCで開く場合:

```text
http://localhost:8501
```

社内LAN内の別PCからアクセスする場合:

```text
http://<起動PCのIPアドレス>:8501
```

## 同時アクセスについて

SQLiteはWALモード、busy_timeoutを設定しています。  
数人程度の同時入力であれば検証可能です。

ただし、全社・大人数利用に広げる場合は、SQLiteではなく PostgreSQL / SQL Server を推奨します。

## 運用時に追加を推奨する機能

- 社内ログイン認証
- 部署マスタ
- ユーザーマスタ
- 入力期間の開閉管理
- 確定取消・承認フロー
- Excelバックアップ世代管理
- 権限別メニュー制御

## ファイル構成

```text
streamlit_sqlite_excel_course_system/
├─ app.py
├─ requirements.txt
├─ schema.sql
├─ README.md
├─ run_windows.bat
├─ run_mac_linux.sh
├─ .streamlit/
│  └─ config.toml
└─ data/
   └─ .gitkeep
```
