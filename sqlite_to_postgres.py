"""SQLite votes.db faylini Railway PostgreSQL bazasiga ko'chirish scripti.

Ishlatish:
  export DATABASE_URL="postgresql://..."
  python sqlite_to_postgres.py /path/to/votes.db

Agar path berilmasa, DATA_DIR/votes.db yoki /app/data/votes.db ishlatiladi.
"""
import os
import sys
import sqlite3
from pathlib import Path

import psycopg2
import psycopg2.extras

DATABASE_URL = os.getenv("DATABASE_URL")
if not DATABASE_URL:
    raise SystemExit("DATABASE_URL topilmadi. Avval Railway PostgreSQL DATABASE_URL ni muhit o'zgaruvchisiga qo'ying.")

DATA_DIR = os.getenv("DATA_DIR", "/app/data")
SQLITE_PATH = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(DATA_DIR) / "votes.db"
if not SQLITE_PATH.exists():
    raise SystemExit(f"SQLite fayl topilmadi: {SQLITE_PATH}")

TABLES = [
    "settings",
    "user_prefs",
    "db_subjects",
    "db_teachers",
    "votes",
    "teacher_ratings",
    "complaints",
]

SCHEMA_SQL = [
    """
    CREATE TABLE IF NOT EXISTS votes (
        user_id BIGINT PRIMARY KEY,
        full_name TEXT,
        username TEXT,
        subject_key TEXT NOT NULL,
        teacher_key TEXT NOT NULL,
        voted_at TEXT
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS user_prefs (
        user_id BIGINT PRIMARY KEY,
        script TEXT DEFAULT 'latin',
        access_granted INTEGER DEFAULT 0
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS teacher_ratings (
        user_id BIGINT NOT NULL,
        full_name TEXT,
        username TEXT,
        subject_key TEXT NOT NULL,
        teacher_key TEXT NOT NULL,
        rating TEXT NOT NULL CHECK(rating IN ('like', 'dislike')),
        rated_at TEXT,
        PRIMARY KEY (user_id, subject_key, teacher_key)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS complaints (
        id SERIAL PRIMARY KEY,
        user_id BIGINT NOT NULL,
        full_name TEXT,
        username TEXT,
        type TEXT DEFAULT 'general',
        subject_key TEXT,
        teacher_key TEXT,
        message_text TEXT NOT NULL,
        created_at TEXT
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS db_subjects (
        subject_key TEXT PRIMARY KEY,
        subject_name TEXT NOT NULL,
        sort_order INTEGER DEFAULT 0
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS db_teachers (
        teacher_key TEXT NOT NULL,
        subject_key TEXT NOT NULL,
        teacher_name TEXT NOT NULL,
        student_count INTEGER DEFAULT 0,
        PRIMARY KEY (subject_key, teacher_key)
    )
    """,
]

CONFLICTS = {
    "settings": "key",
    "user_prefs": "user_id",
    "db_subjects": "subject_key",
    "db_teachers": "subject_key, teacher_key",
    "votes": "user_id",
    "teacher_ratings": "user_id, subject_key, teacher_key",
    "complaints": "id",
}


def sqlite_table_exists(conn, table: str) -> bool:
    cur = conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,))
    return cur.fetchone() is not None


def get_columns(conn, table: str):
    return [row[1] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()]


def copy_table(sqlite_conn, pg_conn, table: str):
    if not sqlite_table_exists(sqlite_conn, table):
        print(f"SKIP: {table} SQLite'da yo'q")
        return

    columns = get_columns(sqlite_conn, table)
    if not columns:
        print(f"SKIP: {table} ustunlari topilmadi")
        return

    rows = sqlite_conn.execute(f"SELECT {', '.join(columns)} FROM {table}").fetchall()
    if not rows:
        print(f"OK: {table} bo'sh")
        return

    placeholders = ", ".join(["%s"] * len(columns))
    col_sql = ", ".join(columns)
    conflict = CONFLICTS.get(table)
    if table == "complaints":
        update_cols = [c for c in columns if c != "id"]
    else:
        update_cols = [c for c in columns if c not in conflict.replace(" ", "").split(",")]

    if conflict and update_cols:
        set_sql = ", ".join([f"{c}=EXCLUDED.{c}" for c in update_cols])
        sql = f"INSERT INTO {table} ({col_sql}) VALUES ({placeholders}) ON CONFLICT ({conflict}) DO UPDATE SET {set_sql}"
    elif conflict:
        sql = f"INSERT INTO {table} ({col_sql}) VALUES ({placeholders}) ON CONFLICT ({conflict}) DO NOTHING"
    else:
        sql = f"INSERT INTO {table} ({col_sql}) VALUES ({placeholders})"

    with pg_conn.cursor() as cur:
        psycopg2.extras.execute_batch(cur, sql, rows, page_size=200)
    pg_conn.commit()
    print(f"OK: {table} -> {len(rows)} qator ko'chirildi")


def main():
    sqlite_conn = sqlite3.connect(str(SQLITE_PATH))
    pg_conn = psycopg2.connect(DATABASE_URL, sslmode=os.getenv("PGSSLMODE", "require"))

    try:
        with pg_conn.cursor() as cur:
            for sql in SCHEMA_SQL:
                cur.execute(sql)
        pg_conn.commit()

        for table in TABLES:
            copy_table(sqlite_conn, pg_conn, table)

        with pg_conn.cursor() as cur:
            cur.execute("SELECT setval(pg_get_serial_sequence('complaints','id'), COALESCE((SELECT MAX(id) FROM complaints), 1), true)")
        pg_conn.commit()
        print("Tayyor: SQLite ma'lumotlari PostgreSQL'ga ko'chirildi.")
    finally:
        sqlite_conn.close()
        pg_conn.close()


if __name__ == "__main__":
    main()
