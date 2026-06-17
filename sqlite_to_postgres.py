"""
SQLite bazadan PostgreSQL ga asyncpg orqali ko'chirish skripti.

Ishlatish:
  python sqlite_to_postgres.py --sqlite bot.db --postgres "postgresql://user:pass@host/dbname"

Yoki muhit o'zgaruvchilari orqali:
  SQLITE_PATH=bot.db DATABASE_URL=postgresql://... python sqlite_to_postgres.py
"""

from __future__ import annotations

import argparse
import asyncio
import logging
import os
import sqlite3
import sys
from typing import Iterable

import asyncpg

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

DDL = """
CREATE TABLE IF NOT EXISTS votes (
    user_id BIGINT PRIMARY KEY,
    full_name TEXT,
    username TEXT,
    subject_key TEXT NOT NULL,
    teacher_key TEXT NOT NULL,
    voted_at TEXT
);

CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
);

CREATE TABLE IF NOT EXISTS user_prefs (
    user_id BIGINT PRIMARY KEY,
    script TEXT DEFAULT 'latin',
    access_granted INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS teacher_ratings (
    user_id BIGINT NOT NULL,
    full_name TEXT,
    username TEXT,
    subject_key TEXT NOT NULL,
    teacher_key TEXT NOT NULL,
    rating TEXT NOT NULL,
    rated_at TEXT,
    PRIMARY KEY (user_id, subject_key, teacher_key)
);

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
);

CREATE TABLE IF NOT EXISTS db_subjects (
    subject_key TEXT PRIMARY KEY,
    subject_name TEXT NOT NULL,
    sort_order INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS db_teachers (
    teacher_key TEXT NOT NULL,
    subject_key TEXT NOT NULL,
    teacher_name TEXT NOT NULL,
    student_count INTEGER DEFAULT 0,
    PRIMARY KEY (subject_key, teacher_key)
);
"""


def get_args():
    parser = argparse.ArgumentParser(description="SQLite -> PostgreSQL migratsiya")
    parser.add_argument("--sqlite", default=os.getenv("SQLITE_PATH", "bot.db"), help="SQLite fayl yo'li")
    parser.add_argument("--postgres", default=os.getenv("DATABASE_URL", ""), help="PostgreSQL DATABASE_URL")
    return parser.parse_args()


def connect_sqlite(path: str):
    if not os.path.exists(path):
        logging.error("SQLite fayl topilmadi: %s", path)
        sys.exit(1)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    logging.info("SQLite ulandi: %s", path)
    return conn


def pg_placeholders(count: int) -> str:
    return ", ".join(f"${i}" for i in range(1, count + 1))


def get_sqlite_tables(sq_conn) -> list[str]:
    cur = sq_conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    return [row[0] for row in cur.fetchall()]


async def ensure_pg_tables(pg: asyncpg.Connection) -> None:
    await pg.execute(DDL)
    migrations_path = os.path.join(os.path.dirname(__file__), "migrations_indexes.sql")
    if os.path.exists(migrations_path):
        with open(migrations_path, "r", encoding="utf-8") as f:
            await pg.execute(f.read())
    logging.info("PostgreSQL jadvallar va indekslar tayyor.")


async def migrate_table(sq_conn, pg: asyncpg.Connection, table: str) -> int:
    sq_cur = sq_conn.cursor()
    sq_cur.execute(f"PRAGMA table_info({table})")
    cols_info = sq_cur.fetchall()
    if not cols_info:
        logging.warning("  Jadval topilmadi yoki bo'sh: %s", table)
        return 0

    columns = [row["name"] for row in cols_info]
    col_list = ", ".join(columns)
    placeholders = pg_placeholders(len(columns))

    sq_cur.execute(f"SELECT * FROM {table}")
    rows = sq_cur.fetchall()
    if not rows:
        logging.info("  %s: 0 ta yozuv (bo'sh)", table)
        return 0

    values = [tuple(row[col] for col in columns) for row in rows]
    await pg.executemany(
        f"INSERT INTO {table} ({col_list}) VALUES ({placeholders}) ON CONFLICT DO NOTHING",
        values,
    )
    logging.info("  %s: %s ta yozuv ko'chirildi.", table, len(values))
    return len(values)


async def migrate_complaints(sq_conn, pg: asyncpg.Connection) -> int:
    sq_cur = sq_conn.cursor()
    sq_cur.execute("PRAGMA table_info(complaints)")
    cols_info = sq_cur.fetchall()
    if not cols_info:
        logging.info("  complaints: SQLite da topilmadi.")
        return 0

    cols = [row["name"] for row in cols_info if row["name"].lower() != "id"]
    if not cols:
        return 0
    col_list = ", ".join(cols)
    placeholders = pg_placeholders(len(cols))

    sq_cur.execute("SELECT * FROM complaints ORDER BY id ASC")
    rows = sq_cur.fetchall()
    if not rows:
        logging.info("  complaints: 0 ta yozuv (bo'sh)")
        return 0

    values = [tuple(row[col] for col in cols) for row in rows]
    await pg.executemany(f"INSERT INTO complaints ({col_list}) VALUES ({placeholders})", values)
    logging.info("  complaints: %s ta yozuv ko'chirildi.", len(values))
    return len(values)


async def print_summary(sq_conn, pg: asyncpg.Connection, tables: Iterable[str]) -> None:
    print("\n" + "=" * 55)
    print("  MIGRATSIYA NATIJASI")
    print("=" * 55)
    sq_cur = sq_conn.cursor()
    for table in tables:
        try:
            sq_cur.execute(f"SELECT COUNT(*) FROM {table}")
            sq_count = sq_cur.fetchone()[0]
        except Exception:
            sq_count = "—"
        try:
            pg_count = await pg.fetchval(f"SELECT COUNT(*) FROM {table}")
        except Exception:
            pg_count = "—"
        status = "✅" if sq_count == pg_count else "⚠️"
        print(f"  {status}  {table:<30} SQLite: {sq_count:<8} PG: {pg_count}")
    print("=" * 55)


async def async_main() -> None:
    args = get_args()
    if not args.postgres:
        logging.error("DATABASE_URL bo'sh. --postgres parametri yoki DATABASE_URL o'zgaruvchisini bering.")
        sys.exit(1)

    sq_conn = connect_sqlite(args.sqlite)
    pg = await asyncpg.connect(args.postgres)
    logging.info("PostgreSQL ulandi.")

    try:
        await ensure_pg_tables(pg)
        tables = get_sqlite_tables(sq_conn)
        logging.info("SQLite jadvallar: %s", ", ".join(tables) if tables else "yo'q")

        migration_order = ["settings", "user_prefs", "db_subjects", "db_teachers", "votes", "teacher_ratings"]
        migrated_tables: list[str] = []
        async with pg.transaction():
            for table in migration_order:
                if table in tables:
                    await migrate_table(sq_conn, pg, table)
                    migrated_tables.append(table)
            if "complaints" in tables:
                await migrate_complaints(sq_conn, pg)
                migrated_tables.append("complaints")

        await print_summary(sq_conn, pg, migrated_tables)
        logging.info("Migratsiya tugadi.")
    finally:
        sq_conn.close()
        await pg.close()


if __name__ == "__main__":
    asyncio.run(async_main())
