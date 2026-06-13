"""
SQLite bazadan PostgreSQL ga ko'chirish skripti.

Ishlatish:
  python sqlite_to_postgres.py --sqlite bot.db --postgres "postgresql://user:pass@host/dbname"

Yoki muhit o'zgaruvchilari orqali:
  SQLITE_PATH=bot.db DATABASE_URL=postgresql://... python sqlite_to_postgres.py
"""

import os
import sys
import argparse
import logging

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

def get_args():
    parser = argparse.ArgumentParser(description="SQLite -> PostgreSQL migratsiya")
    parser.add_argument("--sqlite", default=os.getenv("SQLITE_PATH", "bot.db"), help="SQLite fayl yo'li")
    parser.add_argument("--postgres", default=os.getenv("DATABASE_URL", ""), help="PostgreSQL DATABASE_URL")
    return parser.parse_args()


def check_deps():
    try:
        import sqlite3
    except ImportError:
        logging.error("sqlite3 moduli topilmadi (Python standart kutubxonasi).")
        sys.exit(1)
    try:
        import psycopg2
    except ImportError:
        logging.error("psycopg2 topilmadi. 'pip install psycopg2-binary' bajaring.")
        sys.exit(1)


def connect_sqlite(path: str):
    import sqlite3
    if not os.path.exists(path):
        logging.error(f"SQLite fayl topilmadi: {path}")
        sys.exit(1)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    logging.info(f"SQLite ulandi: {path}")
    return conn


def connect_pg(url: str):
    import psycopg2
    if not url:
        logging.error("DATABASE_URL bo'sh. --postgres parametri yoki DATABASE_URL o'zgaruvchisini bering.")
        sys.exit(1)
    conn = psycopg2.connect(url)
    logging.info("PostgreSQL ulandi.")
    return conn


def ensure_pg_tables(pg_conn):
    """PostgreSQL jadvallarini yaratadi (agar mavjud bo'lmasa)."""
    ddl = """
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
    with pg_conn.cursor() as cur:
        cur.execute(ddl)
    pg_conn.commit()
    logging.info("PostgreSQL jadvallar tayyor.")


def get_sqlite_tables(sq_conn):
    cur = sq_conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    return [row[0] for row in cur.fetchall()]


def migrate_table(sq_conn, pg_conn, table: str, conflict_action: str = "DO NOTHING"):
    """
    Bir jadval ma'lumotlarini ko'chiradi.
    conflict_action: 'DO NOTHING' yoki 'DO UPDATE SET ...'
    """
    sq_cur = sq_conn.cursor()

    # Ustunlarni olish
    sq_cur.execute(f"PRAGMA table_info({table})")
    cols_info = sq_cur.fetchall()
    if not cols_info:
        logging.warning(f"  Jadval topilmadi yoki bo'sh: {table}")
        return 0

    columns = [row["name"] for row in cols_info]
    col_list = ", ".join(columns)
    placeholders = ", ".join(["%s"] * len(columns))

    sq_cur.execute(f"SELECT * FROM {table}")
    rows = sq_cur.fetchall()

    if not rows:
        logging.info(f"  {table}: 0 ta yozuv (bo'sh)")
        return 0

    pg_cur = pg_conn.cursor()
    inserted = 0
    skipped = 0

    for row in rows:
        values = tuple(row[col] for col in columns)
        try:
            if conflict_action == "DO NOTHING":
                pg_cur.execute(
                    f"INSERT INTO {table} ({col_list}) VALUES ({placeholders}) ON CONFLICT DO NOTHING",
                    values
                )
            else:
                pg_cur.execute(
                    f"INSERT INTO {table} ({col_list}) VALUES ({placeholders}) ON CONFLICT DO NOTHING",
                    values
                )
            inserted += 1
        except Exception as e:
            pg_conn.rollback()
            logging.warning(f"  Yozuv o'tkazilmadi ({table}): {e} | Qiymatlar: {values}")
            skipped += 1
            continue

    pg_conn.commit()
    logging.info(f"  {table}: {inserted} ta yozuv ko'chirildi, {skipped} ta o'tkazib yuborildi.")
    return inserted


def migrate_complaints(sq_conn, pg_conn):
    """
    complaints jadvalini ID bilan emas, mazmun bo'yicha ko'chiradi
    (chunki PostgreSQL da SERIAL id bor).
    """
    table = "complaints"
    sq_cur = sq_conn.cursor()
    sq_cur.execute("PRAGMA table_info(complaints)")
    cols_info = sq_cur.fetchall()
    if not cols_info:
        logging.info("  complaints: SQLite da topilmadi.")
        return 0

    all_cols = [row["name"] for row in cols_info]

    # 'id' ustunini olib tashlaymiz - PostgreSQL SERIAL o'zi beradi
    cols = [c for c in all_cols if c.lower() != "id"]
    col_list = ", ".join(cols)
    placeholders = ", ".join(["%s"] * len(cols))

    sq_cur.execute("SELECT * FROM complaints ORDER BY id ASC")
    rows = sq_cur.fetchall()

    if not rows:
        logging.info("  complaints: 0 ta yozuv (bo'sh)")
        return 0

    pg_cur = pg_conn.cursor()
    inserted = 0
    skipped = 0

    for row in rows:
        values = tuple(row[col] for col in cols)
        try:
            pg_cur.execute(
                f"INSERT INTO {table} ({col_list}) VALUES ({placeholders})",
                values
            )
            inserted += 1
        except Exception as e:
            pg_conn.rollback()
            logging.warning(f"  complaints yozuv o'tkazilmadi: {e}")
            skipped += 1
            continue

    pg_conn.commit()
    logging.info(f"  complaints: {inserted} ta yozuv ko'chirildi, {skipped} ta o'tkazib yuborildi.")
    return inserted


def print_summary(sq_conn, pg_conn, tables):
    """Ko'chirishdan keyin hisobot chiqaradi."""
    print("\n" + "=" * 55)
    print("  MIGRATSIYA NATIJASI")
    print("=" * 55)
    sq_cur = sq_conn.cursor()
    pg_cur = pg_conn.cursor()

    for table in tables:
        try:
            sq_cur.execute(f"SELECT COUNT(*) FROM {table}")
            sq_count = sq_cur.fetchone()[0]
        except Exception:
            sq_count = "—"

        try:
            pg_cur.execute(f"SELECT COUNT(*) FROM {table}")
            pg_count = pg_cur.fetchone()[0]
        except Exception:
            pg_count = "—"

        status = "✅" if sq_count == pg_count else "⚠️"
        print(f"  {status}  {table:<30} SQLite: {sq_count:<8} PG: {pg_count}")

    print("=" * 55)
    print()


def main():
    check_deps()
    args = get_args()

    logging.info("=" * 50)
    logging.info("SQLite -> PostgreSQL migratsiya boshlandi")
    logging.info("=" * 50)

    sq_conn = connect_sqlite(args.sqlite)
    pg_conn = connect_pg(args.postgres)

    # Jadvallarni yaratish
    ensure_pg_tables(pg_conn)

    sq_tables = get_sqlite_tables(sq_conn)
    logging.info(f"SQLite jadvallar: {sq_tables}")

    # Ko'chirish tartibi (bog'liqliklar bo'yicha)
    migration_order = [
        "settings",
        "user_prefs",
        "votes",
        "teacher_ratings",
        "db_subjects",
        "db_teachers",
    ]

    for table in migration_order:
        if table not in sq_tables:
            logging.info(f"  {table}: SQLite da yo'q, o'tkazildi.")
            continue
        logging.info(f"Ko'chirilmoqda: {table}")
        if table == "complaints":
            continue  # alohida ko'chiramiz
        migrate_table(sq_conn, pg_conn, table)

    # complaints alohida (chunki id SERIAL)
    if "complaints" in sq_tables:
        logging.info("Ko'chirilmoqda: complaints")
        migrate_complaints(sq_conn, pg_conn)

    # Natija hisoboti
    all_tables = list(set(migration_order + ["complaints"]))
    print_summary(sq_conn, pg_conn, [t for t in all_tables if t in sq_tables])

    sq_conn.close()
    pg_conn.close()
    logging.info("Migratsiya yakunlandi!")


if __name__ == "__main__":
    main()
