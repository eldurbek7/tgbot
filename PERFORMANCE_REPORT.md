# Telegram Bot PostgreSQL Performance Optimization Report

## Scope analyzed
Project files in the ZIP:

- `main.py` — main Aiogram bot, handlers, keyboards, reports, exports, database logic.
- `sqlite_to_postgres.py` — SQLite to PostgreSQL migration utility.
- `requirements.txt` — dependencies.
- `railway.json` — Railway start command.

## Database-related files

- `main.py` before optimization contained all PostgreSQL code inline.
- `sqlite_to_postgres.py` used psycopg2 for migration.
- New `db.py` contains the global asyncpg pool and reusable helpers.
- New `migrations_indexes.sql` contains safe idempotent PostgreSQL indexes.

## Main bottlenecks found

1. **New PostgreSQL connection per DB request**
   - `psycopg2.connect(DATABASE_URL)` was called from `get_connection()`, `execute_query_plain()`, `init_db()`, `_sync_subjects_to_db()`, `migrate_old_subject_keys()`, and `save_vote()`.
   - This is the biggest reason buttons could take 800–3000 ms on Railway.

2. **Blocking psycopg2 inside async Aiogram handlers**
   - Handlers were `async`, but DB operations were synchronous.
   - This blocked the event loop during callback processing.

3. **Repeated settings/user preference queries**
   - `get_user_script()`, `tr()`, labels, keyboards, and text builders repeatedly queried `user_prefs`.
   - `is_voting_open()` repeatedly queried `settings`.

4. **N+1 query patterns**
   - `rating_rows()` queried rating counts once per teacher.
   - `get_top_votes_text()` queried vote count once per teacher.
   - `get_top_votes_by_subject_text()` queried vote count inside nested subject/teacher loops.

5. **Repeated subject/teacher lookups**
   - `get_subject_name()`, `get_teacher_name()`, and student counts queried DB frequently even though these values change rarely.

6. **Missing indexes for callback-heavy queries**
   - Vote checks by `user_id`, ranking by `subject_key/teacher_key`, complaint filters, rating stats, and ordered exports had no supporting indexes.

7. **Synchronous exports and report generation**
   - Excel/DOCX/ZIP generation remains file/CPU-heavy. These actions are admin-only and not on normal vote buttons; they are isolated from most user button paths.

## Fixes applied

### 1. psycopg2 removed

- Removed `psycopg2-binary` from `requirements.txt`.
- Replaced with `asyncpg>=0.29.0`.
- `main.py` no longer imports or calls psycopg2.
- `sqlite_to_postgres.py` was rewritten to use asyncpg.

### 2. Global asyncpg pool added

New `db.py` implements:

- `init_pool(DATABASE_URL, min_size=5, max_size=20, command_timeout=30)`
- `fetch(...)`
- `fetchrow(...)`
- `fetchval(...)`
- `execute(...)`
- `executemany(...)`
- sync compatibility wrappers backed by the same pool to preserve all existing business logic safely.

The pool is initialized once in `on_startup()` and closed on shutdown.

### 3. Startup optimization

`on_startup()` now initializes:

- asyncpg pool
- DB schema
- indexes
- settings cache
- subject/teacher cache
- auto voting scheduler

### 4. Index migration added

New file: `migrations_indexes.sql`

Indexes include:

- `idx_users_user_id` on `user_prefs(user_id)`
- `idx_votes_user_id`
- `idx_votes_subject_key`
- `idx_votes_teacher_key`
- `idx_votes_subject_teacher`
- `idx_teacher_ratings_user_id`
- `idx_teacher_ratings_subject_teacher`
- `idx_complaints_user_id`
- `idx_complaints_type_id`
- plus date/order indexes for exports and recent lists.

### 5. TTL caches added

Installed `cachetools>=5.3.3`.

Caches:

- `SETTINGS_CACHE` — voting status and schedule settings.
- `USER_PREFS_CACHE` — user language/access preferences.
- `SUBJECTS_CACHE` — departments, teachers, student counts.
- `STATS_CACHE` — vote count maps, total vote counts, rating rows.

TTL defaults:

- Settings: 60 seconds.
- Subject data: 60 seconds.
- Stats: 60 seconds.
- User prefs: 300 seconds.

Write operations invalidate affected caches immediately.

### 6. N+1 query reductions

- `rating_rows()` now uses one grouped query instead of one query per teacher.
- `get_top_votes_text()` now uses one grouped vote query.
- `get_top_votes_by_subject_text()` now uses one grouped vote query.
- Subject, teacher, and student count reads now use cached data.

### 7. Performance logging middleware

Added callback middleware logging:

```text
BUTTON=top10 SQL_TIME=0.0200s TOTAL_TIME=0.0800s
```

This makes Railway logs useful for measuring exact button bottlenecks.

## Modified files

- `main.py`
- `sqlite_to_postgres.py`
- `requirements.txt`

## New files

- `db.py`
- `migrations_indexes.sql`
- `PERFORMANCE_REPORT.md`

## Expected response-time improvement

Before:

- Every DB operation could pay connection setup cost.
- Button callbacks were commonly 800–3000 ms.

After:

- Reuses one PostgreSQL pool.
- Common settings/subject/stats paths hit memory cache.
- Ranking callbacks use grouped queries instead of per-teacher loops.

Expected typical button latency on Railway after warm startup:

- Cached buttons: ~50–150 ms.
- Simple DB buttons: ~80–250 ms depending on Railway/PostgreSQL region.
- Heavy admin exports: still slower because Excel/DOCX/ZIP generation is inherently file/CPU-heavy.

## Notes

- All original handlers, commands, callback data, FSM-like in-memory states, admin panel, complaints, suggestions, settings, scheduled voting, exports, top results, and department results were preserved.
- The code compiles with `python -m py_compile main.py db.py sqlite_to_postgres.py`.
- For maximum future improvement, the remaining legacy helper calls can be gradually converted from sync compatibility wrappers to native `await fetchrow(...)` / `await execute(...)` without changing business logic.
