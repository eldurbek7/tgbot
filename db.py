"""Async PostgreSQL access layer for the Telegram bot.

Uses one global asyncpg connection pool initialized at startup.  The async
helpers are preferred from handlers; sync wrappers are kept for legacy helper
functions during the migration and still reuse the same pool.
"""

from __future__ import annotations

import asyncio
import contextvars
import logging
import re
import threading
import time
from typing import Any, Iterable, Optional

import asyncpg

_pool: Optional[asyncpg.Pool] = None
_pool_loop: Optional[asyncio.AbstractEventLoop] = None
_pool_ready = threading.Event()
_loop_thread: Optional[threading.Thread] = None
_init_lock = threading.Lock()
_sql_time_var: contextvars.ContextVar[float] = contextvars.ContextVar("sql_time", default=0.0)

_PLACEHOLDER_RE = re.compile(r"%s")


def _normalize_sql(sql: str) -> str:
    """Accept legacy %s placeholders and send asyncpg $1/$2 form."""
    index = 0

    def repl(_: re.Match[str]) -> str:
        nonlocal index
        index += 1
        return f"${index}"

    return _PLACEHOLDER_RE.sub(repl, sql)


def reset_request_sql_time() -> None:
    _sql_time_var.set(0.0)


def get_request_sql_time() -> float:
    return _sql_time_var.get()


def _add_sql_time(delta: float) -> None:
    _sql_time_var.set(_sql_time_var.get() + delta)


def _start_background_loop() -> asyncio.AbstractEventLoop:
    global _pool_loop, _loop_thread
    if _pool_loop and _pool_loop.is_running():
        return _pool_loop
    with _init_lock:
        if _pool_loop and _pool_loop.is_running():
            return _pool_loop
        loop = asyncio.new_event_loop()

        def run() -> None:
            asyncio.set_event_loop(loop)
            loop.run_forever()

        thread = threading.Thread(target=run, name="postgres-asyncpg-loop", daemon=True)
        thread.start()
        _pool_loop = loop
        _loop_thread = thread
        return loop


async def init_pool(database_url: str, *, min_size: int = 5, max_size: int = 20, command_timeout: int = 30) -> None:
    """Initialize the global asyncpg pool exactly once."""
    global _pool
    if _pool is not None:
        return
    loop = _start_background_loop()

    async def _init() -> None:
        global _pool
        if _pool is None:
            _pool = await asyncpg.create_pool(
                database_url,
                min_size=min_size,
                max_size=max_size,
                command_timeout=command_timeout,
                statement_cache_size=512,
            )
            logging.info("PostgreSQL asyncpg pool initialized: min=%s max=%s", min_size, max_size)
        _pool_ready.set()

    fut = asyncio.run_coroutine_threadsafe(_init(), loop)
    await asyncio.wrap_future(fut)


async def close_pool() -> None:
    global _pool
    if _pool is None:
        return
    loop = _start_background_loop()

    async def _close() -> None:
        global _pool
        if _pool is not None:
            await _pool.close()
            _pool = None
            logging.info("PostgreSQL asyncpg pool closed")

    fut = asyncio.run_coroutine_threadsafe(_close(), loop)
    await asyncio.wrap_future(fut)


async def _run_on_pool(coro_factory):
    if _pool is None:
        raise RuntimeError("PostgreSQL pool is not initialized. Call init_pool() during startup.")
    loop = _start_background_loop()
    start = time.perf_counter()
    fut = asyncio.run_coroutine_threadsafe(coro_factory(), loop)
    try:
        return await asyncio.wrap_future(fut)
    finally:
        _add_sql_time(time.perf_counter() - start)


def _run_on_pool_sync(coro_factory):
    if _pool is None:
        raise RuntimeError("PostgreSQL pool is not initialized. Call init_pool() during startup.")
    loop = _start_background_loop()
    start = time.perf_counter()
    fut = asyncio.run_coroutine_threadsafe(coro_factory(), loop)
    try:
        return fut.result()
    finally:
        _add_sql_time(time.perf_counter() - start)


async def fetch(sql: str, *args: Any) -> list[asyncpg.Record]:
    sql = _normalize_sql(sql)

    async def _query():
        async with _pool.acquire() as conn:  # type: ignore[union-attr]
            return await conn.fetch(sql, *args)

    return await _run_on_pool(_query)


async def fetchrow(sql: str, *args: Any) -> Optional[asyncpg.Record]:
    sql = _normalize_sql(sql)

    async def _query():
        async with _pool.acquire() as conn:  # type: ignore[union-attr]
            return await conn.fetchrow(sql, *args)

    return await _run_on_pool(_query)


async def fetchval(sql: str, *args: Any) -> Any:
    sql = _normalize_sql(sql)

    async def _query():
        async with _pool.acquire() as conn:  # type: ignore[union-attr]
            return await conn.fetchval(sql, *args)

    return await _run_on_pool(_query)


async def execute(sql: str, *args: Any) -> str:
    sql = _normalize_sql(sql)

    async def _query():
        async with _pool.acquire() as conn:  # type: ignore[union-attr]
            return await conn.execute(sql, *args)

    return await _run_on_pool(_query)


async def executemany(sql: str, args: Iterable[tuple[Any, ...]]) -> None:
    sql = _normalize_sql(sql)
    args = list(args)
    if not args:
        return

    async def _query():
        async with _pool.acquire() as conn:  # type: ignore[union-attr]
            await conn.executemany(sql, args)

    await _run_on_pool(_query)


def fetch_sync(sql: str, *args: Any) -> list[asyncpg.Record]:
    sql = _normalize_sql(sql)

    async def _query():
        async with _pool.acquire() as conn:  # type: ignore[union-attr]
            return await conn.fetch(sql, *args)

    return _run_on_pool_sync(_query)


def fetchrow_sync(sql: str, *args: Any) -> Optional[asyncpg.Record]:
    sql = _normalize_sql(sql)

    async def _query():
        async with _pool.acquire() as conn:  # type: ignore[union-attr]
            return await conn.fetchrow(sql, *args)

    return _run_on_pool_sync(_query)


def fetchval_sync(sql: str, *args: Any) -> Any:
    sql = _normalize_sql(sql)

    async def _query():
        async with _pool.acquire() as conn:  # type: ignore[union-attr]
            return await conn.fetchval(sql, *args)

    return _run_on_pool_sync(_query)


def execute_sync(sql: str, *args: Any) -> str:
    sql = _normalize_sql(sql)

    async def _query():
        async with _pool.acquire() as conn:  # type: ignore[union-attr]
            return await conn.execute(sql, *args)

    return _run_on_pool_sync(_query)


def parse_rowcount(status: str) -> int:
    """Parse asyncpg status strings: 'INSERT 0 1', 'UPDATE 3', 'DELETE 4'."""
    if not status:
        return 0
    parts = status.split()
    for part in reversed(parts):
        if part.isdigit():
            return int(part)
    return 0
