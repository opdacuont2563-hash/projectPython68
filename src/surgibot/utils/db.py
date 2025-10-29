"""SQLite helpers for SurgiBot."""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

from ..config import CONFIG

_CONNECTIONS: dict[Path, sqlite3.Connection] = {}


def _configure(conn: sqlite3.Connection) -> None:
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    conn.execute("PRAGMA cache_size=-20000;")
    conn.execute("PRAGMA temp_store=MEMORY;")


def get_connection(db_name: str) -> sqlite3.Connection:
    path = CONFIG.data_dir / db_name
    if path not in _CONNECTIONS:
        path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(path, check_same_thread=False)
        _configure(conn)
        _CONNECTIONS[path] = conn
    return _CONNECTIONS[path]


@contextmanager
def db_cursor(db_name: str) -> Iterator[sqlite3.Cursor]:
    conn = get_connection(db_name)
    cur = conn.cursor()
    try:
        yield cur
        conn.commit()
    finally:
        cur.close()


__all__ = ["get_connection", "db_cursor"]
