"""Read-only PostgreSQL helpers using the backend's shared DB config."""

import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, Iterator, List, Sequence

from psycopg2.extensions import connection as PgConnection
from psycopg2.extras import RealDictCursor


BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from backend.database import get_db_connection  # noqa: E402


@contextmanager
def readonly_connection() -> Iterator[PgConnection]:
    """Open the shared backend DB connection and mark it as read-only."""
    conn = get_db_connection()
    try:
        conn.set_session(readonly=True, autocommit=False)
        yield conn
    finally:
        # Read tools never need to commit. Rollback also closes any open txn.
        conn.rollback()
        conn.close()


def fetch_all(sql: str, params: Sequence[Any] = ()) -> List[Dict[str, Any]]:
    """Execute a hard-coded SELECT and return JSON-friendly dictionaries."""
    normalized = sql.lstrip().upper()
    if not (normalized.startswith("SELECT") or normalized.startswith("WITH")):
        raise ValueError("Only SELECT/CTE queries are allowed")

    with readonly_connection() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cursor:
            cursor.execute(sql, params)
            return [dict(row) for row in cursor.fetchall()]


def fetch_one(sql: str, params: Sequence[Any] = ()) -> Dict[str, Any]:
    rows = fetch_all(sql, params)
    return rows[0] if rows else {}
