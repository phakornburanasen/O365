"""Shared PostgreSQL connection configuration for the backend and MCP server."""

import os
from pathlib import Path

import psycopg2
from dotenv import load_dotenv


BACKEND_DIR = Path(__file__).resolve().parent
load_dotenv(BACKEND_DIR / ".env")


def get_db_connection():
    """Create a PostgreSQL connection using the backend's configuration."""
    return psycopg2.connect(
        host=os.getenv("PGHOST", "10.0.32.71"),
        port=int(os.getenv("PGPORT", "5432")),
        user=os.getenv("PGUSER", "pguser"),
        password=os.getenv("PGPASSWORD", "pgpass123"),
        dbname=os.getenv("PGDATABASE", "postgres"),
    )
