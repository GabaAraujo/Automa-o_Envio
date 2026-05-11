"""Caminhos do projeto e string de conexão ODBC (variáveis de ambiente / .env)."""

from __future__ import annotations

import os
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

BASE_DIR = Path(__file__).resolve().parent
ROOT = BASE_DIR.parent
CONSULTAS_DIR = ROOT / "Consultas"


def load_env() -> None:
    if not load_dotenv:
        return
    load_dotenv(BASE_DIR / ".env")
    load_dotenv(ROOT / ".env")


def get_output_dir() -> Path:
    raw = os.environ.get("REPORT_OUTPUT_DIR", "").strip()
    if raw:
        return Path(raw).expanduser().resolve()
    return (BASE_DIR / "saida").resolve()


def connection_string() -> str:
    full = os.environ.get("MSSQL_ODBC_CONNECTION_STRING")
    if full:
        return full.strip()

    driver = os.environ.get(
        "MSSQL_ODBC_DRIVER", "ODBC Driver 17 for SQL Server"
    ).strip()
    server = os.environ.get("MSSQL_SERVER", "").strip()
    database = os.environ.get("MSSQL_DATABASE", "BMCLIENTES").strip()
    user = os.environ.get("MSSQL_USER", "").strip()
    password = os.environ.get("MSSQL_PASSWORD", "").strip()

    if not server or not user:
        raise SystemExit(
            "Defina MSSQL_SERVER, MSSQL_USER e MSSQL_PASSWORD "
            "(ou MSSQL_ODBC_CONNECTION_STRING completa). "
            "Opcional: MSSQL_DATABASE (padrão BMCLIENTES)."
        )

    return (
        f"DRIVER={{{driver}}};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"UID={user};"
        f"PWD={password};"
        "TrustServerCertificate=yes;"
    )
