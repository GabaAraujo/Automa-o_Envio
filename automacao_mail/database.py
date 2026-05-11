"""Leitura de arquivos SQL e execução no SQL Server via pyodbc."""

from __future__ import annotations

import re
import warnings
from pathlib import Path

import pandas as pd
import pyodbc


def prepare_sql(raw: str) -> str:
    """Converte comentários estilo // para -- (SQL Server não aceita //)."""
    return re.sub(r"(?m)^(\s*)//", r"\1--", raw)


def run_query(conn_str: str, sql_path: Path) -> pd.DataFrame:
    sql = prepare_sql(sql_path.read_text(encoding="utf-8"))
    conn = pyodbc.connect(conn_str)
    try:
        with warnings.catch_warnings():
            warnings.filterwarnings(
                "ignore", message=".*SQLAlchemy.*", category=UserWarning
            )
            return pd.read_sql(sql, conn)
    finally:
        conn.close()
