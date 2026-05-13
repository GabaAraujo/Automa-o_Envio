#!/usr/bin/env bash
# Gera os Excels e envia por e-mail (--email). Uso com cron ou systemd.
# Pré-requisitos: Python + deps, driver ODBC Microsoft, arquivo .env nesta pasta.
set -euo pipefail

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$APP_DIR"

LOG_DIR="${LOG_DIR:-$APP_DIR/logs}"
mkdir -p "$LOG_DIR"
LOG_FILE="${LOG_FILE:-$LOG_DIR/relatorio.log}"

if [[ -x "$APP_DIR/.venv/bin/python" ]]; then
  PY="$APP_DIR/.venv/bin/python"
elif [[ -n "${PYTHON:-}" ]]; then
  PY="$PYTHON"
else
  PY="python3"
fi

{
  echo "======== $(date -Iseconds) ========"
  "$PY" main.py --email
  echo "======== fim (exit 0) ========"
} >>"$LOG_FILE" 2>&1
