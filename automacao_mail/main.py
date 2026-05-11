"""
Orquestração: carrega config, executa consultas e gera um .xlsx por relatório (d-1 e d-30).

Uso:
  python main.py              # gera Excel em automacao_mail/saida/
  python main.py --print      # também imprime prévia no console
  python main.py --email      # após gerar, envia os .xlsx por SMTP (ver email_send.py / .env)

Variáveis de ambiente: ver config.py, email_send.py e automacao_mail/.env
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

from config import CONSULTAS_DIR, connection_string, get_output_dir, load_env
from database import run_query
from excel_export import write_report_excel


def gerar_relatorio(
    conn_str: str,
    query_path: Path,
    titulo: str,
    arquivo_saida: Path,
    imprimir: bool,
) -> Path:
    df = run_query(conn_str, query_path)
    if imprimir:
        pd.set_option("display.max_columns", None)
        pd.set_option("display.width", None)
        pd.set_option("display.max_colwidth", 80)
        print("\n" + "=" * 72)
        print(f" {titulo}")
        print(f" Consulta: {query_path.name}")
        print("=" * 72 + "\n")
        print(df.to_string(index=False))
        print(f"\n--- Linhas: {len(df)} ---\n")

    caminho = write_report_excel(df, arquivo_saida, report_title=titulo)
    print(f"[OK] {titulo} -> {caminho}")
    return caminho


def main() -> None:
    parser = argparse.ArgumentParser(description="Relatórios d-1 / d-30 → Excel")
    parser.add_argument(
        "--print",
        action="store_true",
        help="Imprime os dados no console (além de gerar o Excel)",
    )
    parser.add_argument(
        "--email",
        action="store_true",
        help="Após gerar os Excels, envia por e-mail (SMTP no .env)",
    )
    args = parser.parse_args()

    load_env()
    conn_str = connection_string()
    out_dir = get_output_dir()

    relatorios: list[tuple[str, Path, Path]] = [
        (
            "Vistorias - dia anterior (d-1)",
            CONSULTAS_DIR / "d-1",
            out_dir / "relatorio_d-1.xlsx",
        ),
        (
            "Vistorias - janela 16 a 15 (d-30)",
            CONSULTAS_DIR / "d-30",
            out_dir / "relatorio_d-30.xlsx",
        ),
    ]

    gerados: list[Path] = []
    for titulo, caminho_sql, arquivo_xlsx in relatorios:
        if not caminho_sql.is_file():
            print(f"[ERRO] Consulta não encontrada: {caminho_sql}", file=sys.stderr)
            continue
        gerados.append(
            gerar_relatorio(conn_str, caminho_sql, titulo, arquivo_xlsx, args.print)
        )

    if args.email:
        from email_send import send_report_email

        try:
            send_report_email(gerados)
            print("[OK] E-mail enviado com os anexos gerados.")
        except Exception as e:
            print(f"[ERRO] Falha ao enviar e-mail: {e}", file=sys.stderr)
            raise SystemExit(1) from e


if __name__ == "__main__":
    main()
