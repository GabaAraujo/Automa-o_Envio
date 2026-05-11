"""Envio de relatórios por SMTP (anexos .xlsx + corpo texto). Configuração via variáveis de ambiente."""

from __future__ import annotations

import os
import smtplib
import ssl
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path


@dataclass(frozen=True)
class SmtpConfig:
    host: str
    port: int
    user: str
    password: str
    use_tls: bool
    use_ssl: bool


def _parse_recipients(raw: str) -> list[str]:
    out: list[str] = []
    for part in raw.replace(";", ",").split(","):
        addr = part.strip()
        if addr:
            out.append(addr)
    return out


def load_smtp_config() -> tuple[SmtpConfig, str, list[str], list[str], list[str]]:
    """Lê env e retorna (smtp, mail_from, to, cc, bcc)."""
    host = os.environ.get("SMTP_HOST", "").strip()
    port_s = os.environ.get("SMTP_PORT", "587").strip()
    user = os.environ.get("SMTP_USER", "").strip()
    password = "".join(os.environ.get("SMTP_PASSWORD", "").split())
    use_tls = os.environ.get("SMTP_USE_TLS", "true").lower() in (
        "1",
        "true",
        "yes",
        "on",
    )
    use_ssl = os.environ.get("SMTP_USE_SSL", "false").lower() in (
        "1",
        "true",
        "yes",
        "on",
    )

    mail_from = os.environ.get("MAIL_FROM", "").strip()
    mail_to = _parse_recipients(os.environ.get("MAIL_TO", ""))
    mail_cc = _parse_recipients(os.environ.get("MAIL_CC", ""))
    mail_bcc = _parse_recipients(os.environ.get("MAIL_BCC", ""))

    try:
        port = int(port_s)
    except ValueError as exc:
        raise ValueError(f"SMTP_PORT inválido: {port_s}") from exc

    if not host:
        raise ValueError("Defina SMTP_HOST.")
    if not mail_from:
        raise ValueError("Defina MAIL_FROM.")
    if not mail_to:
        raise ValueError(
            "Defina MAIL_TO (um ou mais e-mails, separados por vírgula)."
        )

    cfg = SmtpConfig(
        host=host,
        port=port,
        user=user,
        password=password,
        use_tls=use_tls,
        use_ssl=use_ssl,
    )
    return cfg, mail_from, mail_to, mail_cc, mail_bcc


def default_subject() -> str:
    return (
        f"Relatórios de vistorias — {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )


def default_body(paths: list[Path]) -> str:
    lines = [
        "Segue em anexo a geração dos relatórios.",
        "",
        "Arquivos:",
    ]
    for p in paths:
        p = p.resolve()
        try:
            size_kb = p.stat().st_size / 1024
            extra = f" ({size_kb:.1f} KB)"
        except OSError:
            extra = ""
        lines.append(f"  - {p.name}{extra}")
    lines.extend(["", "---", "Mensagem automática (automação relatório Vix)."])
    return "\n".join(lines)


def send_email_with_attachments(
    subject: str,
    body_plain: str,
    attachment_paths: list[Path],
) -> None:
    """Envia um e-mail multipart com anexos (apenas arquivos existentes)."""
    cfg, mail_from, mail_to, mail_cc, mail_bcc = load_smtp_config()

    existing = [p.resolve() for p in attachment_paths if p.is_file()]
    if not existing:
        raise FileNotFoundError("Nenhum anexo válido (arquivos inexistentes).")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = mail_from
    msg["To"] = ", ".join(mail_to)
    if mail_cc:
        msg["Cc"] = ", ".join(mail_cc)
    if mail_bcc:
        msg["Bcc"] = ", ".join(mail_bcc)
    msg.set_content(body_plain, charset="utf-8")

    for path in existing:
        data = path.read_bytes()
        msg.add_attachment(
            data,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=path.name,
        )

    all_recipients = list(dict.fromkeys(mail_to + mail_cc + mail_bcc))

    context = ssl.create_default_context()
    if cfg.use_ssl:
        with smtplib.SMTP_SSL(cfg.host, cfg.port, context=context) as server:
            if cfg.user:
                server.login(cfg.user, cfg.password)
            server.send_message(msg, from_addr=mail_from, to_addrs=all_recipients)
    else:
        with smtplib.SMTP(cfg.host, cfg.port) as server:
            server.ehlo()
            if cfg.use_tls:
                server.starttls(context=context)
                server.ehlo()
            if cfg.user:
                server.login(cfg.user, cfg.password)
            server.send_message(msg, from_addr=mail_from, to_addrs=all_recipients)


def send_report_email(
    attachment_paths: list[Path],
    *,
    subject: str | None = None,
    body: str | None = None,
) -> None:
    """Atalho: assunto e corpo padrão a partir da lista de caminhos dos .xlsx."""
    paths = [p for p in attachment_paths if p]
    subj = subject if subject is not None else default_subject()
    txt = body if body is not None else default_body(paths)
    send_email_with_attachments(subj, txt, paths)


def main_cli() -> None:
    """Envia relatorio_d-1/d-30 da pasta de saída, ou caminhos após --."""
    import sys

    from config import get_output_dir, load_env

    load_env()
    out = get_output_dir()
    paths: list[Path] = [
        out / "relatorio_d-1.xlsx",
        out / "relatorio_d-30.xlsx",
    ]
    extra = os.environ.get("MAIL_ATTACH_EXTRA", "").strip()
    if extra:
        for part in extra.split(","):
            p = Path(part.strip())
            if p.is_file():
                paths.append(p)

    if "--" in sys.argv:
        idx = sys.argv.index("--")
        paths = [Path(a) for a in sys.argv[idx + 1 :] if Path(a).is_file()]

    send_report_email(paths)
    n = len([p for p in paths if p.is_file()])
    print(f"[OK] E-mail enviado com {n} anexo(s).")


if __name__ == "__main__":
    main_cli()
