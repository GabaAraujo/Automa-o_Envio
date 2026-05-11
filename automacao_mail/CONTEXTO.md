# Contexto do projeto — Relatório de atendimentos (automação)

Documento vivo para registrar decisões, requisitos e evolução. Atualize este arquivo conforme o projeto avança (datas e autoria opcionais nas entradas).

---

## Objetivo

Automação em **servidor AWS EC2** para:

1. Consultar banco de dados com registros de **atendimentos técnicos**.
2. Tratar e consolidar dados periodicamente (produtividade da equipe).
3. Gerar **Excel (.xlsx)** e enviar por **e-mail** com resumo das métricas.

---

## Requisitos funcionais (checklist)

| # | Requisito | Status |
|---|-----------|--------|
| 1 | Consulta ao BD (**SQL Server** — `BMCLIENTES`) | Em uso (`Consultas/d-1`, `Consultas/d-30`) |
| 2 | Tratamento: remover duplicidades, padronizar datas, validar obrigatórios | Pendente |
| 3 | Normalizar nomes dos técnicos (acentos, case, fuzzy + consolidação canônica) | Pendente |
| 4 | Excel: linha por atendimento + **aba Resumo** (UF, técnico, tipo operação) | Em uso (detalhe + resumos sintéticos; consolidação “canônica” de nomes ainda pendente) |
| 5 | Métrica obrigatória: **total de atendimentos por técnico** | Pendente |
| 6 | Métricas extras: % participação, média por período, distribuição no tempo | Pendente |
| 7 | Relatório legível para gestão | Pendente |
| 8 | E-mail automático com anexo + **corpo com resumo** das principais métricas | Em uso (`email_send.py`; corpo lista anexos — métricas textuais finas podem evoluir) |
| 9 | EC2: cron/timer, env para credenciais, logs e tratamento de erros | Pendente |

---

## Arquitetura (visão)

```
[cron / systemd] → script Python → [SGBD]
                        ↓
              tratamento + normalização de técnicos
                        ↓
                  geração .xlsx (abas detalhe + por técnico)
                        ↓
                  SMTP (ou SES) → destinatários
                        ↓
                  logs (arquivo / CloudWatch opcional)
```

---

## Stack sugerida

| Uso | Biblioteca / ferramenta |
|-----|-------------------------|
| Runtime | Python 3.11+ |
| Dados | pandas |
| BD | **pyodbc** (SQL Server) |
| Excel | **openpyxl** (via pandas `ExcelWriter`) |
| E-mail | smtplib + email.mime (ou boto3 + SES) |
| Normalização fuzzy | RapidFuzz ou thefuzz |
| Acentos / texto | unicodedata ou unidecode |
| Config local (dev) | python-dotenv (produção: env do sistema ou Secrets Manager) |

---

## Fluxo lógico (resumo)

1. Carregar config (período, destinatários, conexão via env).
2. Extrair dados com query parametrizada.
3. Dedupe, datas em timezone definido, validação de campos obrigatórios.
4. Pipeline de nome: limpeza → mapa de aliases → opcional fuzzy com threshold + log de ambiguidades.
5. Calcular métricas (contagem por técnico canônico + extras acordadas).
6. Montar Excel (detalhe + consolidação; opcional série temporal).
7. Enviar e-mail com anexo e resumo textual.
8. Log de sucesso/falha e exit code para monitoração.

---

## Pontos críticos de qualidade

- **Regra de duplicidade:** definir chave natural com o negócio (ex.: id do ticket ou combinação de campos).
- **Fuso horário:** um timezone de referência (ex.: `America/Sao_Paulo`) para agregações por dia.
- **Fuzzy:** não substituir mapeamento manual; alto risco de fundir duas pessoas; usar threshold e log.
- **Secrets:** nunca no código; variáveis de ambiente ou AWS Secrets Manager / SSM.

---

## Variáveis de ambiente (rascunho — preencher)

Documente aqui os nomes reais usados no código quando definidos.

| Variável | Descrição |
|----------|-----------|
| `MSSQL_ODBC_DRIVER` | Nome do driver ODBC (ex.: `SQL Server`, `ODBC Driver 17 for SQL Server`) |
| `MSSQL_SERVER` | Host ou IP do SQL Server |
| `MSSQL_DATABASE` | Banco (ex.: `BMCLIENTES`) |
| `MSSQL_USER` / `MSSQL_PASSWORD` | Credenciais SQL |
| `MSSQL_ODBC_CONNECTION_STRING` | Alternativa: string ODBC completa (sobrepõe os campos acima) |
| `REPORT_OUTPUT_DIR` | Pasta de saída dos `.xlsx` (padrão: `automacao_mail/saida/`) |
| `SMTP_HOST` | Servidor SMTP |
| `SMTP_PORT` | Porta (ex.: **587** STARTTLS, **465** SSL implícito) |
| `SMTP_USE_TLS` | `true`/`false` — STARTTLS após conectar (típico na 587) |
| `SMTP_USE_SSL` | `true`/`false` — conexão SSL desde o início (típico na 465) |
| `SMTP_USER` / `SMTP_PASSWORD` | Autenticação SMTP (se vazio, tenta enviar sem login — raro) |
| `MAIL_FROM` | Endereço remetente |
| `MAIL_TO` | Destinatários (vírgula ou ponto e vírgula) |
| `MAIL_CC` / `MAIL_BCC` | Opcional |
| `MAIL_ATTACH_EXTRA` | Somente em `python email_send.py`: caminhos extras de anexo (vírgula) |

---

## Estrutura de pastas (atual)

```
automacao_mail/
├── CONTEXTO.md
├── main.py           ← orquestração (CLI)
├── config.py         ← paths, load_dotenv, connection_string
├── database.py       ← prepare_sql, run_query
├── excel_export.py   ← prepare_detail_dataframe, abas Detalhe + Resumo
├── email_send.py     ← SMTP, anexos .xlsx, corpo texto
├── requirements.txt
├── .env              ← credenciais (não versionar)
├── saida/            ← relatórios gerados (relatorio_d-1.xlsx, relatorio_d-30.xlsx)
└── ...

../Consultas/
├── d-1               ← SQL: dia anterior
└── d-30              ← SQL: janela 16–15
```

---

## Registro de alterações / decisões

Use a seção abaixo como log cronológico (copie o bloco modelo para cada entrada).

### Modelo de entrada

```
### YYYY-MM-DD — título curto
- O que mudou ou foi decidido:
- Próximos passos:
```

---

### 2026-05-07 — Documento de contexto inicial

- Criado `CONTEXTO.md` com objetivo, requisitos, arquitetura, stack, fluxo, riscos de dados e placeholders de configuração.
- Próximos passos: definir SGBD e schema; escrever query base; implementar pipeline em `main.py`; preencher variáveis de ambiente reais.

---

### 2026-05-08 — Excel modular + resumo sintético

- **Modularização:** `config.py` (env e ODBC), `database.py` (queries em `Consultas/`), `excel_export.py` (processamento + `.xlsx`), `main.py` apenas chama `gerar_relatorio` para `d-1` e `d-30`.
- **Saída:** dois arquivos em `automacao_mail/saida/`: `relatorio_d-1.xlsx` e `relatorio_d-30.xlsx`. Cada um tem aba **Detalhe** (dados tratados e cabeçalhos legíveis) e **Resumo** (totais, distribuição por UF, participações por técnico nas colunas de executante, por tipo de operação).
- **CLI:** `python main.py` gera os Excels; `python main.py --print` mantém prévia no console.
- **Variáveis:** uso de `MSSQL_*` e opcional `REPORT_OUTPUT_DIR` documentados na tabela deste arquivo.
- Próximos passos: normalização canônica de nomes de técnicos; opcional segunda aba “por técnico único” conforme regra de negócio.

---

### 2026-05-08 — Envio por e-mail (SMTP)

- Novo módulo **`email_send.py`**: `send_report_email` / `send_email_with_attachments`, leitura de `SMTP_*`, `MAIL_FROM`, `MAIL_TO`, `MAIL_CC`, `MAIL_BCC`; suporte a STARTTLS (587) e SSL implícito (465).
- **`main.py --email`**: após gerar os dois `.xlsx`, envia ambos em um único e-mail com lista de anexos no corpo.
- **`python email_send.py`**: reenvia os arquivos padrão em `saida/` (ou caminhos após `--`), útil sem rodar o BD.
- Placeholders SMTP adicionados em `.env` (comentados); documentação das variáveis na tabela deste arquivo.

---

## Links e referências internas

- Queries SQL: `Consultas/d-1`, `Consultas/d-30` (pasta pai de `automacao_mail`).

---

## Notas livres

(Espaço para anotações rápidas — convenções de nome de arquivo do relatório, horário do cron, contatos de TI, etc.)

### Lembrete — após a automação: disponibilidade da API

**Quando:** depois que a automação do relatório/e-mail estiver estável.

**Objetivo:** ter um sistema que **confirme com frequência** que a API está respondendo (e alertar se cair), não apenas “achar que está no ar”.

**Ideias para implementar depois (escolher o que couber na stack):**

| Camada | Opções |
|--------|--------|
| Na API | Endpoint leve `GET /health` ou `/ready` (status 200 + JSON mínimo); opcional checar BD ou dependência crítica em `/ready`. |
| Infra / servidor | Na EC2: `systemd` com `Restart=always` para o processo da API; health check do balanceador ou do target group, se houver. |
| Monitor externo | Serviço de uptime (ping HTTP periódico) ou **cron** em outro host que chama a URL da API e registra falha; alerta por e-mail/Slack/Teams. |
| Observabilidade | Logs estruturados + métricas (ex.: CloudWatch, ou contador de erros 5xx). |

**Critério de sucesso sugerido:** falha detectada em até X minutos + notificação para quem opera o sistema + tentativa automática de recuperação quando aplicável.

**Próximo passo quando for a hora:** definir URL base da API, SLA desejado (ex.: verificar a cada 1–5 min) e canal de alerta.
