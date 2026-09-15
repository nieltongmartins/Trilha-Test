# Auditor de Planilhas Excel

Aplicação Python para construir uma trilha de auditoria incremental de
planilhas armazenadas no SharePoint Online. O banco SQLite é a fonte oficial
da trilha; a integração com SharePoint será estritamente de leitura.

## Requisitos

- Python 3.10 ou superior
- dependências de `requirements.txt`

## Preparação

```bash
python -m venv .venv
source .venv/bin/activate  # No Windows: .venv\Scripts\activate
python -m pip install -r requirements.txt
```

As configurações têm valores locais seguros. Para sobrescrevê-las, exporte as
variáveis documentadas em `.env.example`; o projeto não carrega arquivos
`.env` automaticamente nem armazena credenciais.

## Inicialização

```bash
python main.py
```

O comando cria os diretórios operacionais e, se necessário, o banco em
`data/database/auditoria.db`. A inicialização é idempotente.

## Testes

```bash
pytest
```

Nesta fase não existe integração com o SharePoint, leitura de Excel, motor de
comparação, interface ou relatório.
