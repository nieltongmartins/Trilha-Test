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

O núcleo inclui persistência, leitura/comparação Excel, auditoria incremental e a
fonte SharePoint da F4. Interface e relatório pertencem à F5 e não foram iniciados.

## Fonte SharePoint da F4

A V1 usa `BrowserSharePointSource.open_edge(...)` com `SHAREPOINT_SITE_URL` e uma
ou mais raízes server-relative em `SHAREPOINT_SCOPE_PATHS` (separadas por `;`). O
Edge é aberto visivelmente para o usuário concluir login e MFA. A aplicação não
solicita senha, não acessa cookies/tokens e não transfere a sessão para outro
cliente HTTP.

A fonte descobre `.xlsx` recursivamente apenas nos escopos configurados, usa o
`UniqueId` no contexto do site como identidade e trata nome/caminho como atributos
mutáveis. Histórico e arquivo atual são consultados e baixados por endpoints REST
GET distintos; os downloads temporários somente são entregues ao motor após
validação Open XML. O SharePoint permanece estritamente somente leitura.
