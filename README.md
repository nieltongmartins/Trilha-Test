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
`data/database/auditoria.db`. Em um ambiente gráfico, abre o Edge visível para
autenticação manual e apresenta a interface Tkinter. Nela é possível selecionar
a planilha, consultar o checkpoint, auditar sem bloquear a janela e gerar ou abrir
o relatório consolidado. Conclua o login/MFA no Edge e somente então clique em
**Atualizar lista**; a abertura da janela não dispara consultas enquanto a
autenticação manual ainda está em andamento. A inicialização é idempotente.

## Distribuição para Windows

A distribuição operacional recomendada é uma pasta autocontida criada com
PyInstaller; ela não requer Node.js e não incorpora configuração corporativa nem
credenciais. Em uma máquina Windows com Python 3.10 ou superior, execute uma vez:

```bat
build_windows.bat
```

O pacote será criado em `dist\AuditorPlanilhas`. Distribua a pasta inteira. Na
primeira utilização, copie `configuracao.exemplo.bat` para `configuracao.bat`, informe
somente `SHAREPOINT_SITE_URL` e `SHAREPOINT_SCOPE_PATHS` e dê duplo clique em
`executar_auditor.bat`. O arquivo real de configuração é ignorado pelo Git e não deve
conter senha, token, cookie ou credencial; login e MFA continuam manuais no Edge.

O launcher fixa a pasta distribuída como diretório de trabalho. Assim, por padrão, o
SQLite canônico, logs, temporários e relatórios ficam respectivamente em
`data\database`, `logs`, `data\temp` e `data\reports` dentro dessa pasta. O operador
precisa ter permissão de escrita nela. Mantenha e proteja `data\database\auditoria.db`:
relatórios e XLSX temporários não substituem o banco. A assinatura do executável, a
política de distribuição, o diretório corporativo definitivo, backup, retenção e
permissões continuam decisões do ambiente operacional.

## Testes

```bash
pytest
```

O núcleo inclui persistência, leitura/comparação Excel, auditoria incremental,
fonte SharePoint, interface desktop e relatório Excel regenerável com as abas
`RESUMO`, `VERSOES` e `TRILHA`.

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
