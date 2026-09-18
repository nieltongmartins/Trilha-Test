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

O site e os escopos SharePoint podem ser configurados na própria interface e são
persistidos no perfil do usuário (`%APPDATA%\Trilha de Auditoria\config.json` no
Windows; `$XDG_CONFIG_HOME/trilha-de-auditoria/config.json` no Linux). O arquivo
contém **somente** URL e caminhos operacionais: senha, token, cookie e qualquer
outro dado de autenticação nunca são lidos nem persistidos.

A precedência é: variáveis de ambiente > configuração local persistida > valor
inicial seguro. `SHAREPOINT_SITE_URL` e `SHAREPOINT_SCOPE_PATHS` (escopos separados
por `;`) continuam aceitas para execução administrada. O valor inicial do site é
`https://hypermarcas.sharepoint.com/controle_qualidade`; não existe escopo padrão,
pois a pasta auditada deve ser escolhida pelo usuário. O histórico de uso dessas
variáveis durante o desenvolvimento permanece documentado em `.env.example`, mas
não é necessário defini-las a cada novo PowerShell.

## Inicialização

```bash
python main.py
```

O comando cria os diretórios operacionais e, se necessário, o banco em
`data/database/auditoria.db`, e abre primeiro a janela **Trilha de Auditoria**.
Nenhum Edge ou acesso de rede é iniciado nessa etapa. Confira URL/escopo e clique
em **Conectar**; somente então o Edge visível é aberto para login/MFA manual. A
conexão valida a autenticação por uma leitura REST do site e, quando a tela indicar
**Conectado ao SharePoint**, permite clicar em **Atualizar lista**. Inicialização
Selenium, descoberta, consulta de versões, auditoria e geração de relatório são
executadas fora da thread do Tkinter, mantendo a janela responsiva.

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

A V1 usa `BrowserSharePointSource.open_edge(...)` com a URL e uma ou mais raízes
server-relative configuradas (na tela ou pelas variáveis de ambiente). O
Edge é aberto visivelmente para o usuário concluir login e MFA. A aplicação não
solicita senha, não acessa cookies/tokens e não transfere a sessão para outro
cliente HTTP.

A fonte descobre `.xlsx` recursivamente apenas nos escopos configurados, usa o
`UniqueId` no contexto do site como identidade e trata nome/caminho como atributos
mutáveis. Histórico e arquivo atual são consultados e baixados por endpoints REST
GET distintos; os downloads temporários somente são entregues ao motor após
validação Open XML. O SharePoint permanece estritamente somente leitura.

## Validação manual no Windows

1. Feche o PowerShell atual.
2. Abra um novo PowerShell.
3. Entre na pasta do projeto e ative o ambiente virtual, se aplicável.
4. Execute `python main.py` sem definir `$env:SHAREPOINT_SITE_URL` ou
   `$env:SHAREPOINT_SCOPE_PATHS`.
5. Confirme a abertura imediata da janela **Trilha de Auditoria** e que o Edge
   ainda não foi aberto.
6. Informe/confira o site e o caminho completo da biblioteca/pasta no campo
   **Escopo(s)**.
7. Clique em **Conectar**.
8. Autentique-se manualmente no Edge e conclua o MFA.
9. Clique em **Atualizar lista** e confirme a atualização do estado, a seleção e a
    auditoria das planilhas.
10. Feche a aplicação e o Edge.
11. Abra novamente com `python main.py`.
12. Confirme que site e escopo foram recuperados sem comandos `$env:`.

## Auditorias armazenadas e backups locais

A guia **Auditorias armazenadas** consulta exclusivamente o SQLite local e não
abre nem utiliza uma conexão SharePoint. Ela lista a identidade conhecida, o
caminho, a versão/checkpoint, as contagens e a última execução de cada planilha.
As ações da guia afetam somente a trilha local; arquivos no SharePoint nunca são
alterados ou excluídos.

Os backups são gravados em `data/backups/` (ou em
`AUDIT_BACKUPS_DIRECTORY`) sem sobrescrever arquivos anteriores:

- o backup individual é um SQLite dedicado, versionado, com a planilha e todos
  os seus checkpoints, execuções, versões, alterações, erros e hashes. Na
  restauração, uma identidade já existente exige confirmação e é substituída
  integralmente, sem merge; uma identidade ausente é recriada;
- o backup completo é uma imagem consistente produzida pela API oficial de
  backup do SQLite. A restauração valida metadados, versão do schema,
  `integrity_check`, Foreign Keys e SHA-256, e cria automaticamente um backup de
  segurança do estado corrente antes da substituição atômica;
- cada `.sqlite3` possui um arquivo `.sha256`; backups completos também possuem
  `.meta.json`. A divergência, a ausência desses dados ou um formato/schema
  incompatível impede a restauração.

Excluir uma auditoria remove atomicamente somente seus registros dependentes.
Excluir todas limpa os dados na ordem das Foreign Keys, mas mantém tabelas,
índices e versão do schema; a interface oferece backup antes da confirmação
final. A restauração repõe exatamente o checkpoint armazenado: versões posteriores
do SharePoint não são marcadas como processadas e a próxima auditoria incremental
continua pelas regras existentes.
