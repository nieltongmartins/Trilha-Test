# HISTÓRICO DE IMPLEMENTAÇÕES
## Auditor de Planilhas Excel — SharePoint Online

**Documento:** 04_HISTORICO_IMPLEMENTACOES.md
**Versão:** 1.1
**Status:** Oficial
**Data de criação:** 15/09/2026
**Última revisão:** 16/09/2026

---

# CICLO FINAL, EMPACOTAMENTO E ENCERRAMENTO DA V1 — 16/09/2026

**Autorização:** decisão expressa do responsável autorizou inventariar, executar e
concluir todas as tarefas oficiais restantes sem nova autorização entre elas, além dos
commits tecnicamente exigidos pelo ambiente. A exceção e seu objetivo de encerramento
foram aplicados sem alterar as invariantes de segurança ou o escopo funcional.

## Estado inicial e inventário

As fases F1 a F5 estavam concluídas. A F6 estava em andamento com testes finais,
performance, logs, temporários, SHA-256 e backup concluídos. A seção 43 exigia apenas
documentar a estratégia mínima e já possuía evidência suficiente; não foi reimplementada.
A única tarefa de implementação pendente era empacotamento (seção 44), seguida da
validação integrada, reconciliação documental e encerramento. Não foi encontrada tarefa
oficial posterior.

## Empacotamento e alterações

Foi definida distribuição Windows one-folder por `auditor_planilhas.spec`, com
PyInstaller como dependência exclusiva de build em `requirements-build.txt`.
`build_windows.bat` gera a pasta `dist\AuditorPlanilhas` e inclui o launcher e o exemplo
de configuração. `executar_auditor.bat` fixa sua própria pasta como diretório de
trabalho, exige `configuracao.bat`, não versionado, e inicia a aplicação por duplo
clique. O exemplo contém apenas site, escopos e caminhos opcionais, sem credenciais.
README e documentos oficiais receberam instruções e limites operacionais. Não houve
alteração no código de produção, schema, migração, protocolo SharePoint, checkpoint,
relatório, logs, temporários ou SHA-256.

A tentativa de instalar PyInstaller neste Linux falhou porque o índice configurado foi
inacessível (`403 Forbidden`). Além disso, PyInstaller não faz cross-build de Windows.
Isso limita a materialização do `.exe` neste ambiente, não a definição reproduzível do
pacote: a geração/homologação final deve ocorrer em Windows. Assinatura, canal e pasta
corporativa, permissões e implantação da estratégia de backup continuam decisões
operacionais.

## Validação final integrada

- `pytest -q`: **52 passed em 1,74 s**, sem falhas, skips ou warnings relevantes;
- seleção integrada cobrindo banco legado, rollback/retomada, reexecução, incremental,
  relatório, múltiplas planilhas, SHA-256, temporários, logs e empacotamento:
  **16 passed em 1,12 s**;
- inicialização sem UI em diretório temporário: código 0, banco novo criado e fechado;
- banco novo: `PRAGMA integrity_check` = `ok` e `PRAGMA foreign_key_check` = `[]`;
- `python -m compileall -q app main.py tests`, compilação do spec, `ruff check .` e
  `git diff --check`: aprovados.

A suíte preservou ADD/MOD/DEL, fórmulas, zero/False/vazio, múltiplas abas, versão sem
alterações, pares adjacentes, checkpoint, rollback, retomada, isolamento, migração
legada, hash do conteúdo processado, limpeza/órfãos, rotação/redação de logs e relatório
`RESUMO`/`VERSOES`/`TRILHA` regenerável sem XLSX históricos. A prova corporativa real
SharePoint/Edge já aceita na F5 não foi repetida no Linux; nenhum SharePoint real foi
acessado ou alterado neste ciclo.

## Estado final, riscos e limites

F1, F2, F3, F4, F5 e F6 estão concluídas. Todos os critérios da F6 foram marcados como
atendidos e não existe tarefa oficial pendente. SQLite permanece canônico; backup
continua como estratégia documentada, não rotina operacional nem restauração. O gargalo
conhecido do relatório com 100.000 alterações permanece aceito, sem SLA definido e sem
redesenho preventivo. Disponibilidade histórica depende do SharePoint; autoria/horário
são da versão, não da célula; fórmulas não são recalculadas; operação real depende de
Windows, Edge, políticas corporativas, permissões de escrita locais e autenticação
manual. SQL Server, agendamento, lote, dashboards, notificações, política avançada de
retenção e hardening adicional permanecem melhorias futuras fora do escopo.

**Estado oficial:** ELABORAÇÃO DA VERSÃO ATUAL CONCLUÍDA PARA UTILIZAÇÃO E TESTES
REAIS. Problemas descobertos em utilização/homologação passam a ser correção, manutenção,
hardening, melhoria ou evolução.

# ESTRATÉGIA MÍNIMA DE BACKUP DA F6 — 16/09/2026

**Fase:** F6 — Robustez e Preparação para Produção. **Status:** tarefa de backup
concluída; a F6 permanece em andamento.

## Confirmação documental e decisão

A seção 43 continuava sendo a próxima tarefa oficial e possui exatamente dois
critérios: documentar uma estratégia mínima de backup do banco e considerar o banco de
auditoria um ativo crítico. A arquitetura já estabelecia o SQLite em
`data/database/auditoria.db` como fonte canônica e indicava que uma estratégia de
produção deveria tratar backup, controle de acesso, retenção e restauração, mas não
definia infraestrutura, frequência, prazo de retenção ou procedimento de restauração.
Não existia estratégia anterior nem código específico de backup.

Por isso, esta tarefa é exclusivamente documental, conforme o verbo e o limite da
seção 43. Não foi criada funcionalidade de produção, política arbitrária de retenção ou
restauração fora do escopo. A ausência das decisões operacionais foi registrada como
limitação a resolver antes de depender do backup em produção, sem conflito documental,
técnico ou de governança que impedisse a documentação da estratégia.

## Estratégia documentada

- o único ativo canônico incluído é o arquivo SQLite configurado, contendo planilhas,
  versões processadas, alterações, checkpoints, execuções, erros, hashes, metadados e
  relacionamentos definidos pelo schema;
- XLSX históricos de `data/temp/`, relatórios regeneráveis, logs, Git e código-fonte são
  excluídos;
- uma implementação futura deverá usar `sqlite3.Connection.backup`, API oficial de
  backup online do SQLite, em vez de copiar ingenuamente um banco aberto. Isso permite
  snapshot consistente diante de escrita por outra conexão;
- o destino deverá ser protegido e separado do arquivo ativo. O nome deverá usar o
  identificador do banco e timestamp UTC com microssegundos; arquivo existente não
  poderá ser silenciosamente substituído;
- somente uma cópia concluída e aprovada por `PRAGMA integrity_check` e
  `PRAGMA foreign_key_check` poderá ser publicada como válida. Falhas e artefatos
  parciais não poderão alterar o banco original nem sua trilha/checkpoint;
- não haverá upload externo, criptografia própria, retenção ou limpeza automática sem
  decisão formal. Frequência, destino concreto, responsável, acesso e retenção seguem
  pendentes para o ambiente de produção;
- restauração não faz parte do critério oficial. Foram registrados apenas guardrails
  para sua futura definição: aplicação parada, cópia validada, ausência de substituição
  silenciosa e nova validação antes do uso.

## Validação e limites

Não houve alteração de código ou schema e, portanto, não foram criados testes de
backup/restauração nem executada restauração. A revisão estática confirmou que o banco
canônico concentra as seis tabelas oficiais e que temporários, relatórios e logs são
diretórios separados. `pytest -q` foi executado como regressão da documentação e teve
50 testes aprovados. `git diff --check` foi concluído sem erros.

Não há resultado novo de `integrity_check` ou `foreign_key_check` sobre backup, pois
nenhum arquivo de backup foi criado; alegar esses resultados ampliaria ou falsearia o
escopo documental. Os resultados da tarefa de integridade anterior permanecem válidos
somente para seus bancos temporários de teste. Não houve medição nem impacto de
performance.

O commit adicional é realizado exclusivamente porque o ambiente exige commit e criação
de pull request, apesar do limite ordinário da F6 já ultrapassado. O hash é informado no
relatório da sessão.

**Próxima tarefa oficial:** empacotamento, seção 44 do plano, pendente de nova
autorização. Não foi executada. A F6 não está concluída.

---

# INTEGRIDADE POR SHA-256 DA F6 — 16/09/2026

**Fase:** F6 — Robustez e Preparação para Produção. **Status:** tarefa de integridade
concluída; a F6 permanece em andamento.

## Confirmação documental e decisão

A seção 42 continuava sendo a próxima tarefa oficial e continha somente dois critérios:
avaliar/implementar SHA-256 das versões processadas quando tecnicamente adequado e
registrar claramente sua finalidade. A especificação proibia falsa garantia de
segurança ou imutabilidade, e a arquitetura já previa o campo opcional
`versao_processada.hash_origem` para uma impressão digital do binário usado. Portanto,
foi considerado adequado preencher o campo existente, sem ampliar a tarefa para uma
revisão geral de integridade e sem mudança conceitual do schema.

O SHA-256 que já existia no controle de temporários não atendia a esse objetivo: ele é
derivado dos identificadores contextuais e serve somente para produzir um nome seguro,
sem colisões de caminho. O mecanismo desta tarefa usa os bytes reais do XLSX e permite
conferir que outro binário é idêntico ao conteúdo adquirido e processado naquela
comparação. Não comprova autoria nem autenticidade da origem; não substitui trilha,
checkpoint, identidade SharePoint, assinatura, backup ou proteção do SQLite; e não
transforma o XLSX temporário em evidência permanente.

## Implementação e comportamento

- `app/integrity.py` calcula SHA-256 em blocos de 1 MiB com `hashlib`, produzindo
  hexadecimal minúsculo determinístico sem carregar o arquivo inteiro em memória;
- o `AuditService` calcula o digest antes de ler o workbook e antes de liberar o
  temporário, sobre o binário exato adquirido para o lado atual da comparação;
- `hash_origem` é gravado atomicamente com `versao_processada`, alterações e checkpoint,
  para versões históricas ou para a versão atual conforme a ordem oficial da fonte;
- o schema já continha `hash_origem`. A migração idempotente passou a acrescentá-lo
  apenas em bancos legados que não o possuem, preservando registros existentes, e o
  `user_version` passou a 2;
- auditorias inicial e incremental registram o hash de cada versão consolidada como
  lado atual. Retomada conserva hashes confirmados e grava apenas pares pendentes;
  execução sem novidades não relê XLSX nem cria/duplica hashes;
- a primeira versão usada somente como baseline inicial não possui linha própria de
  versão processada nem hash isolado, em conformidade com o modelo singular existente;
- os temporários continuam removidos imediatamente, e relatórios continuam regeneráveis
  exclusivamente a partir do banco.

## Validação e limites

`pytest -q` — `50 passed in 1.71s`.

O teste novo valida mesmo conteúdo, inclusive lido em blocos pequenos, produzindo o
mesmo SHA-256 esperado e conteúdo diferente produzindo digest diferente. Testes do
serviço validam a persistência para versões históricas, versão sem alterações e fluxo
incremental; a suíte existente cobre versão atual SharePoint, reinicialização,
idempotência, falha/retomada, descarte de temporários e relatório sem XLSX históricos.
A migração foi validada contra banco legado sem a coluna, com reinicialização
idempotente e preservação dos dados. Nenhum SharePoint real foi acessado.

`PRAGMA integrity_check` retornou `ok` e `PRAGMA foreign_key_check` não encontrou
violações nos cenários de auditoria incremental, reexecução e retomada da suíte.

O cálculo acrescenta uma leitura sequencial do XLSX por versão. O uso de blocos limita
memória e não altera o arquivo; não foi repetido o benchmark completo porque não houve
evidência objetiva para fazê-lo. A conferência futura depende de o binário candidato
ainda estar disponível por meio autorizado; o hash sozinho não o recupera.

O commit adicional é realizado exclusivamente porque o ambiente exige commit e criação
de pull request, apesar do limite ordinário da F6 já ultrapassado. O hash é informado no
relatório da sessão.

**Próxima tarefa oficial:** backup, seção 43 do plano, pendente de nova autorização.
Não foi executada. A F6 não está concluída.

---

# CONTROLE DE ARQUIVOS TEMPORÁRIOS DA F6 — 16/09/2026

**Fase:** F6 — Robustez e Preparação para Produção. **Status:** tarefa de arquivos
temporários concluída; a F6 permanece em andamento.

## Confirmação documental e estado anterior

A seção 41 do plano era a próxima tarefa oficial e exigia validar criação,
utilização, remoção, recuperação após falha e a separação entre temporários e
evidência oficial. A arquitetura permitia downloads em `data/temp/`, descarte após
sucesso e proibia acúmulo indefinido. Não houve conflito de requisito, segurança ou
governança.

As fontes Edge/REST e Graph já usavam subdiretórios temporários por instância e os
removiam no fechamento normal. Entretanto, os XLSX permaneciam até o fechamento da
fonte, uma interrupção abrupta podia deixar diretórios órfãos sem tratamento na
reinicialização, o Graph numerava arquivos pela quantidade existente e a fonte Edge
incorporava identificadores externos ao nome. Em auditorias extensas isso permitia
crescimento até o encerramento e faltava uma garantia explícita contra colisão ou
componentes de caminho maliciosos.

## Alterações e comportamento

- um workspace comum cria subdiretório exclusivo e marcador de propriedade/processo;
- nomes determinísticos usam SHA-256 sobre site, contexto/drive, identidade da
  planilha e ID oficial da versão, sem incorporar esses valores ao caminho;
- Edge/REST e Graph escrevem apenas dentro do workspace; o Graph publica o download
  por substituição atômica de arquivo parcial;
- o `AuditService` remove cada XLSX adquirido logo após a leitura, inclusive quando a
  leitura ou uma etapa posterior falha. Falha de limpeza é registrada como aviso e não
  transforma temporário em evidência nem desfaz dados canônicos;
- o fechamento normal remove o workspace inteiro. Na reinicialização, somente
  diretórios com prefixo e marcador válidos, pertencentes à aplicação e cujo PID não
  existe mais são removidos; arquivos externos, marcadores inválidos e processos ativos
  são preservados;
- fontes locais nunca removem fixtures/arquivos fornecidos pelo usuário.

Na auditoria incremental, a baseline e a versão atual coexistem logicamente como
snapshots em memória durante a comparação, embora os XLSX sejam liberados assim que
lidos. Falha preserva as transações já confirmadas e o checkpoint; retomada readquire a
baseline oficial da fonte. Se ela não estiver disponível, a regra existente registra
falha e impede comparação não adjacente. Reinicialização não depende dos XLSX antigos.
Relatórios continuam regeneráveis exclusivamente pelo SQLite.

## Validação e limites

`pytest -q` — `48 passed in 1.26s`.

`ruff check .` — aprovado.

`python -m compileall -q app main.py tests` — concluído sem erros.

Os testes cobriram isolamento entre workspaces, nomes sem colisão entre versões e
planilhas, remoção imediata, limpeza após falha de download/comparação, preservação de
arquivo externo e recuperação de órfão na inicialização. A suíte existente manteve
cobertura de auditoria inicial/incremental, sem novidades, falha/retomada, múltiplas
planilhas, relatório sem XLSX e integridade transacional. `PRAGMA integrity_check`
retornou `ok` e `PRAGMA foreign_key_check` permaneceu sem violações nesses cenários.
Nenhum SharePoint real foi acessado.

Uma terminação abrupta pode deixar resíduos até a próxima criação de fonte. Um PID
reutilizado pelo sistema operacional pode adiar conservadoramente a remoção para evitar
apagar workspace possivelmente ativo; isso não cria dependência funcional ou canônica.

O commit adicional é realizado exclusivamente porque o ambiente exige commit e criação
de pull request, apesar do limite ordinário da F6 já ultrapassado. O hash é informado no
relatório da sessão.

**Próxima tarefa oficial:** integridade, seção 42 do plano, pendente de nova
autorização. Não foi executada. A F6 não está concluída.

---

# REVISÃO E FORTALECIMENTO DE LOGS DA F6 — 16/09/2026

**Fase:** F6 — Robustez e Preparação para Produção. **Status:** tarefa de logs
concluída; a F6 permanece em andamento.

## Confirmação documental anterior às alterações

A documentação obrigatória, o estado limpo da branch `work`, os commits alcançáveis e
a implementação foram inspecionados antes das alterações. A seção 40 do plano define
quatro critérios: revisar logs técnicos, mensagens de erro, ausência de segredos e
rastreabilidade das execuções. Testes finais e performance já estavam concluídos, e o
histórico e o estado oficial do prompt apontavam logs como a próxima tarefa. A seção 55
do plano ainda citava a antiga tarefa de falha/retomada; tratava-se de texto obsoleto,
resolvido de forma inequívoca pela autorização expressa atual, pelos status das seções
38/39 e pelo histórico factual. Não houve conflito de requisito, segurança ou
arquitetura.

Antes desta tarefa, `app/logging_config.py` já criava `logs/auditoria.log` (ou o caminho
configurado), escrevia em UTF-8 e acrescentava conteúdo entre execuções. Entretanto,
somente `main.py` emitia uma mensagem, depois da abertura do banco. Auditoria,
checkpoint, aquisição SharePoint, autenticação manual, descoberta, comparação, falha,
retomada, execução sem novidades e relatório não produziam evidência operacional; o
arquivo também crescia sem limite e não tinha proteção defensiva caso um erro contivesse
um campo de autenticação. Essa divergência em relação à seção 40 da arquitetura exigia
fortalecimento, mas não mudança arquitetural: foi mantido o módulo central e utilizado o
logging padrão já existente.

## Alterações realizadas

- o arquivo permanece persistente e UTF-8, no diretório configurável `logs/` por
  padrão, agora com rotação de 5 MiB e cinco cópias de retenção;
- o formato ganhou timestamp, nível e logger de origem; um formatador redige valores
  associados a senha, token, segredo, cookie e autorização, inclusive em texto de
  exceção;
- a inicialização registra abertura/configuração do SQLite e encerramento; erros
  inesperados são registrados como `CRITICAL` com traceback e continuam sendo
  propagados;
- autenticação manual no Edge, operação SharePoint read-only, descoberta de planilhas e
  contagens de versões passaram a ser registradas sem cookies, tokens ou credenciais;
- cada auditoria recebe nos logs o mesmo `codigo_execucao` persistido no SQLite, junto
  de planilha/identidade, checkpoint, contagens, conclusão, ausência de novidades ou
  falha. Comparações e aquisições individuais usam somente `DEBUG`, evitando milhares
  de mensagens no nível operacional `INFO`; nenhuma alteração célula a célula é
  registrada;
- geração de relatório registra início, arquivo e totais. Falhas controladas da
  interface registram tipo e mensagem úteis ao diagnóstico;
- foram adicionados testes determinísticos com diretórios temporários para arquivo,
  persistência entre reinicializações, auditoria concluída, execução sem novidades,
  relatório, falha controlada, redação de segredos e rotação.

## Validação realizada

`pytest -q tests/test_logging.py tests/test_main.py` — `4 passed in 0.36s`.

`ruff check .` — aprovado.

`python -m compileall -q app main.py tests` — concluído sem erros.

`pytest -q` — `46 passed in 1.40s`.

`git diff --check` — concluído sem erros.

No cenário de reinicialização do teste de logs, `PRAGMA integrity_check` retornou `ok`
e `PRAGMA foreign_key_check` não retornou violações. Nenhum SharePoint real foi
acessado; os testes usaram fonte local/fakes. O provider continua estritamente
read-only e nenhuma captura de senha, cookie, token, cabeçalho de autenticação ou
credencial foi adicionada.

## Limitações, riscos e decisão

A redação é uma defesa adicional para campos de autenticação identificáveis, não uma
autorização para capturar dados de sessão nem uma garantia de classificar todo texto
arbitrário fornecido por sistemas externos. Por isso, o código continua sem coletar
segredos. Nome, identidade e caminho de planilha são dados operacionais presentes no
log e o diretório deve receber as permissões adequadas do usuário operacional. Se o
diretório/arquivo não puder ser criado ou escrito, a configuração falha explicitamente
em vez de simular persistência. A política local mantém no máximo o arquivo ativo mais
cinco cópias de 5 MiB; não foi introduzida observabilidade externa.

Não foi repetido o cenário completo de performance. O nível `INFO` registra eventos por
execução e contagens agregadas; os eventos por versão/comparação ficam em `DEBUG`, e não
há logging por célula. Não foi identificado impacto relevante no caminho operacional
padrão. A retomada transacional continua coberta pelos testes anteriores da F6 e agora
cada nova tentativa possui código correlacionável próprio.

O commit adicional é realizado exclusivamente porque o ambiente exige commit e criação
de pull request, apesar de o limite ordinário da F6 já ter sido ultrapassado. O hash é
informado no relatório da sessão, pois um commit não pode conter o próprio hash.

**Próxima tarefa oficial:** arquivos temporários, seção 41 do plano, pendente de nova
autorização. Não foi executada nesta sessão. A F6 não está concluída.

---

# ACEITAÇÃO REAL DA F5 E INÍCIO DA F6 — 16/09/2026

## Encerramento formal da F5

A validação real foi concluída com sucesso no Windows corporativo, usando SharePoint
Online real, Selenium + Microsoft Edge e o banco SQLite existente. A planilha
`TEST001_RENOMEADO.xlsx` apresentou última versão auditada `1.3`, última disponível
`1.3` e zero versões pendentes.

Após a migração aditiva, o banco continha 2 planilhas, 3 versões processadas, 90
alterações, 1 checkpoint, 16 execuções e 10 erros de processamento. Os 10 erros
anteriores permaneceram como histórico, comprovando que a migração não apagou os dados
existentes.

Foram aceitos: aquisição e processamento de versões históricas, comparação, persistência
das alterações, avanço do checkpoint e relatório produzido do banco com as abas
`RESUMO`, `VERSOES` e `TRILHA`. O SharePoint permaneceu estritamente read-only; os XLSX
históricos permaneceram temporários; o SQLite continuou como fonte canônica da trilha e
dos checkpoints. Mensagens visuais de erro de download ou verificação do Edge podem
ocorrer mesmo quando o arquivo foi criado e validado no filesystem; isso não autoriza
alterar políticas do Edge, implementar bypass ou extrair credenciais ou artefatos de
sessão.

**Status final:** 🟢 F5 CONCLUÍDA.

## Requisitos e riscos de produção registrados

O cenário informado prevê aproximadamente 2.000 planilhas, algumas com mais de 3.000
versões, e potencialmente milhões de alterações. A F6 deverá obter evidência de
capacidade para índices, consultas, crescimento do banco, memória, relatórios, carga
inicial, incrementais e lotes. SQLite não será substituído preventivamente.

Um único banco canônico deverá suportar muitas planilhas, cada uma com identidade,
versões, alterações, checkpoint, execuções e erros independentes. Também permanece um
requisito de configuração/usabilidade tratar `SHAREPOINT_SITE_URL` e
`SHAREPOINT_SCOPE_PATHS` sem hardcode de escopos transitórios como `Z-Testes VSC`.

## Primeira tarefa da F6 — testes unitários

**Status:** concluída. Foram executados somente os testes unitários isolados de
configuração, banco, leitura, comparação, fonte SharePoint com fakes, relatório e
interface com fakes. Testes de integração e os demais itens de Testes Finais não foram
antecipados.

**Próxima tarefa:** testes de integração, pendente de nova autorização.

## Segunda tarefa da F6 — testes de integração

**Status:** concluída. Foram executados exclusivamente os testes automatizados que
integram componentes, sem antecipar os cenários independentes de reexecução, auditoria
incremental, falha/retomada ou relatório. A seleção exercitou: fonte local →
`AuditService` → reader/comparator → SQLite; inicialização da aplicação → configuração
→ criação do banco/log; e provider Edge/REST fake → endpoints históricos/atual →
download e validação Open XML.

**Comando executado:**

`pytest -q tests/test_audit_service.py::test_complete_audit_records_changes_empty_version_checkpoint_and_execution tests/test_main.py::test_application_starts_and_creates_database tests/test_sharepoint_source.py::test_downloads_historical_and_current_using_distinct_read_only_endpoints`

**Resultado:** `3 passed in 0.47s`. Nenhuma conexão externa, credencial ou operação no
SharePoint foi utilizada. A integração real já aceita na F5 não foi repetida neste
ambiente, que não possui a sessão corporativa Edge/SharePoint. Nenhum problema ou
bloqueio foi encontrado.

**Próxima tarefa:** reexecução, pendente de nova autorização.

## Terceira tarefa da F6 — reexecução

**Status:** concluída. O teste automatizado executou uma auditoria completa, encerrou e
reabriu a conexão da aplicação com o mesmo SQLite e reexecutou a auditoria sem novas
versões. Antes da reexecução, os XLSX temporários foram removidos para comprovar que o
conteúdo já processado não é relido nem readquirido quando não há novidades.

Foram comprovados: nenhum novo registro em `versao_processada`, `alteracao` ou
`planilha`; checkpoint integralmente preservado; dados anteriormente persistidos sem
modificação; resultado determinístico `CONCLUIDA_SEM_NOVIDADES`, com zero versões e
zero alterações; nova execução com código próprio, checkpoint inicial/final `0.99`,
término registrado e sem mensagem de erro; `PRAGMA integrity_check` igual a `ok`; e
`PRAGMA foreign_key_check` sem violações. A fonte da reexecução foi local e somente
listou metadados. O teste de segurança do provider confirmou novamente que o código
embutido usa apenas GET same-origin sob `/_api/`; nenhum SharePoint real foi acessado
ou alterado.

**Comando executado:**

`pytest -q tests/test_audit_service.py::test_reexecution_without_new_versions_is_idempotent tests/test_sharepoint_source.py::test_only_get_same_origin_api_is_embedded`

**Resultado:** `2 passed in 0.35s`. Nenhum problema ou bloqueio foi encontrado.

**Próxima tarefa:** auditoria incremental, pendente de nova autorização. O terceiro
commit ordinário da F6 foi consumido nesta tarefa; não realizar outro commit da fase
sem autorização expressa ou procedimento previsto pela governança.

## Quarta tarefa da F6 — auditoria incremental

**Status:** concluída. O teste automatizado partiu de um histórico já consolidado até o
checkpoint `0.99` e acrescentou as versões `1.00`, `1.01` e `1.02`. Os arquivos das
versões anteriores a `0.99` foram removidos antes da continuação para comprovar que o
histórico antigo não é readquirido; a aquisição observada ficou restrita à versão-base
`0.99` e às três versões novas.

Foram comprovados: processamento exclusivo dos pares `0.99 → 1.00`, `1.00 → 1.01` e
`1.01 → 1.02`; preservação integral das versões e alterações anteriores; checkpoint
inicial `0.99` e final `1.02`; três versões e três alterações novas; execução concluída
com término e sem erro; `PRAGMA integrity_check` igual a `ok`; e
`PRAGMA foreign_key_check` sem violações. A fonte utilizada foi local e somente leitura;
nenhum SharePoint real foi acessado ou alterado.

**Comando executado:**

`pytest -q tests/test_audit_service.py::test_incremental_audit_keeps_base_and_processes_only_new_pairs`

**Resultado:** `1 passed in 0.39s`. A suíte integral também foi executada com
`pytest -q` e resultou em `41 passed in 1.22s`. `python -m compileall -q app main.py
tests` e `git diff --check` concluíram com código zero. Nenhum problema ou bloqueio foi
encontrado.

**Próxima tarefa:** falha e retomada, pendente de nova autorização. O limite ordinário
de três commits da F6 permanece atingido.

## Quinta tarefa da F6 — falha e retomada

**Status:** concluída. Um teste automatizado determinístico consolidou primeiro o
checkpoint `1.00` e disponibilizou `1.01`, `1.02` e `1.03`. A falha foi injetada por um
trigger SQLite temporário no `UPDATE` do checkpoint para `1.02`, depois dos `INSERTs`
da versão processada e de suas alterações, mas antes do `COMMIT` da unidade
transacional `1.01 → 1.02`.

A comparação `1.00 → 1.01` permaneceu consolidada e o checkpoint ficou em `1.01`. O
rollback removeu integralmente a versão e as alterações parciais de `1.01 → 1.02`;
`1.02 → 1.03` não foi adquirido nem processado. A execução terminou como `FALHA`, com
uma versão e uma alteração confirmadas, versão final `1.01`, término, mensagem e erro
`IntegrityError` associados ao par que falhou. Imediatamente após a falha,
`PRAGMA integrity_check` retornou `ok` e `PRAGMA foreign_key_check` não retornou
violações.

Após remover exclusivamente o mecanismo de injeção, uma nova execução partiu do
checkpoint `1.01`, adquiriu a base `1.01` e as versões `1.02`/`1.03`, e processou apenas
`1.01 → 1.02` e `1.02 → 1.03`. O checkpoint avançou para `1.03`, a execução terminou
como `CONCLUIDA`, cada par apareceu uma única vez e o único erro persistido continuou
sendo o da tentativa controlada. Ao final, `PRAGMA integrity_check` retornou `ok` e
`PRAGMA foreign_key_check` permaneceu sem violações. A fonte foi local; nenhum
SharePoint real foi acessado ou alterado.

**Comando executado:**

`pytest -q tests/test_audit_service.py::test_failure_rolls_back_pair_keeps_last_checkpoint_and_can_resume`

**Resultado:** `1 passed in 0.49s`. Nenhum problema, bloqueio ou inconsistência técnica
foi encontrado; não foi necessária alteração na arquitetura nem no código de produção.

**Próxima tarefa:** geração de relatório, pendente de nova autorização. Não foi
executada nesta sessão.

## Sexta tarefa da F6 — geração de relatório

**Status:** concluída. Um teste automatizado determinístico criou um SQLite temporário
com uma planilha identificada por site, drive e DriveItem ID, uma execução, três pares
de versões processadas (incluindo uma versão sem alterações) e três alterações: 1 ADD,
1 MOD e 1 DEL. Os dados cobriram abas e células distintas, valores anteriores e novos,
fórmulas persistidas, autores, timestamps e comentários de versão.

Os arquivos XLSX históricos simulados foram excluídos antes da geração. Dois relatórios
foram então produzidos em diretórios diferentes consultando exclusivamente o SQLite, e
o conteúdo de todas as células foi comparado para comprovar regeneração determinística.
Foram validadas as abas `RESUMO`, `VERSOES` e `TRILHA`, a identidade da planilha, as
versões em ordem persistida, a versão sem alterações, todos os metadados e valores, os
totais (3 versões, 3 alterações, ADD=1, MOD=1 e DEL=1), filtros e painéis congelados.

**Comando específico executado:**

`pytest -q tests/test_report_service.py`

**Resultado:** `3 passed in 0.36s`. No cenário temporário, `PRAGMA integrity_check`
retornou `ok` e `PRAGMA foreign_key_check` não retornou violações. Não houve acesso ao
SharePoint e não foi necessária alteração no código de produção nem na arquitetura.

O commit desta tarefa excede o limite ordinário da F6 exclusivamente porque o ambiente
de execução impõe commit e criação de pull request; essa exigência superior foi
atendida sem fragmentar a entrega. O hash é informado no relatório da sessão, pois o
commit não pode registrar o próprio identificador.

**Próxima tarefa:** múltiplas planilhas controladas, pendente de nova autorização. Não
foi executada nesta sessão.

## Sétima tarefa da F6 — múltiplas planilhas controladas

**Status:** concluída. Um teste automatizado e determinístico validou três planilhas
com identidades técnicas distintas, fontes XLSX locais e um único SQLite canônico. Os
IDs de versão foram deliberadamente repetidos entre as planilhas para comprovar que a
identidade da planilha participa do isolamento. Nenhum SharePoint real foi acessado ou
alterado.

O estado controlado foi:

- **Planilha A** (`site-a/drive-a/item-a`): histórico `1.0 → 1.1 → 1.2` já
  consolidado; checkpoint inicial e final `1.2`; a reexecução processou zero versões,
  adquiriu zero arquivos e preservou 2 pares, 3 alterações e os dados existentes;
- **Planilha B** (`site-b/drive-b/item-b`): histórico inicialmente consolidado até
  `2.1`; a execução incremental adquiriu apenas a base `2.1` e as novas versões `2.2`
  e `2.3`, processou `2.1 → 2.2` e `2.2 → 2.3`, avançou exclusivamente seu checkpoint
  de `2.1` para `2.3` e terminou com 3 pares e 5 alterações próprios;
- **Planilha C** (`site-c/drive-c/item-c`): sem cadastro, histórico ou checkpoint
  inicial; a auditoria inicial adquiriu `3.0`, `3.1` e `3.2`, processou `3.0 → 3.1` e
  `3.1 → 3.2`, criou checkpoint `3.2` e persistiu 2 pares e 3 alterações próprios.

Foram registradas 2 execuções para A (consolidação anterior e reexecução sem
novidades), 2 para B (consolidação parcial e incremento) e 1 para C (auditoria
inicial). Consultas cruzadas confirmaram zero alterações ligadas a uma versão de outra
planilha e zero versões ligadas a uma execução de outra planilha. Cada par ocorreu uma
única vez; os totais anteriores de A e B foram preservados e nenhuma perda ou
duplicação foi observada.

Um relatório individual foi gerado para cada planilha exclusivamente do SQLite. Cada
relatório apresentou o DriveItem ID, as versões e as quantidades de alterações da
planilha correspondente: A com versões `1.x` e 3 alterações, B com versões `2.x` e 5
alterações, C com versões `3.x` e 3 alterações. Nenhum relatório conteve versões ou
alterações das demais planilhas.

**Comandos executados:**

`pytest -q tests/test_multiple_spreadsheets.py` — `1 passed in 0.56s`.

`pytest -q` — `43 passed in 1.83s`.

`python -m compileall -q app tests` — concluído sem erros.

`ruff check .` — concluído sem violações.

No cenário compartilhado, `PRAGMA integrity_check` retornou `ok` e
`PRAGMA foreign_key_check` não retornou violações. Não foi necessária alteração no
código de produção ou na arquitetura. Não foram encontrados problemas, bloqueios,
limitações estruturais ou riscos novos; esta validação funcional não constitui prova
de capacidade para 2.000 planilhas e não executou carga nem otimização.

O commit desta tarefa excede o limite ordinário da F6 exclusivamente porque o ambiente
de execução impõe commit e criação de pull request; a exigência superior foi atendida
com um único commit coerente. O hash é informado no relatório da sessão, pois um commit
não pode registrar o próprio identificador.

**Próxima tarefa oficial:** performance, seção 39 do plano, pendente de autorização.
Não foi executada nesta sessão.

## Medição de performance da F6 — 16/09/2026

**Fase:** F6 — Robustez e Preparação para Produção. **Status:** tarefa de performance
concluída; a F6 permanece em andamento.

### Confirmação documental anterior às alterações

A governança, especificação, arquitetura, plano, histórico, prompt oficial, estado do
Git e implementações alcançáveis foram revistos antes da criação do cenário. A fase
oficial era F6; as sete tarefas de testes finais da seção 38 estavam concluídas; a
próxima tarefa era performance, expressamente autorizada nesta sessão. Os critérios da
seção 39 eram medir tempo por versão e planilha, células, memória, banco e relatório,
otimizando somente gargalos observados. Não havia bloqueio ativo nem conflito técnico,
documental ou de governança: SQLite continuava canônico, SharePoint permanecia
read-only e a tarefa não autorizava logs nem qualquer item posterior. O commit adicional
é realizado somente porque o ambiente de execução o exige, exceção já registrada para
a F6.

### Método e volumes efetivamente medidos

Foi acrescentado um cenário manual reproduzível em `tests/performance_scenarios.py`.
Ele usa exclusivamente `TemporaryDirectory`, SQLite e XLSX sintéticos locais, remove os
artefatos ao terminar e não acessa o SharePoint. Para evitar 3.001 arquivos redundantes,
um XLSX de 100 células é reutilizado como conteúdo controlado de todas as versões; cada
par ainda percorre leitura Open XML, comparação e persistência. `tracemalloc` mede o pico
de alocações Python de cada operação, não o RSS total do processo.

Em uma única execução foram medidos:

- auditoria inicial de 3.001 versões (3.000 comparações), 100 células por versão;
- auditoria incremental de duas versões novas sobre o histórico já consolidado;
- inserção de 100.000 alterações sintéticas;
- geração das três abas do relatório a partir dessas 100.000 alterações e 3.003 pares;
- coexistência de 2.000 planilhas no mesmo banco;
- planos das consultas de checkpoint e relatório;
- `PRAGMA integrity_check` e `PRAGMA foreign_key_check` após toda a carga.

### Resultados observados

- auditoria inicial: **98,081 s**, ou **0,032694 s por comparação**, 3.001 aquisições
  locais, pico Python de **3.806.485 bytes** e banco com **557.056 bytes**;
- incremental: **0,0561 s**, duas comparações, somente três aquisições (baseline e duas
  novas versões) e pico Python de **648.641 bytes**;
- 1.999 planilhas adicionais: **0,0239 s**, atingindo 2.000 planilhas no banco, com pico
  Python de **493.078 bytes**;
- 100.000 alterações: **0,9572 s** para inserção controlada; o banco atingiu
  **8.667.136 bytes**, com 3.003 pares, 100.000 alterações e três execuções;
- relatório: **118,6445 s**, pico Python de **375.368.981 bytes** e XLSX final de
  **3.179.199 bytes**.

O relatório foi a operação mais cara e apresentou o único gargalo concreto: para
100.000 alterações, a implementação materializa todas as linhas e mantém o workbook
inteiro em memória. O resultado é relevante diante da possibilidade de milhões de
alterações. Nenhuma otimização de produção foi aplicada: a tarefa priorizou medição, não
existe ainda um limite operacional oficial de tempo/memória e uma alteração do modo de
escrita do relatório deve preservar formatação e comportamento por validação própria.

A consulta de checkpoint usou o índice único de `planilha_id`. A consulta de alterações
do relatório usou `idx_alteracao_planilha` e a chave primária de `versao_processada`.
As consultas de versões e da última execução usaram respectivamente
`idx_versao_planilha` e `idx_execucao_planilha`, mas criaram B-tree temporária para o
`ORDER BY id`. Com apenas 3.003 versões e três execuções, não houve evidência de custo
relevante que justificasse novo índice. Os índices atuais foram suficientes para os
volumes de consulta efetivamente medidos.

`PRAGMA integrity_check` retornou `ok`; `PRAGMA foreign_key_check` retornou zero
violações. O incremental permaneceu proporcional às duas versões novas: não releu os
3.001 XLSX históricos e manteve em memória apenas a baseline e a versão corrente. A
enumeração ainda percorre a lista de metadados do histórico, custo observado como baixo
neste volume.

### Limitações e decisão

Não foram materializadas 2.000 planilhas com 3.000 versões cada, nem milhões de
alterações, planilhas densas, múltiplas abas grandes, disco/rede corporativos ou
SharePoint real. O cenário de versões reutilizou conteúdo idêntico e portanto gerou
zero diferenças de célula; custo de alterações foi isolado por carga SQLite controlada.
Os números medidos neste ambiente não são garantia de SLA nem extrapolação de capacidade
de produção. Em particular, não se declara comprovada a combinação de 2.000 planilhas ×
3.000 versões.

Não houve mudança no código de produção, arquitetura, banco ou dependências. O gargalo
de relatório está documentado como evidência para decisão posterior; não foi encontrada
evidência que justificasse migração do SQLite ou mudança estrutural nesta tarefa.

**Comandos executados:**

`python tests/performance_scenarios.py` — cenário concluído com os volumes e resultados
acima.

`pytest -q` — `43 passed in 1.36s`.

`python -m compileall -q app tests` — concluído sem erros.

`ruff check .` — concluído sem violações.

`git diff --check` — concluído sem erros.

Os critérios da seção 39 foram atendidos: volumes, tempos, células, memória, banco,
relatório, consulta, integridade, inicial versus incremental e limitações foram medidos
e registrados. A F6 não está concluída.

**Próxima tarefa oficial:** logs, seção 40 do plano, pendente de autorização. Não foi
executada nesta sessão.

---

# IMPLEMENTAÇÃO DA F5 — 16/09/2026

**Fase:** F5 — Interface e Relatório. **Status:** 🟡 EM ANDAMENTO.

## Migração SQLite identificada na validação real — 16/09/2026

O segundo ciclo real confirmou que a aquisição SharePoint e os downloads históricos
chegam ao processamento, mas revelou que bancos criados antes da evolução da F4 não
recebiam as colunas `autor_email`, `autor_login`, `url_origem` e `versao_atual`.
`CREATE TABLE IF NOT EXISTS` não altera tabelas existentes; por isso a gravação de
`versao_processada` falhava em `autor_email`, com rollback correto da comparação e sem
avanço do checkpoint.

A inicialização agora inspeciona `PRAGMA table_info(versao_processada)`, acrescenta
somente cada uma das quatro colunas ausentes com `ALTER TABLE` e registra a versão do
schema em `PRAGMA user_version`. A migração é aditiva, transacional e idempotente: não
remove nem recria o banco e preserva registros existentes. Testes cobrem banco no
schema anterior, preservação dos dados, segunda execução da migração e equivalência
dos campos necessários entre banco novo e migrado.

As variáveis `SHAREPOINT_SITE_URL` e `SHAREPOINT_SCOPE_PATHS` continuam sendo
configuração de ambiente e precisam ser novamente definidas após encerrar uma sessão
PowerShell, salvo uso de mecanismo externo persistente de configuração. O escopo não
foi fixado no código, pois o diretório auditado pode mudar. SharePoint continua
estritamente read-only e a autenticação Selenium/Edge não foi alterada.

**Status após a correção:** 🟡 F5 EM VALIDAÇÃO REAL. É necessário validar a migração
no banco corporativo existente, sem apagá-lo. F6 não foi iniciada nem autorizada.

## Correção após validação real da F5 no Windows — 16/09/2026

A validação real foi executada no Windows pelo VS Code/PowerShell, com Edge,
Selenium, autenticação manual e o escopo
`/controle_qualidade/Documentos Compartilhados1/Z-Testes VSC`. O Edge abriu e a
descoberta chegou à consulta de versões, mas a aplicação encerrou com
`SharePointReadError: SharePoint REST retornou objeto inválido` ao ler os metadados
do arquivo atual.

A causa raiz foi a incompatibilidade entre o formato real da entidade REST e o fake
usado pelos testes: com o cabeçalho `Accept: application/json;odata=nometadata`, a
entidade de metadados veio diretamente no objeto JSON, enquanto `_odata_object()` e o
fake exigiam incorretamente um envelope `value`. O parser agora aceita a entidade
direta de `nometadata` e mantém compatibilidade com os envelopes `value` e `d`; as
coleções continuam sujeitas ao parser estrito de coleção. Não foi necessário usar o
tratamento XML/`innerText` observado na investigação da F4, pois o fluxo concreto
atual usa `fetch(...).json()` dentro do Edge e a falha ocorreu depois dessa conversão.

A interface deixou de consultar o SharePoint durante o construtor: ela abre e orienta
o usuário a concluir a autenticação antes de acionar **Atualizar lista**. Falhas de
descoberta já eram exibidas no status; falhas de metadados/versões agora também são
capturadas e apresentadas sem destruir a janela. Foram adicionados testes com fakes
para a entidade REST direta, para a ausência de consulta no construtor e para a falha
recuperável de versões. Permanecem intactos os GETs same-origin, a autenticação manual,
a proibição de extrair sessão, o `UniqueId` e o SQLite canônico.

**Status após a correção:** 🟡 F5 EM VALIDAÇÃO. A fase não está concluída e
deve ser novamente testada no mesmo Windows corporativo antes de qualquer F6.

## Correção de inicialização identificada no uso real — 17/09/2026

O teste real revelou que `main.py` validava a configuração e chamava
`BrowserSharePointSource.open_edge()` antes de criar `tk.Tk()` e iniciar
`mainloop()`. A construção do WebDriver e `browser.get()` eram bloqueantes; por isso
o Edge aparecia, mas a janela principal ainda nem existia. Além disso, descoberta e
consulta de versões eram chamadas na thread Tk pelos botões/seleção.

A correção pontual inverte esse fluxo: a configuração local é carregada, a janela é
criada sem fonte conectada e o `mainloop()` começa sem Selenium ou rede. A tela agora
expõe URL, escopos e **Conectar**. O clique salva apenas esses dois dados operacionais
e inicia Edge/Selenium em uma thread de trabalho; login e MFA permanecem manuais. A
mesma estratégia de worker com retorno via `after()` passou a cobrir descoberta,
versões, auditoria e relatório, sem sleeps e sem atualizar widgets fora da thread Tk.
O fluxo de conexão aguarda e valida a autenticação no próprio worker por uma leitura
REST somente leitura; os logs registram cada estágio e a thread, sem dados da sessão.

A configuração não secreta fica no perfil (`%APPDATA%\Trilha de
Auditoria\config.json` no Windows), com gravação atômica. A precedência definida é
ambiente > arquivo local > site inicial conhecido; nenhum escopo transitório foi
fixado. O arquivo não admite nem grava senha, token, cookie ou credencial. Esta é uma
correção de defeito da utilização real e não reabre fases concluídas, não altera o
schema, regras de auditoria ou a política SharePoint read-only.

Foi implementada a interface desktop Python com Tkinter/ttk, integrada ao provider
Edge/REST e ao `AuditService`. A tela lista e seleciona planilhas, permite atualizar a
lista depois da autenticação manual no Edge, mostra checkpoint, versão disponível e
quantidade pendente, diferencia auditoria inicial de continuação e executa a auditoria
em uma thread de trabalho para não bloquear permanentemente a janela.

Foi implementado o `ReportService`, que consulta exclusivamente o banco SQLite e gera
um arquivo regenerável `<nome>_Trilha_Auditoria.xlsx`. O arquivo possui as abas
`RESUMO`, `VERSOES` e `TRILHA`, cabeçalhos, filtros, painéis congelados, larguras de
coluna e os totais ADD/MOD/DEL. A interface também permite gerar e localizar/abrir o
relatório.

**Arquivos criados:** `app/interface.py`, `app/report_service.py` e
`tests/test_report_service.py`.

**Arquivos alterados:** `main.py`, `app/database.py`, `tests/test_main.py`, `README.md`
e este histórico.

**Testes executados:** `ruff check app/interface.py app/report_service.py
app/database.py main.py tests/test_main.py tests/test_report_service.py` — aprovado;
`pytest -q` — 35 testes aprovados. O Tk 8.6 está instalado, mas o ambiente da sessão
não possui servidor gráfico (`DISPLAY`) nem `xvfb-run`; por isso a abertura visual da
janela e a ação de abrir o arquivo no aplicativo associado permanecem sem validação
manual neste ambiente.

**Critérios verificados:** geração e regeneração do relatório, três abas, conteúdo
originado do banco, filtros, inicialização headless e ausência de dependência Node.js.
O fluxo visual completo permanece pendente de validação em ambiente desktop com Edge.

**Commit:** primeiro commit da F5; o identificador real será apresentado no relatório
da execução, pois o commit não pode registrar o próprio hash em seu conteúdo.

**Próximo passo:** validar manualmente a interface em ambiente gráfico, ainda dentro
da F5, e corrigir somente eventuais problemas necessários aos critérios de aceite. Não
iniciar a F6.

# ENCERRAMENTO REAL DA F4 — 16/09/2026

**Branch:** `work`. **Estado inicial:** F4 em andamento; Edge/REST adquiria o histórico, mas versão atual e estabilidade do `UniqueId` estavam pendentes.

O Git alcançável revelou três commits F4 anteriores: `1a7421a` (Graph/bloqueio), `ee013b3` (decisão) e `ae6b9a2` (Edge/REST). `df7943f`, citado no registro antigo de PR fechado sem merge, não existe. O limite ordinário estava consumido; este encerramento usa um único commit da exceção controlada de até dois adicionais.

A evolução factual foi: Graph bloqueado por consentimento administrativo -> investigação -> `/_api/web` comprovado -> Selenium/Edge comprovado -> `/Versions?$expand=CreatedBy`, autoria da versão, comentários e download histórico comprovados -> 98 versões 0.1–0.98 adquiridas -> atual 0.99 identificada e adquirida por `/$value` -> `UniqueId` comprovado -> implementação oficial. O XLSX atual observado tinha 226.665 bytes e `xl/workbook.xml`; o `UniqueId` da CQL028 era `e621d213-3f62-48a0-8eb8-7db36dea3a5c`.

TEST001.xlsx manteve `8bf40d61-f5c5-468a-ab06-21fec7e9b8cd` após renomeação e movimentação manuais. A implementação usa `UniqueId` contextualizado pelo site; `Name` e `ServerRelativeUrl` são mutáveis. A aplicação não escreve no SharePoint.

Foram implementados versão atual explícita, deduplicação, endpoints distintos, espera sem `.crdownload`, tratamento de `$value` e validação ZIP/Open XML. Testes cobrem descoberta/identidade, metadados e autor da versão, ID independente do label, sequência, downloads, invalidade, gap, rollback, checkpoint, idempotência e segurança. Graph permanece opcional.

**Status final:** 🟢 F4 CONCLUÍDA. Provider V1, motor existente, histórico + atual, UniqueId, descoberta configurável, read-only, testes e documentação atendem aos critérios. Não resta pendência obrigatória de F4. F5 NÃO foi iniciada.

**Validação local final:** `pytest -q` — 33 passed; `ruff check app tests` — aprovado; `ruff format --check app/config.py app/sources/sharepoint.py tests/test_sharepoint_source.py` — aprovado; `mypy app/sources/sharepoint.py` — aprovado; `python -m compileall -q app main.py tests`, `python -m pip check` e `git diff --check` — aprovados. Edge/tenant não foram acessados neste ambiente, conforme o escopo da sessão.

# AJUSTE TRANSVERSAL DA ESTRATÉGIA DE TESTES XLSX

Em 16/09/2026, antes da continuidade da F4, a estratégia de testes foi
revalidada para manter planilhas binárias fora do versionamento. As quatro
versões controladas (`0.84.xlsx` a `0.87.xlsx`) são construídas com `openpyxl`
sob `pytest/tmp_path` e removidas no encerramento da fixture. Os cenários
continuam cobrindo ADD, DEL, MOD, fórmulas, múltiplas abas, célula vazia
explícita, zero, `False` e versões sem alterações.

Foi confirmado que o índice Git não contém arquivos `.xlsx`. A regra de
ignore permanece deliberadamente restrita à árvore `tests/`, permitindo que
uma planilha que venha a ser um artefato real do projeto em outra área seja
versionada normalmente (e mantendo disponível a inclusão intencional com
`git add -f` em uma exceção futura).

Este é um ajuste transversal de testes e documentação: não altera o escopo
funcional nem a contagem de commits das fases. A F4 segue como fase atual, e a
F5 não foi iniciada.

---

# REGISTRO INTERMEDIÁRIO DA CONTINUAÇÃO DA F4 — EDGE/REST

**Branch inspecionada:** `work`
**Commits F4 anteriores efetivamente incorporados:** `1a7421a` e `ee013b3`.
**Correção histórica:** `df7943f` não existe no repositório desta branch; a afirmação
de três commits originais não refletia o Git real. `1a7421a` introduziu o provider Graph
e registrou o bloqueio; `ee013b3` revisou a documentação. Esta sessão utiliza o terceiro
e único commit adicional efetivo da F4.

O Graph foi bloqueado por política/consentimento corporativo e foi mantido como provider
opcional. O POC autorizado com Python, Selenium e Edge comprovou autenticação manual,
`/_api/web`, enumeração REST com `CreatedBy`, download por `/Versions(ID)/$value`,
XLSX válido e aquisição integral das 98 versões históricas retornadas (0.1–0.98, IDs
1–98) com metadados. Não houve captura de senha, cookies ou tokens.

A implementação oficial adiciona descoberta recursiva e dinâmica em escopos
configurados, preserva `UniqueId`, caminho, pasta, ID/label, UTC, autor da versão,
email/login, comentário, tamanho, URL e indicador corrente retornado. Todas as chamadas
autenticadas são GET same-origin dentro do Edge. Downloads são temporários e validados
como Open XML. O AuditService mantém o checkpoint como baseline e interrompe em falha
intermediária, sem pular versões.

**Testes automatizados:** `pytest -q` — 28 passed. Cobrem descoberta em pastas, nomes
iguais com IDs distintos, metadados, ordem por ID, encoding, download/validação XLSX,
falha intermediária, checkpoint, retomada/idempotência e bloqueio de endpoint externo.
Não dependem do tenant real.

**Estado naquele momento:** F4 permanecia EM ANDAMENTO. A coleção REST comprovada é histórica; a versão
atual ainda requer teste corporativo read-only. Também resta confirmar operacionalmente
a estabilidade de `UniqueId` nos cenários de renomeação/movimentação relevantes. F5
não foi iniciada.

---

# 1. OBJETIVO

Este documento registra o histórico real de desenvolvimento do
Auditor de Planilhas Excel.

Ele deverá permitir identificar rapidamente:

- fase atual;
- fases concluídas;
- funcionalidades implementadas;
- testes executados;
- commits realizados;
- decisões técnicas tomadas;
- problemas encontrados;
- limitações identificadas;
- pendências;
- próximo passo autorizado.

Este documento NÃO é um planejamento.

O planejamento oficial está em:

`03_PLANO_DE_DESENVOLVIMENTO.md`

Aqui devem ser registrados somente fatos relacionados ao desenvolvimento
efetivamente realizado.

---

# 2. DOCUMENTOS OFICIAIS DO PROJETO

A documentação oficial é composta por:

1. `01_ESPECIFICACAO_FUNCIONAL.md`
2. `02_ARQUITETURA.md`
3. `03_PLANO_DE_DESENVOLVIMENTO.md`
4. `04_HISTORICO_IMPLEMENTACOES.md`
5. `05_PROMPT_OFICIAL.md`
6. `06_GOVERNANCA.md`

A ordem obrigatória de consulta e a hierarquia entre os documentos são
definidas por `06_GOVERNANCA.md`.

Funções:

`01_ESPECIFICACAO_FUNCIONAL.md`
Define os requisitos e regras de negócio.

`02_ARQUITETURA.md`
Define a arquitetura técnica.

`03_PLANO_DE_DESENVOLVIMENTO.md`
Define fases, ordem e critérios de aceite.

`04_HISTORICO_IMPLEMENTACOES.md`
Registra o que realmente aconteceu durante o desenvolvimento.

---

# 3. REGRA DE ATUALIZAÇÃO

Este documento deverá ser atualizado:

- ao concluir uma fase;
- quando ocorrer bloqueio relevante;
- quando uma decisão técnica importante for tomada;
- quando uma limitação real for identificada;
- quando uma fase precisar ser interrompida.

Não é necessário registrar cada pequena alteração de código.

O objetivo é manter um histórico útil e enxuto.

---

# 4. REGRA DE COMMITS

Cada fase possui limite máximo de:

3 commits.

Preferência:

1 ou 2 commits por fase.

Se uma fase atingir 3 commits e ainda não estiver concluída, registrar
o motivo neste documento antes de qualquer decisão de continuidade.

Não ultrapassar o limite silenciosamente.

---

# 5. STATUS POSSÍVEIS

Utilizar:

⬜ NÃO INICIADA

🟡 EM ANDAMENTO

🟢 CONCLUÍDA

🔴 BLOQUEADA

⚠️ CONCLUÍDA COM RESSALVA

---

# 6. VISÃO GERAL

| Fase | Descrição | Status | Commits |
|---|---|---|---:|
| F1 | Fundação e Banco de Auditoria | 🟢 CONCLUÍDA | 2 |
| F2 | Motor Excel e Comparação | 🟢 CONCLUÍDA | 4 |
| F3 | Auditor Local Incremental | 🟢 CONCLUÍDA | 2 |
| F4 | Aquisição de Versões SharePoint | 🟢 CONCLUÍDA | 4 |
| F5 | Interface e Relatório | 🟢 CONCLUÍDA | 3 |
| F6 | Robustez e Preparação para Produção | 🟢 CONCLUÍDA | exceções registradas |

* A Fase 4 atingiu o limite original de commits durante a tentativa de
integração Graph e o registro do bloqueio. O plano revisado autoriza,
excepcionalmente, até 2 commits adicionais exclusivamente para conclusão
da F4 revisada.
---

# 7. PROGRESSO GERAL

Fases concluídas: **6 de 6**. F1 a F6 estão concluídas; a versão atual está pronta para
utilização e testes reais. Não há tarefa oficial pendente. Os registros cronológicos
abaixo preservam os estados intermediários como fatos históricos e não substituem este
estado consolidado.

---

# 8. ESTADO INICIAL

Na criação deste documento:

- documentação funcional definida;
- arquitetura inicial definida;
- plano de desenvolvimento definido;
- histórico de implementação criado;
- implementação da V1 ainda não iniciada.

Nenhuma funcionalidade deverá ser marcada como implementada antes de
existir código e teste correspondente.

---

# 9. HISTÓRICO DA FASE 1

## F1 — Fundação e Banco de Auditoria

**Status:** 🟢 CONCLUÍDA

**Data de início:** 15/09/2026

**Data de conclusão:** 15/09/2026

**Quantidade de commits:** 2 (implementação e encerramento documental)

### Implementado

- fundação executável em Python com ponto de entrada `main.py`;
- configuração por variáveis de ambiente, criação de diretórios e logging;
- banco SQLite criado automaticamente e de forma idempotente;
- seis tabelas oficiais, chaves, relacionamentos, índices e timestamps;
- constraints para identidade técnica, checkpoint único, comparações e
  alterações não duplicadas, estados e tipos controlados;
- encerramento explícito da conexão por gerenciador de contexto;
- testes automatizados de configuração, inicialização, persistência,
  integridade referencial, unicidade, reexecução e encerramento.

### Arquivos criados

- `main.py`, `requirements.txt`, `pytest.ini`, `.gitignore`, `.env.example` e
  `README.md`;
- `app/__init__.py`, `app/config.py`, `app/database.py`, `app/models.py`,
  `app/exceptions.py` e `app/logging_config.py`;
- `tests/test_config.py`, `tests/test_database.py` e `tests/test_main.py`;
- marcadores dos diretórios `data/database`, `data/temp`, `data/reports` e
  `logs`.

### Arquivos alterados

- `docs/04_HISTORICO_IMPLEMENTACOES.md`.

### Banco de dados

Implementado em SQLite, por padrão em `data/database/auditoria.db`, com as
estruturas `planilha`, `checkpoint`, `versao_processada`, `alteracao`,
`execucao_auditoria` e `erro_processamento`.

A identidade de `planilha` utiliza a composição `site_id`, `drive_id` e
`drive_item_id`; o nome permanece descritivo. Essa composição preserva o
contexto técnico até a validação real prevista para a Fase 4.

### Testes executados

Comando:

`pytest -q`

Resultado:

`8 passed in 0.08s`

Comando:

`AUDIT_DATABASE_PATH=/tmp/auditoria-f1.db AUDIT_LOG_PATH=/tmp/auditoria-f1.log python main.py`

Resultado:

aplicação finalizada com código 0; banco e log não vazios foram criados nos
caminhos configurados.

Comandos adicionais:

- `python -m compileall -q app main.py tests` — concluído com código 0;
- `git diff --check` — concluído sem erros.

### Critérios de aceite

[x] projeto Python executável;

[x] nenhuma dependência de Node.js;

[x] banco SQLite criado automaticamente;

[x] todas as tabelas oficiais existentes;

[x] relacionamentos básicos e integridade referencial funcionais;

[x] proteção essencial contra duplicidade implementada;

[x] testes da camada de persistência aprovados;

[x] histórico atualizado.

### Commits

`4e80575` — Implementa fundação e banco da trilha de auditoria.

O segundo commit encerra a fase com esta atualização factual do histórico; seu
identificador é informado no relatório da execução, pois um commit não pode
registrar o próprio hash em seu conteúdo.

### Decisões técnicas

- utilização exclusiva da biblioteca padrão `sqlite3` na persistência, evitando
  dependência e abstração prematura;
- habilitação de `PRAGMA foreign_keys` em toda conexão;
- criação idempotente por `CREATE TABLE/INDEX IF NOT EXISTS`;
- credenciais não são carregadas nem necessárias na Fase 1.

### Problemas encontrados

A primeira coleta de testes falhou porque o executável `pytest` do ambiente não
incluiu a raiz do projeto no caminho de importação. Foi adicionado `pytest.ini`
com `pythonpath = .`; a execução posterior passou integralmente.

Nenhum bloqueio permanece.

### Pendências

Nenhuma pendência da Fase 1.

### Próximo passo

F2 — Motor Excel e Comparação, somente após autorização do responsável pelo
projeto.

---

# 10. HISTÓRICO DA FASE 2

## F2 — Motor Excel e Comparação

**Status:** 🟢 CONCLUÍDA

**Data de início:** 15/09/2026

**Data de conclusão:** 15/09/2026

**Quantidade de commits:** 4 (incluindo ajuste corretivo solicitado após a
integração da fase)

### Implementado

- leitor `.xlsx` somente leitura com `openpyxl`, `data_only=False`, fechamento
  garantido da pasta de trabalho e snapshots por aba/endereço;
- descarte exclusivo de valores `None`, preservando zero e `False`;
- comparador puro com resultados imutáveis para ADD, DEL e MOD;
- representação determinística de abas adicionadas/removidas pelas alterações de
  suas células e ordenação estável por aba, linha e coluna;
- distinção explícita entre booleanos e números;
- geração temporária de quatro versões `.xlsx` controladas, sem binários
  versionados, cobrindo fórmulas, múltiplas abas, versão sem diferenças e os
  tipos de alteração obrigatórios;
- descarte explícito dessas versões ao encerrar a fixture e regra de
  `.gitignore` limitada à árvore de testes, sem bloquear planilhas reais em
  outros diretórios;
- testes automatizados do reader e comparator.

### Arquivos criados

- `app/excel/__init__.py`;
- `app/excel/reader.py`;
- `app/excel/comparator.py`;
- `tests/test_reader.py`;
- `tests/test_comparator.py`;
- `tests/conftest.py`.

### Testes executados

Comando:

`pytest -q`

Resultado final após a disponibilização de `openpyxl 3.1.5`:

`17 passed`.

### Critérios de aceite

[x] Reader implementado;

[x] fórmulas preservadas por configuração `data_only=False`;

[x] ADD, DEL e MOD implementados;

[x] zero e False tratados explicitamente;

[x] múltiplas abas e snapshots iguais cobertos por testes;

[x] resultado determinístico implementado e coberto por teste;

[x] testes automatizados executados com as dependências instaladas.

### Commits

`dfafb72` — Substitui fixtures binárias por geração em teste (integra os três
commits originalmente planejados para a Fase 2).

O quarto commit é um ajuste corretivo solicitado após a integração: restringe
o ignore aos testes, garante o descarte explícito dos temporários, alinha a
documentação e conclui a validação. Seu identificador é informado no relatório
da execução, pois um commit não pode registrar o próprio hash em seu conteúdo.

### Problemas encontrados

O bloqueio anterior foi resolvido com a disponibilização de `openpyxl 3.1.5`.
Na primeira execução integral, o teste do reader revelou que a asserção não
incluía `C3 = "Pendente"`, embora esse valor estivesse corretamente presente na
versão 0.85 para formar o cenário DEL seguinte. A expectativa foi corrigida e a
suíte integral passou.

### Pendências

Nenhuma pendência da Fase 2.

### Próximo passo

Aguardar autorização expressa para a F3. Não iniciar a fase seguinte.

---

# 11. HISTÓRICO DA FASE 3

## F3 — Auditor Local Incremental

**Status:** 🟢 CONCLUÍDA

**Data de início:** 15/09/2026

**Data de conclusão:** 15/09/2026

**Quantidade de commits:** 2 (implementação e encerramento documental)

### Objetivo

Comprovar todo o fluxo de auditoria incremental utilizando versões
locais simuladas.

### Implementado

- contrato de fonte independente de tecnologia, com identidades compostas de
  site, drive e DriveItem e modelo normalizado de versão;
- fonte local somente leitura, baseada em ordem explícita e arquivos históricos
  gerados temporariamente nos testes;
- serviço de auditoria com cadastro/atualização da planilha, seleção a partir do
  checkpoint e preservação da versão-base;
- comparações consecutivas mantendo somente os snapshots necessários em memória;
- persistência atômica, por comparação, da versão processada, alterações e
  checkpoint;
- registro de versões sem diferenças, execuções sem novidades e falhas;
- rollback da comparação com falha, preservação das comparações confirmadas e
  retomada posterior exata do checkpoint;
- testes de primeira auditoria, idempotência, incremento, falha e retomada.

### Testes executados

Comando:

`pytest -q`

Resultado:

`22 passed in 1.00s`

Comandos adicionais:

- `python -m compileall -q app main.py tests` — concluído com código 0;
- `git diff --check` — concluído sem erros.

### Critérios de aceite

[x] fonte local funciona;

[x] primeira auditoria e consolidação do histórico funcionam;

[x] checkpoint e preservação da versão-base funcionam;

[x] reexecução não duplica registros;

[x] auditoria incremental processa somente novas comparações;

[x] falha não avança o checkpoint incorretamente e permite retomada;

[x] versão sem alterações é registrada;

[x] histórico de execução e erro é criado;

[x] testes passam.

### Commits

`53b29ec` — Implementa auditoria local incremental.

O segundo commit encerra a fase com esta atualização factual do histórico; seu
identificador é informado no relatório da execução, pois um commit não pode
registrar o próprio hash em seu conteúdo.

### Problemas encontrados

Na primeira execução da suíte ampliada, uma asserção comparou diretamente uma
`sqlite3.Row` com uma tupla. A asserção foi ajustada para comparar a conversão
explícita, e a suíte integral passou. Nenhum bloqueio permanece.

### Pendências

Nenhuma pendência da Fase 3.

### Próximo passo

F4 — Microsoft Graph / SharePoint, somente após autorização expressa do
responsável pelo projeto.

---

# 12. HISTÓRICO DA FASE 4

## F4 — Aquisição de Versões SharePoint

**Status:** 🟡 EM ANDAMENTO / REAVALIAÇÃO TÉCNICA

**Data de início:** 15/09/2026

**Data de conclusão:** —

**Quantidade de commits:** 3 no ciclo original da F4

### Objetivo

Integrar o núcleo validado com o SharePoint Online através de mecanismos
de leitura suportados.

### Implementado

- configuração por ambiente para tenant, aplicativo, segredo, site e drive,
  validada sem expor o valor do segredo;
- autenticação OAuth 2.0 `client_credentials` contra a plataforma de identidade
  Microsoft, solicitando o escopo `.default` do Microsoft Graph;
- fonte SharePoint com chamadas Graph exclusivamente `GET` para listar arquivos
  `.xlsx`, paginar resultados, enumerar versões e baixar conteúdo histórico;
- preservação de site, drive, DriveItem ID, nome, caminho, identificador da
  versão, data/hora, autor e tamanho disponíveis;
- comentário mantido nulo quando o contrato `DriveItemVersion` usado não fornece
  esse dado, evitando atribuição indevida de metadados;
- arquivos históricos em diretório temporário descartável;
- integração automatizada da fonte com o `AuditService`, sem reescrever reader,
  comparator ou persistência;
- falhas na listagem de versões agora registradas como falhas de execução pelo
  serviço de auditoria.

### Testes executados

Comando:

`pytest -q`

Resultado:

`28 passed in 1.04s`

Comandos adicionais:

- `python -m compileall -q app main.py tests` — concluído com código 0;
- `git diff --check` — concluído sem erros.

Tentativa de instalar as dependências:

`python -m pip install -r requirements.txt`

Resultado:

falhou porque o índice configurado no ambiente retornou `403 Forbidden` ao
consultar `msal`. A dependência foi eliminada: o fluxo OAuth suportado foi
implementado com a biblioteca padrão e o `requirements.txt` permaneceu enxuto.

### Critérios de aceite

[x] configuração e autenticação seguras implementadas;

[x] cliente limitado a operações de leitura no Microsoft Graph;

[x] DriveItem ID, versões, metadados disponíveis e conteúdo histórico cobertos
por testes automatizados;

[x] `AuditService` funciona com `SharePointSource` em teste isolado;

[x] motor Excel não foi reescrito;

[ ] autenticação e leitura comprovadas em SharePoint controlado;

[ ] versões reais, especialmente secundárias, enumeradas e recuperadas;

[ ] teste controlado concluído e ausência de escrita comprovada no ambiente.

### Commits

`df7943f` — hash citado à época, mas inexistente no Git real; a implementação Graph/bloqueio alcançável está em `1a7421a`.

O segundo commit registra este bloqueio e o estado factual da fase; seu
identificador é informado no relatório da execução.

### BLOQUEIO

**Problema:** os critérios obrigatórios de prova real não podem ser executados
sem tenant, aplicativo autorizado, segredo, site, drive e uma planilha
controlada acessível.

**Causa:** o ambiente da sessão não contém credenciais SharePoint nem o cenário
corporativo controlado.

**Impacto:** não é possível afirmar que o ambiente enumera ou permite baixar as
versões secundárias necessárias, nem concluir a F4.

**Alternativas:**

1. fornecer ao ambiente as cinco variáveis documentadas em `.env.example`, com
   permissões mínimas de leitura, e indicar uma planilha controlada;
2. executar o teste controlado externamente e fornecer evidências técnicas dos
   endpoints, metadados e versões recuperadas.

**Recomendação:** disponibilizar credenciais de aplicação com menor privilégio e
uma planilha controlada; então validar uma única planilha antes de qualquer uso
operacional.

### Pendências

- comprovar acesso ao site e à biblioteca no ambiente real;
- confirmar a ordem efetiva retornada pela coleção de versões;
- confirmar enumeração e download das versões secundárias;
- registrar metadados reais e executar a auditoria controlada.

### Próximo passo

Retomar somente a validação controlada da F4 após remoção do bloqueio. Não
iniciar a F5.

### Evolução do bloqueio

Após o registro inicial do bloqueio, foi confirmado pelo responsável do
projeto que o ambiente corporativo não disponibilizará acesso ao Microsoft
Entra Admin Center nem realizará o registro de aplicação originalmente
necessário para a estratégia `client_credentials`.

Consequentemente, a alternativa originalmente recomendada de fornecer
tenant, client ID, client secret, site ID e drive ID deixou de ser uma
dependência operacional viável para a V1 no ambiente atual.

A implementação Graph existente não deverá ser removida, pois permanece
arquiteturalmente válida para ambientes onde essa integração seja
autorizada.

Entretanto, Microsoft Graph deixou de ser considerado o único mecanismo
possível de aquisição SharePoint.

### Investigação `_vti_history`

Foi realizada investigação controlada utilizando a planilha real:

CQL028.xlsx

Foi confirmado que o SharePoint disponibiliza acesso à versão histórica
0.97, através da sessão autenticada do usuário, por endereço contendo:

`_vti_history/97/.../CQL028.xlsx`

Ao abrir esse recurso através do ambiente autenticado do usuário, foi
apresentada exatamente a versão histórica esperada.

A mesma versão também pôde ser aberta no Microsoft Excel como versão
anterior somente leitura.

Não foi executada restauração nem qualquer operação de escrita.

### Teste programático controlado

Foi executado um teste Python isolado, fora da aplicação, utilizando uma
requisição HTTP GET sem credenciais, cookies ou tokens para o mesmo
recurso histórico.

Resultado observado:

Status HTTP: 403
Content-Type: text/plain; charset=utf-8
Redirecionamento: nenhum
Corpo da resposta: 13 bytes

Conclusão:

a existência de uma URL `_vti_history` acessível pela sessão autenticada
do usuário não implica que o mesmo recurso esteja disponível
programaticamente sem autenticação apropriada.

Não foram extraídos cookies, tokens ou credenciais da sessão existente.

### Decisão decorrente

A Fase 4 foi revisada de:

"Microsoft Graph / SharePoint"

para:

"Aquisição de Versões SharePoint".

O objetivo permanece adquirir automaticamente as versões históricas
necessárias utilizando mecanismo suportado, autorizado e exclusivamente
de leitura.

A importação manual permanece somente como contingência e não foi
adotada como solução oficial da V1.

A Fase 5 permanece bloqueada até conclusão da F4 revisada ou nova decisão
expressa do responsável pelo projeto.
---

# 13. HISTÓRICO DA FASE 5

## F5 — Interface e Relatório

**Status:** 🟢 CONCLUÍDA

**Data de início:** 16/09/2026

**Data de conclusão:** 16/09/2026

**Quantidade de commits:** 4

### Objetivo

Disponibilizar interface simples para operação e geração do relatório
Excel consolidado.

### Implementado

Interface Tkinter/ttk e relatório consolidado implementados conforme o registro
factual no início deste documento.

### Testes executados

`pytest -q` — 35 testes aprovados. `ruff check` nos arquivos da entrega — aprovado.

### Commits

`0f95c32` implementou interface e relatório, `fdc12ad` corrigiu a validação real e
`1b71aec` adicionou a migração de bancos SQLite legados. Os três commits constituem o
limite ordinário da F5.

### Problemas encontrados

A migração aditiva do SQLite foi necessária e posteriormente aceita no banco real. As
mensagens visuais de download do Edge não representam, por si sós, falha do arquivo já
validado no filesystem.

### Pendências

Nenhuma pendência obrigatória da fase.

### Próximo passo

F6 iniciada por autorização expressa; executar uma tarefa por vez.

---

# 14. HISTÓRICO DA FASE 6

## F6 — Robustez e Preparação para Produção

**Status:** 🟢 CONCLUÍDA

**Data de início:** 16/09/2026

**Data de conclusão:** 16/09/2026

**Quantidade de commits:** limite ordinário excedido pelas exceções impostas pelo ambiente e pela autorização final do responsável

### Objetivo

Validar robustez, desempenho, integridade, logs, empacotamento e
preparação da V1 para homologação.

### Implementado

As sete tarefas de Testes Finais: testes unitários isolados, testes de
integração automatizados entre os componentes, reexecução idempotente após reabertura
do banco, auditoria incremental a partir do checkpoint, falha/retomada transacional e
geração determinística de relatório exclusivamente a partir do SQLite, além do cenário
controlado de três planilhas isoladas no mesmo banco canônico.
Também foram concluídos performance com carga sintética de 3.001 versões, 100.000
alterações e 2.000 planilhas, logs, temporários, SHA-256, estratégia mínima de backup e
empacotamento Windows one-folder. A validação final e os documentos oficiais foram
reconciliados.

### Testes executados

Testes unitários de configuração, persistência, reader, comparator, fonte SharePoint
com fakes, relatório e interface com fakes — aprovados. Seleção de três testes de
integração automatizados — `3 passed in 0.47s`. Reexecução idempotente e invariante
read-only do provider — `2 passed in 0.35s`. Auditoria incremental — `1 passed in
0.39s`; falha/retomada — `1 passed in 0.49s`; relatório — `3 passed in 0.36s`;
múltiplas planilhas — `1 passed in 0.56s`; suíte completa — `43 passed in 1.83s`.

### Commits

`e34c1f8` — inicia a F6 com testes unitários.

`2539e9b` — executa testes de integração da F6.

O terceiro commit registra a reexecução e tem seu identificador informado no relatório
da sessão, pois um commit não pode registrar o próprio hash. Com ele, o limite ordinário
de três commits da F6 foi atingido. Um quarto commit registra falha/retomada por
determinação do ambiente de execução; seu identificador é informado no relatório da
sessão.

### Problemas encontrados

Nenhum.

### Pendências

Nenhuma tarefa oficial da V1.

### Próximo passo

Utilização e homologação real.

---

# 15. REGISTRO DE DECISÕES TÉCNICAS

Esta seção registra somente decisões relevantes tomadas durante o
desenvolvimento.

---

## DEC-001 — Aplicação sem Node.js

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

A aplicação será baseada em Python e não possuirá dependência
obrigatória de Node.js, npm ou frameworks frontend baseados nesse
ecossistema.

### Motivo

Manter execução simples e reduzir dependências desnecessárias.

---

## DEC-002 — SQLite como banco inicial

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

A V1 utilizará SQLite como banco inicial.

### Motivo

Simplicidade operacional e facilidade de implantação.

### Observação

A arquitetura deverá permitir migração futura para SQL Server.

---

## DEC-003 — Banco como fonte oficial

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

O banco de auditoria será a fonte oficial da trilha consolidada.

O relatório Excel será uma representação gerada a partir do banco.

### Consequência

Alterações ou exclusão do relatório não deverão destruir a trilha
armazenada.

---

## DEC-004 — Auditoria incremental

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

Cada planilha possuirá checkpoint individual.

Exemplo:

Primeira execução:

0.84 → ... → 0.99

Checkpoint:

0.99

Execução posterior:

0.99 → ... → 1.20

Novo checkpoint:

1.20

O histórico consolidado não deverá ser reprocessado durante execução
incremental normal.

---

## DEC-005 — Identidade independente do nome

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

O nome da planilha não será utilizado como identidade técnica.

A aplicação utilizará identificadores fornecidos pelo
SharePoint/Microsoft Graph, preservando DriveItem ID e demais
identificadores necessários.

---

## DEC-006 — SharePoint somente leitura

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

A aplicação não poderá realizar operações de escrita no SharePoint.

Isso inclui:

- alteração;
- exclusão;
- criação de versões;
- check-in;
- check-out;
- alteração de metadados;
- alteração de comentários.

---

## DEC-007 — Limite de commits

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

Cada fase possuirá no máximo 3 commits.

Preferência:

1 ou 2 commits.

### Motivo

Evitar desenvolvimento excessivamente fragmentado e manter o projeto
curto e controlável.

## DEC-008 — Desacoplamento da aquisição SharePoint do Microsoft Graph

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

Microsoft Graph deixa de constituir mecanismo obrigatório e exclusivo
para aquisição das versões SharePoint na V1.

A arquitetura continuará permitindo Graph quando autorizado, mas a F4
passará a investigar mecanismo alternativo suportado e autorizado no
ambiente corporativo.

### Motivo

A integração Graph implementada não pôde ser validada no ambiente real
porque a organização não disponibiliza ao projeto o registro de aplicação
e as autorizações necessárias.

### Consequência

O motor Excel, AuditService, persistência, checkpoint e trilha consolidada
permanecem inalterados.

A aquisição deverá continuar isolada atrás do contrato de fonte.

## DEC-009 — Importação manual somente como contingência

**Data:** 15/09/2026

**Status:** APROVADA

### Decisão

O download e a importação manual de versões históricas não serão adotados
neste momento como mecanismo oficial da V1.

### Motivo

O cenário previsto inclui centenas de planilhas e potencialmente centenas
ou milhares de versões, tornando a aquisição manual inadequada como fluxo
operacional principal.

### Consequência

A investigação de aquisição automatizada deverá ser concluída antes de
decisão sobre contingência manual.

---

# 16. REGISTRO DE BLOQUEIOS

Nenhum bloqueio registrado até o momento.

Quando necessário utilizar:

## BLOQ-001 — Autenticação programática SharePoint no ambiente corporativo

**Data:** 15/09/2026

**Fase:** F4 — Aquisição de Versões SharePoint

**Status:** EM REAVALIAÇÃO

### Problema

A integração Microsoft Graph implementada não pode ser validada no
ambiente real utilizando o modelo de autenticação originalmente previsto.

### Causa

O projeto não dispõe de App Registration/autorização corporativa
necessária e essa disponibilização não está prevista no ambiente atual.

### Impacto

A aplicação ainda não consegue adquirir automaticamente as versões
históricas reais necessárias para concluir a F4.

### Evidências adicionais

O acesso humano autenticado à versão histórica 0.97 da CQL028.xlsx através
de `_vti_history` foi confirmado.

Uma requisição Python HTTP não autenticada ao mesmo recurso retornou
HTTP 403.

### Alternativas

1. identificar outro mecanismo Microsoft/SharePoint suportado e autorizado;
2. manter Microsoft Graph disponível para ambientes onde seja autorizado;
3. avaliar importação assistida somente como contingência caso nenhuma
   alternativa automatizada seja viável.

### Recomendação

Prosseguir com investigação técnica controlada das alternativas
suportadas, sem contornar mecanismos corporativos de autenticação.

### Decisão

F4 revisada e mantida em andamento/reavaliação.

F5 permanece não autorizada.

---

# 17. REGISTRO DE LIMITAÇÕES CONFIRMADAS

Nenhuma limitação técnica da implementação foi confirmada até o momento.

As limitações descritas nos documentos anteriores que ainda dependem
de validação técnica não deverão ser registradas aqui como fatos
confirmados.

Quando uma limitação for comprovada:

## LIM-001 — URL histórica não é acesso programático anônimo

**Data:** 15/09/2026

**Fase:** F4

### Comportamento esperado

Verificar se uma URL histórica `_vti_history` acessível pelo usuário
autenticado também poderia ser recuperada diretamente por requisição
Python sem autenticação adicional.

### Comportamento observado

A URL abriu corretamente a versão histórica 0.97 no ambiente autenticado.

A requisição Python HTTP sem autenticação retornou HTTP 403.

### Impacto

A URL `_vti_history` isoladamente não resolve a aquisição automatizada.

### Tratamento adotado

Investigar mecanismo suportado de autenticação/aquisição.

Não reutilizar cookies ou tokens de sessões existentes como contorno.

---

# 18. MODELO DE ATUALIZAÇÃO DE FASE

Ao concluir uma fase, utilizar aproximadamente:

## FASE X — NOME

**Status:** 🟢 CONCLUÍDA

**Início:** DD/MM/AAAA

**Conclusão:** DD/MM/AAAA

**Commits:** X

### Implementado

- [...]
- [...]
- [...]

### Arquivos principais

- [...]
- [...]
- [...]

### Testes

Executados:

`pytest ...`

Resultado:

XX passed

### Critérios de aceite

[x] requisito 1

[x] requisito 2

[x] requisito 3

### Commits

`abcdef1` — descrição

`abcdef2` — descrição

### Decisões

[...]

### Problemas encontrados

[...]

### Pendências

[...]

### Próximo passo

Fase X+1 aguardando autorização.

---

# 19. MODELO PARA FASE BLOQUEADA

Quando uma fase não puder ser concluída:

**Status:** 🔴 BLOQUEADA

Registrar obrigatoriamente:

- último ponto concluído;
- teste que apresentou problema;
- erro observado;
- impacto;
- alternativas;
- recomendação;
- estado do repositório.

Não marcar como concluída.

Não avançar para a próxima fase.

---

# 20. REGRA PARA TESTES

Não registrar:

"Testes OK"

sem informar quais testes foram executados.

Preferir:

Comando:

`pytest`

Resultado:

`24 passed`

ou equivalente.

Testes manuais relevantes também poderão ser registrados.

---

# 21. REGRA PARA COMMITS

Registrar o identificador real do commit.

Exemplo:

`a12bc34` — Implementa motor de comparação Excel

Não inventar hashes de commits.

Se ainda não houver commit:

registrar:

"Commit ainda não realizado."

---

# 22. REGRA PARA ALTERAÇÕES DOCUMENTAIS

Pequenos ajustes documentais realizados como consequência direta da
implementação poderão ser incluídos no commit correspondente.

Mudanças relevantes de:

- requisito;
- arquitetura;
- escopo;
- segurança;

deverão ser explicitamente registradas.

---

# 23. REGRA PARA NOVAS FUNCIONALIDADES

Sugestões surgidas durante o desenvolvimento não deverão ser
automaticamente implementadas.

Registrar como:

MELHORIA FUTURA

quando não forem necessárias para os critérios de aceite da V1.

Isso evita crescimento descontrolado do projeto.

---

# 24. MELHORIAS FUTURAS

Nenhuma melhoria adicional aprovada neste momento.

Possibilidades já identificadas, mas fora do escopo automático da V1:

- SQL Server;
- auditoria automática agendada;
- processamento em lote;
- Power BI;
- dashboards;
- notificações;
- múltiplos ambientes SharePoint.

A presença nesta seção não significa autorização para implementação.

---

# 25. ESTADO ATUAL OFICIAL

# 25. ESTADO ATUAL OFICIAL

**Data:** 16/09/2026

**Projeto:** Auditor de Planilhas Excel — SharePoint Online

**Versão planejada:** V1

**Fase atual:** F6 — Robustez e Preparação para Produção

**Status:** 🟢 CONCLUÍDA PARA UTILIZAÇÃO E TESTES REAIS

**Implementação:** núcleo local, aquisição SharePoint Edge/REST, interface, migração
SQLite e relatório validados no ambiente corporativo real; testes unitários, testes
de integração automatizados, reexecução idempotente, auditoria incremental,
falha/retomada, geração de relatório, múltiplas planilhas controladas e medição de
performance, logs, temporários, integridade por SHA-256, estratégia mínima de backup,
empacotamento e validação final da F6 concluídos. A estratégia de backup foi somente
documentada, como exigido pela seção 43, sem implementação operacional.

**Fases concluídas:** 6/6

**Commits do ciclo original da Fase 4:** 3

**Bloqueios ativos:** 0

**Limitações confirmadas:** 1

**Próxima ação:**

Utilização e homologação real. Não existe tarefa oficial pendente na V1.

---

# 26. INSTRUÇÃO AO CODEX

Antes de iniciar qualquer implementação:

1. consultar a documentação obrigatória na ordem definida por
   `06_GOVERNANCA.md`;
2. verificar o estado real do repositório;
3. consultar este histórico;
4. identificar a fase e a tarefa expressamente autorizadas;
5. respeitar bloqueios e decisões registrados.
6. identificar a fase autorizada.

Após executar a fase:

1. atualizar este documento;
2. registrar apenas fatos reais;
3. registrar testes efetivamente executados;
4. registrar hashes reais dos commits;
5. informar bloqueios e limitações;
6. parar antes de iniciar a fase seguinte.

---

FIM DO DOCUMENTO
