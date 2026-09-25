# Pool dinâmico de WebDrivers

## Arquitetura

Antes desta alteração, uma única thread `webdriver-owner` fazia toda aquisição.
Agora o scheduler aceita de um a quatro drivers e entrega tarefas por prioridade a
uma fila central. Cada `DriverWorker` possui exatamente uma thread e somente essa
thread valida, recupera e usa sua instância Selenium. Slots de comparação não são
associados a drivers. O registry por `technical_version_id` devolve o mesmo
`Future` para pedidos concorrentes, impedindo download duplicado; arquivos somente
entram no pipeline existente depois da validação XLSX, SHA e rename atômico já
existentes.

Com `driver_count=1`, não é criado pool adicional: permanecem o driver, prefetch,
autenticação e caminho serial anteriores. Com mais drivers, o prefetch JavaScript
do driver primário é desativado; os atores paralelos são os produtores da fila
READY. Comparação, staging, promoção ordenada e checkpoint não foram alterados.

## Autenticação e isolamento

A sessão atual é uma sessão interativa Edge, com autenticação federada mantida no
perfil do navegador e requisições `fetch(..., credentials='same-origin')`. A
aplicação deliberadamente não lê nem copia cookies, tokens, `localStorage` ou
`sessionStorage`: copiar apenas parte desse estado não é uma derivação segura de
sessão e perfis Edge não podem ser compartilhados simultaneamente.

Cada driver adicional usa um `user-data-dir` temporário exclusivo. Ele navega para
o site e depende do SSO integrado do ambiente. Antes de entrar no pool são
validados `session_id`, window handle, origem, `document.readyState` e um GET REST
autorizado (`web?$select=Id`). Se o tenant não autenticar automaticamente o novo
perfil, a auditoria falha de modo controlado, sem solicitar logins manuais
repetidos nem iniciar downloads. Perfis temporários são removidos no encerramento.

## Falhas, stop e telemetria

Uma falha recoloca a mesma tarefa na fila uma vez; outros atores continuam. O stop
impede novas distribuições e o fechamento acorda os atores, encerra os drivers
adicionais e mantém o checkpoint governado pelo prefixo contíguo. Os eventos
`WEBDRIVER_CREATED`, `WEBDRIVER_READY`, `WEBDRIVER_DOWNLOAD_STARTED`,
`WEBDRIVER_DOWNLOAD_FINISHED`, `WEBDRIVER_RECOVERY`, `WEBDRIVER_FAILED` e
`WEBDRIVER_POOL_SUMMARY` permitem calcular downloads/minuto e utilização.

## Benchmark controlado

O ambiente automatizado não contém credenciais, Edge nem acesso ao tenant da
CQLPA123. Portanto, resultados reais para 1/2/3/4 drivers, CPU/RAM, throttling e
equivalência dos 25 pares **não foram inventados** e permanecem pendentes de
execução no ambiente autenticado. Os testes automatizados cobrem deduplicação,
owner exclusivo, retry isolado, compatibilidade do modo padrão e checkpoint.
