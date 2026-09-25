# Processamento independente de pares

## Referência histórica

O commit `b21993a` (pai de `18f7830`, que introduziu o cache de snapshots) foi
usado como referência para restaurar a unidade de trabalho completa. A
restauração foi seletiva: o pool de WebDrivers, catálogo de versões, perfil de
runtime, staging e promoção ordenada posteriores foram preservados.

## Caminho principal

1. O scheduler mantém uma janela de `slots + 2 * slots` pares.
2. Até `driver_count` atores de download preparam pares em paralelo.
3. Um registry sincronizado por `technical_version_id` e uma `Condition`
   garantem que apenas um ator baixe cada arquivo. Os demais reutilizam o mesmo
   caminho local depois de `FILE_READY`.
4. Quando os dois caminhos estão prontos, o scheduler envia uma
   `ComparisonTask` autocontida ao primeiro slot livre.
5. O próprio worker abre o XLSX anterior e o atual e executa o comparator. Não
   há snapshot compartilhado, serialização de workbook ou espera por eventos de
   snapshot.
6. O resultado vai para o SQLite de staging. O `OrderedCommitCoordinator`
   promove apenas o maior prefixo contíguo para o SQLite oficial.

Arquivos intermediários são deliberadamente lidos duas vezes por pares
adjacentes, mas são baixados uma única vez durante a execução. Eles só são
liberados depois que a execução encerra, portanto nenhum par ativo ou futuro da
janela perde o arquivo.

## Telemetria

O fluxo emite `PAIR_SUBMITTED`, `PAIR_READ_STARTED`, `PAIR_READ_FINISHED`,
`PAIR_COMPARE_FINISHED`, `PAIR_STAGED` e `PAIR_PROMOTED`, além de
`FILE_DOWNLOAD_REQUESTED`, `FILE_READY`, `FILE_REUSED` e `FILE_RELEASED`.
`WORKER_POOL_SUMMARY` registra leituras esperadas/reais, utilização, workers
ativos (média/pico), esperas de download/scheduler/commit e
`snapshot_mode=disabled`.

## Benchmark real

As quatro execuções solicitadas para `CQLPA123` (5/7 slots, 1/2 WebDrivers)
dependem de autenticação SharePoint, Edge e do histórico corporativo. Elas não
podem ser reproduzidas somente com os fixtures locais do repositório. Devem ser
executadas no ambiente autenticado com o mesmo intervalo de 30 ou 50 pares; o
baseline histórico informado é de aproximadamente 15 versões/minuto com sete
slots.
