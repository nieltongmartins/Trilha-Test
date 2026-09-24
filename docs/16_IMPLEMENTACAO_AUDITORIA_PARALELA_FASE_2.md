# Fase 2 — protótipo de auditoria paralela (1/2 slots)

Data: 2026-09-24. Escopo exclusivo: `Trilha-Test`.

## Arquitetura implementada

`ParallelAuditService` conserva as três unidades da Fase 1: `VersionTask`
(reserva, aquisição e identidade), `ComparisonTask` (paths e IDs serializáveis) e
`PromotionTask` (descrição da promoção). O scheduler central é o único dono dos
estados, tokens de reserva e fila. Só aceita **1 ou 2** slots; qualquer outro
valor falha imediatamente. A janela é `2 * slots` (2 no controle serial, 4 no
ensaio paralelo) e um slot liberado rouba a próxima tarefa elegível sem esperar
promoção.

A thread coordenadora é o único ator que chama `get_version`,
`verify_download_digest` e `release_version`. Portanto nenhum comando concorre
no mesmo WebDriver. SHA-256 continua sendo calculado sobre o arquivo adquirido;
a validação ZIP e o rename atômico permanecem responsabilidade da fonte
SharePoint existente. Um ID técnico é adquirido uma vez por execução.

O backend padrão usa `ProcessPoolExecutor`: cada worker recebe apenas paths,
metadados imutáveis, sequence e token; cria duas leituras completas independentes
com `read_workbook` e compara os snapshots. Não compartilha reader, snapshot,
banco ou WebDriver. Há também backends `thread` e `sync` para comparação
controlada.

## Staging e promoção

Cada execução cria um SQLite temporário `parallel-<execution>.sqlite3`, separado
do banco oficial. O resultado e métricas são gravados antes de `complete=1` na
mesma transação. `OrderedCommitCoordinator` procura exclusivamente
`next_sequence`; se ausente, não promove posteriores. Para cada item do prefixo
contíguo chama a transação oficial já existente, que insere versão, alterações e
checkpoint atomicamente. Workers jamais recebem a conexão oficial.

Arquivos de staging órfãos são descartados no início. A retomada reconstrói o
plano a partir do checkpoint oficial. Pause impede reservas novas e espera os
workers atingirem staging; todos os slots então publicam `PAUSADO`. Stop impede
reservas, fecha o executor em boundary seguro, preserva o prefixo promovido e
remove o staging descartável. Falha de worker tem um retry; esgotado o limite,
marca `ERRO`, registra slot/ID/label/etapa/tentativas e impede a promoção pela
lacuna.

## Interface e telemetria

A tela possui blocos fixos SLOT 1 e SLOT 2 com label, ID técnico, estado, etapa,
percentual e duração. O progresso global informa committed/total, checkpoint,
ativos, staged, taxa móvel de commits nos últimos 60 segundos e ETA calculada
pela taxa agregada. Eventos continuam atravessando filas e só a thread Tk altera
widgets.

São emitidos registros por slot (task/technical ID/sequence/stage, download,
parse, compare, staging e duração), promoção (espera no staging e commit) e
amostras globais (ativos, staged, checkpoint sequence, janela, throughput e RSS
do Python/workers; RSS do Edge é zero quando não há medidor externo).

## Provas automatizadas

`tests/test_parallel_audit.py` prova equivalência oficial 1×2 slots, unicidade,
conclusão física fora de ordem com bloqueio do checkpoint e descarte do staging
após crash simulado. A suíte completa preserva os testes seriais existentes.

## Benchmark reproduzível

O repositório não contém nem permite obter nesta execução as 20–50 versões reais
da CQLPA123 autenticada. Para não fabricar uma medição de produção, foi criado e
executado um corpus **CQLPA123 sintético**, 24 versões × 2.500 linhas × 4 células,
com a mesma leitura XLSX e persistência oficial. Resultado detalhado e
machine-readable: `docs/16_BENCHMARK_FASE_2.json`.

| modo | tempo | comparações/min | RSS Python | CPU Python | speedup |
|---|---:|---:|---:|---:|---:|
| síncrono, 1 slot | 5,306 s | 260,092 | 37,38 MiB | 5,239 s | 1,000× |
| threads, 2 slots | 5,441 s | 253,609 | 40,12 MiB | 5,519 s | 0,975× |
| processos, 2 slots | 3,735 s | 369,449 | 40,12 MiB | 2,010 s* | **1,421×** |

\* CPU da linha de processos é a do coordenador; CPU dos filhos é exposta pela
telemetria runtime, mas não somada pelo relatório portátil. Todos produziram
57.500 alterações, 23 pares, zero erro/retry e bancos equivalentes. Download
local médio foi 0,000099/0,000077/0,000149 s; parsing médio
0,152480/0,193792/0,180721 s; comparação média
0,027311/0,028614/0,027414 s (serial/threads/processos).

## Decisão e riscos

O protótipo prova ganho no corpus controlado e confirma processos como candidato;
threads não trouxeram ganho. Ainda **não se recomenda avançar à Fase 3** antes de
rodar a mesma ferramenta com 20–50 binários reais da CQLPA123, medir RSS total
dos filhos e Edge em Windows, e exercitar encerramento forçado do processo pai.
Outros riscos: duas leituras completas por par aumentam I/O/RAM, pickle de uma
lista grande de alterações pode pressionar IPC, e pause aguarda o término do
parse corrente. Não foram implementados 3–5 slots, cache mutável, sharedStrings,
row/worksheet reuse, ArrayFormula, relatório ou mudança de schema oficial.
