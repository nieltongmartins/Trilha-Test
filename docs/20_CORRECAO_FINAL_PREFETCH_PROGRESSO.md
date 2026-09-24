# Correção final: prefetch, workers e progresso

## Política e limites

O alvo do buffer é calculado por `prefetch_target = 2 * slots_selected`. A
janela limitada do scheduler é `slots_selected + prefetch_target`, portanto
varia de 3 (um slot) a 24 (oito slots). O reabastecimento roda em toda iteração
do scheduler e conta exclusivamente aquisições `PREFETCH_PENDING` e
`PREFETCH_READY`; o mapa indexado por `technical_version_id` impede uma segunda
aquisição. O WebDriver continua acessível apenas pelo coordenador.

Sob pressão (90% da memória física ou menos de 512 MiB disponíveis), novas
aquisições são suspensas. Itens já preparados não são descartados. A telemetria
expõe bytes declarados no buffer, quantidade de blobs no Edge, arquivos `.part`,
RSS do coordenador, workers e Edge, além do RSS total.

## Pool real

Não existe limite fixo de quatro processos. No backend de produção,
`ProcessPoolExecutor(max_workers=slots_selected)` configura oito workers para
oito slots. O executor cria processos sob demanda; assim, uma execução curta ou
limitada pelo downloader serial pode observar menos PIDs. O evento
`WORKER_POOL_SUMMARY` separa `configured_workers` de `effective_worker_pids` e
publica tempos ocupado/ocioso de workers e slots. Para oito slots, o resultado
esperado em uma carga longa é **oito configurados e até oito PIDs efetivos**.

## Progresso visual

As fases visuais são DOWNLOAD, SHA, READ_XLSX, COMPARE e STAGING. Seus pesos são
as médias robustas atuais do `SharedExecutionTimingModel`, divididas pela soma
das médias dessas fases. Dentro da fase, a fração é 90% linear até sua duração
estimada e usa uma cauda exponencial assintótica depois dela. A interface recalcula
os slots a cada 50 ms no `MainThread`, aplica `max(progresso_anterior,
progresso_calculado)` e limita tarefas ainda não concluídas a 99%. Somente o
checkpoint confirmado publica 100%.

## Benchmark CQL028

Os dados e a sessão autenticada da CQL028 não fazem parte deste repositório.
Consequentemente, medições reais de hit/miss, starvation, CPU, RAM, tempo e
throughput para 1/4/8 slots devem ser coletadas no ambiente SharePoint pelo log
`PREFETCH_SUMMARY`, `TELEMETRIA_MEMORIA` e `WORKER_POOL_SUMMARY`; não se registram
números sintéticos como se fossem o benchmark real. A suíte automatizada valida
a equivalência do banco oficial entre um e oito slots, unicidade de versões,
ordenação do checkpoint, fórmulas de janela/target e continuidade monotônica do
modelo de progresso.
