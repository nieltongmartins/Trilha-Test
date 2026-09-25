# Parse único de XLSX e cache deslizante de snapshots

## Arquitetura

Antes, cada worker recebia um par de caminhos e executava duas chamadas ao
reader antes de comparar. Em uma cadeia de 25 comparações isso permitia 50
parses. Agora o coordenador registra as dependências por
`(workbook_identity, technical_version_id)`, agenda uma `VersionSnapshotTask`
por versão local e entrega à `ComparisonTask` apenas dois `VersionSnapshot`
prontos. O comparador e o `OrderedCommitCoordinator` não foram modificados.

`VersionSnapshot` e suas abas implementam somente a interface `Mapping`, não
expõem mutação e são serializáveis. Cada parse cria seu próprio reader por meio
de `read_workbook`; nenhum reader é compartilhado. O cache tem capacidade
`scheduler_window + 1`, isto é, acompanha a janela configurada em vez de crescer
com todo o histórico.

O registry distingue `NOT_REQUESTED`, `PARSING`, `READY`, `FAILED` e
`RELEASED`. Uma segunda solicitação reutiliza o mesmo `Future`. As sequências
consumidoras formam o refcount explícito; a retenção é conservadora até o
staging ser promovido e somente a remoção da última sequência libera o objeto.
Uma falha nunca publica resultado parcial e precisa ser resetada explicitamente
antes do retry.

## Telemetria

São emitidos `SNAPSHOT_REQUESTED`, `SNAPSHOT_PARSE_STARTED`,
`SNAPSHOT_PARSE_READY`, `SNAPSHOT_REUSED`, `SNAPSHOT_RELEASED` e
`SNAPSHOT_CACHE_SUMMARY`. O resumo inclui versões únicas, parses, reutilizações,
parses duplicados evitados, pico de residentes, pico estimado de bytes e
amplificação. O benchmark também registra bytes e duração de serialização.

A desserialização automática feita pelo `ProcessPoolExecutor` ocorre antes da
entrada na função worker, portanto sua duração isolada não é observável sem
alterar o protocolo de transporte. Ela é registrada como `null`, e não como um
zero enganoso.

## Benchmark CQLPA123 sintético

Medição reproduzível com `tools/benchmark_parallel_audit.py`, backend process,
5 slots, 1.000 rows por versão, incluindo startup e sem warm-up descartado. A
baseline foi executada em worktree destacada no commit anterior e o novo
pipeline no mesmo host. O conjunto é sintético, pois credenciais/dados reais do
SharePoint CQLPA123 não fazem parte do repositório.

| Métrica | 25 comparações (antes) | 25 (depois) | 50 (antes) | 50 (depois) |
|---|---:|---:|---:|---:|
| versões únicas | 26 | 26 | 51 | 51 |
| parse count | 50 | 26 | 100 | 51 |
| duplicados evitados | 0 | 24 | 0 | 49 |
| amplificação | 1,923 | 1,000 | 1,961 | 1,000 |
| READ_XLSX acumulado | 5,342 s | 3,657 s | 8,473 s | 5,433 s |
| elapsed total | 2,118 s | 2,325 s | 3,716 s | 3,380 s |
| comparações/min | 708,3 | 645,2 | 807,4 | 887,5 |
| RSS total de pico | 167,15 MiB | 184,66 MiB | 166,81 MiB | 183,53 MiB |
| serialização de snapshot | n/a | 0,186 s / 1.688.453 B | n/a | 0,327 s / 3.312.603 B |

Para 25 comparações, `speedup_total=0,911` e `speedup_read=1,461`; essa amostra
curta regrediu em elapsed. Para 50, já estável, `speedup_total=1,099` e
`speedup_read=1,560`, com throughput 9,9% maior. O RSS total subiu cerca de 10%
porque snapshots imutáveis atravessam o limite de processos. O cache permaneceu
limitado; isso é custo de transporte/residência, não crescimento com o histórico.

## Equivalência, testes e riscos

Os testes confrontam o SQLite oficial entre 1 e 8 slots, exercitam ordem de
staging/checkpoint e agora também validam corrida simultânea, quatro parses para
`A-B-C-D`, refcount, release, falha/reset, limite da janela, imutabilidade e
reuso com 1 slot. O benchmark produziu 25.000 e 50.000 mudanças, respectivamente,
iguais à baseline para os mesmos geradores.

O risco restante é o custo de pickle em workbooks maiores. A amostra longa teve
ganho total, mas a curta mostrou ruído/regressão e o RSS cresceu. Se dados reais
mostrarem regressão relevante, deve-se avaliar, em tarefa separada e antes de
implementar, afinidade de worker, armazenamento intermediário ou memória
compartilhada. Nenhuma dessas ampliações foi incluída aqui.
