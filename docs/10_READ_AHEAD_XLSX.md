# Leitura antecipada XLSX

## Decisao

Foi adotado `ProcessPoolExecutor(max_workers=1)`, com buffer nomeado
`READ_AHEAD_BUFFER_SIZE = 1`. A leitura XML/XLSX e predominantemente CPU-bound;
um processo evita a disputa pelo GIL com comparacao e persistencia. Uma thread
teria menor custo de serializacao, mas nao garante paralelismo para os trechos
Python do parser.

Somente a proxima versao e materializada e submetida. O processo principal
continua sendo o unico responsavel por comparar, gravar no SQLite e confirmar o
checkpoint. O resultado transporta ID, rotulo, origem absoluta e token aleatorio;
os quatro campos sao validados antes que o snapshot seja usado.

## Memoria e observabilidade

Cada consumo registra `rss_principal`, `rss_worker`, bytes do snapshot
serializado, numero de celulas, tempo de leitura, espera e se o resultado ja
estava pronto. A fila tem tamanho fixo um. O caminho temporario e liberado assim
que o resultado e consumido; as referencias ao `Future` e a identidade tambem
sao removidas. Em falha, o executor e encerrado e o arquivo pendente e liberado.

## Ordem e falhas

Uma leitura futura pode terminar antes da comparacao atual, mas seu resultado so
e solicitado na iteracao correspondente. Falhas do worker sao propagadas nessa
iteracao; portanto a versao nao e pulada e seu checkpoint nao avanca. Reinicios
seguem o checkpoint transacional ja existente.

## Benchmark

O repositorio nao inclui acesso autenticado ao historico real do SharePoint nem
um conjunto de 50 versoes reais. Por isso nao ha alegacao de ganho real nesta
alteracao. Os logs `PERF read_ahead` fornecem as metricas necessarias para executar
o comparativo A/B no ambiente corporativo (media, mediana e p95 devem ser
calculados sobre pelo menos 50 versoes), junto de CPU, erros e retries da
telemetria operacional existente.

O custo adicional esperado e a serializacao interprocesso e a convivencia
temporaria dos snapshots estritamente limitados. Esse custo, especialmente em
Windows (`spawn`), deve ser confrontado com os tempos reais antes de ampliar o
buffer ou o numero de workers.
