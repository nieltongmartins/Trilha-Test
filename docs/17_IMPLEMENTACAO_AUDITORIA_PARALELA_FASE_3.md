# Auditoria paralela — Fase 3

## Arquitetura

`ParallelAuditService` recebe de 1 a 5 slots. A fila continua dinâmica: o
primeiro slot livre reserva a próxima sequência dentro da janela limitada a
`2 * slots`. A aquisição continua no coordenador (um único proprietário da
fonte/WebDriver); somente leitura XLSX e comparação são enviadas aos processos.
Uma tarefa libera seu slot assim que o resultado entra no SQLite temporário.

`OrderedCommitCoordinator` permanece o único escritor do banco oficial e o
único componente que avança o checkpoint. Resultados fora de ordem aguardam no
staging descartável. No restart, staging órfão é eliminado e todo trabalho após
o checkpoint oficial é refeito.

## Modelo temporal compartilhado

Todos os workers alimentam uma única instância thread-safe de
`SharedExecutionTimingModel`. São medidos separadamente `DOWNLOAD_FETCH`,
`DOWNLOAD_TRANSFER`, `SHA`, `READ_XLSX`, `COMPARE`, `STAGING`,
`WAIT_PROMOTION`, `COMMIT` e `TOTAL_TASK`. Etapas ainda sem amostra usam baseline
visual conservador, sem qualquer efeito funcional.

Cada etapa conserva uma janela móvel das últimas 20 observações. Para ETA, é
calculada a média depois de winsorizar valores nos limites
`mediana ± 3 * 1,4826 * MAD`. Quando MAD é zero, o teto é o maior entre três
vezes a mediana e mediana mais um segundo. A observação bruta, inclusive
outlier, continua registrada na telemetria; apenas sua influência visual é
limitada.

O peso de uma etapa é sua duração esperada dividida pela soma das durações das
etapas operacionais. Dentro da etapa, até o tempo esperado a fração é
`0,9 * elapsed / expected`; depois usa a cauda
`0,9 + 0,1 * (1 - exp(-(ratio - 1)))`. Portanto há movimento contínuo, sem
confirmar uma etapa ou alcançar 100% pelo relógio. Somente o evento funcional
de conclusão produz 100%.

O ETA do slot soma a fração restante da etapa atual às médias das etapas futuras;
etapas concluídas não são incluídas novamente. O ETA global usa a taxa das 20
promoções mais recentes após três commits. Antes disso, faz bootstrap por
`tarefas restantes * tempo médio da tarefa / slots ativos`; ETAs dos slots nunca
são somados.

Pause congela o relógio ativo e exclui o intervalo pausado. Stop congela
definitivamente o relógio e impede novas reservas. Nenhum relógio ou ETA decide
conclusão, promoção ou checkpoint.

## Interface e telemetria

O seletor **Processamentos simultâneos** oferece 1–5, fica bloqueado durante a
auditoria e volta após Stop/conclusão. A tela cria exatamente a quantidade de
barras escolhida. Cada evento mostra versão, identidade técnica, estado,
decorrido, média compartilhada (ou “Calculando...”), estimativa total e restante.

A telemetria de tarefa inclui `slot_id`, `worker_pid`, `worker_thread_id`,
`coordinator_pid`, `technical_version_id` e `sequence`. As métricas incluem
ocupação da janela, staged máximo e RSS do coordenador/workers.

## Validação e benchmark preliminar

Os testes automatizados cobrem janela, compartilhamento entre slots, outlier,
bootstrap, progresso assintótico, ETA individual/global, Pause/Stop, conclusão
fora de ordem, crash/restart e equivalência oficial entre 1, 2, 3, 4 e 5 slots.

O benchmark sintético controlado (`8` versões, `100` linhas) está em
`17_BENCHMARK_FASE_3.json`. Todos os modos foram funcionalmente equivalentes e
produziram 700 alterações. Os resultados são apenas instrumentação preliminar:
1/2/3/4/5 slots obtiveram respectivamente 1591/2311/1909/1449/1274 comparações
por minuto e speedup 1,00/1,45/1,20/0,91/0,80. RSS Python máximo foi cerca de
29,6 MiB e RSS de filhos reportado chegou a 23,8 MiB; Edge foi zero por usar
fonte local. O workload curto é dominado pelo custo de criação dos processos e
não permite escolher um vencedor.

## Riscos para a Fase 4

* validar em SharePoint real a separação entre fetch e transferência, hoje
  limitada pelo contrato síncrono de `get_version`;
* coletar amostra longa para estabilizar throughput e medir RSS individual por
  PID (o benchmark atual usa o máximo agregado de filhos);
* persistir opcionalmente baselines compatíveis entre execuções;
* decidir concorrência padrão somente após benchmark real de CPU, RAM, Edge e
  throttling do SharePoint.
