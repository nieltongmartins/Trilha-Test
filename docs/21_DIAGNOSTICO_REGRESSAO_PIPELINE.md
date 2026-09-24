# Diagnóstico da regressão do pipeline paralelo

## Referência e histórico comparado

O último baseline mensurado antes das alterações de interface/tempo foi o commit
`d7b8e7f` (fase 4). Ele contém o benchmark controlado de 4 slots e preserva o
modelo de um proprietário do downloader, staging ordenado e checkpoint pelo
maior prefixo contíguo. Foram comparados os commits posteriores `720fbe7`,
`7fee31f`, `ab8b367`, `f5fff16`, `d5ddc3c`, `9f1b96c` e `ecbbfcb`.

O benchmark histórico de CQLPA123 sintética em 4 slots foi 2,322 s para 29
comparações (749,360 comparações/minuto, eficiência 0,417). Esse número é uma
referência controlada; o repositório não contém as 25 versões reais, as
credenciais do SharePoint nem o Edge necessários para reproduzir o ensaio real.

## Matriz de diferenças anterior à correção

| Arquivo | Função/área | Comportamento no baseline `d7b8e7f` | Comportamento antes da correção | Possível impacto funcional |
|---|---|---|---|---|
| `app/parallel_audit.py` | `ParallelAuditService.audit` | O coordenador adquiria arquivos e, entre aquisições, coletava futures e promovia o prefixo. | Prefetch e aquisição continuaram síncronos na mesma thread que chama `_collect_done` e `_promote`; um `get_version` lento impede ambos. | Um worker concluído permanece invisível até o WebDriver retornar, inflando `pipeline_latency` e atrasando staging/checkpoint. |
| `app/parallel_audit.py` | reserva de slot | A reserva era feita antes de baixar o par, e o cronômetro da task começava nesse ponto. | As novas métricas tornaram explícito que download/espera eram atribuídos ao slot, embora nenhum processo estivesse executando a comparação. | Estado visual `WORKER_RUNNING` e tempo do slot incluem espera de aquisição; utilização aparente não representa trabalho real. |
| `app/parallel_audit.py` | `_collect_done` | Fazia polling a cada 50 ms quando nenhum future terminara. | Continua sendo o único caminho para staging e usa `future.result()` apenas depois de `done()`, mas só é chamado depois do trecho WebDriver bloqueante. | O `result()` não é a causa direta; a ausência de um consumidor independente durante o fetch é a causa do atraso observado. |
| `app/parallel_audit.py` | `_metrics`/`_slot` | Telemetria temporal e de memória era calculada no coordenador, com logs limitados por transição/3 s. | Estatísticas adicionais, RSS/CPU, callbacks e estimativas continuam no loop funcional. | Custo menor que o fetch, porém ainda acopla observação ao caminho crítico e aumenta jitter/logging. |
| `app/interface.py` | `_poll_work_result` e `_tick_slot_progress` | Tk fazia polling de filas, atualizava variáveis e interpolava progresso. | O ticker lê eventos enfileirados e o modelo compartilhado; não chama WebDriver nem `Future.result()`. | Não explica sozinho dezenas de segundos, mas o uso de objetos temporais compartilhados impede uma fronteira observacional clara. |
| `app/execution_timing.py` | `SharedExecutionTimingModel` | Modelo estatístico compartilhado para ETA e estágios. | Foi ampliado com médias recentes/globais e relógio ativo. | Seguro enquanto somente observado; não deve determinar sleep, reserva, promoção ou prefetch. |
| `app/sources/sharepoint.py` | `get_version`/prefetch | Um único chamador serializa o WebDriver. | `get_version` pode aguardar um fetch iniciado no Edge; `queue_age` e `active_fetch_elapsed` são expostos, mas a espera acontece na thread coordenadora. | Enquanto o owner espera o Edge, staging e promoção ficam parados. |
| `app/sources/sharepoint.py` | decisão de timeout | Timeout técnico nasceu no JavaScript a partir do tempo ativo do fetch. | Métricas distinguem idade desde enqueue e duração ativa, mas faltava log explícito da decisão. | Diagnóstico ambíguo pode atribuir timeout à idade da fila. |
| `app/runtime_profile.py` / `app/parallel_audit.py` | perfil/prefetch inicial | O baseline usava alvo fixo `2 * slots`. | O perfil persistente existe, mas o carregamento depende do caminho da UI; instâncias diretas podem iniciar com `avg_file_bytes=0`, e a calibração por metadados pode sobrescrever a recomendação carregada. | CQLPA123 pode voltar ao alvo 8 apesar do histórico válido. |
| `app/audit_service.py` | persistência oficial | Comparação/checkpoint transacionais. | Sem alteração relevante nos commits investigados. | Deve permanecer inalterado. |

## Causa raiz

A regressão arquitetural é o acoplamento do **owner bloqueante do WebDriver** ao
**único consumidor de conclusões**. O loop agenda prefetch, reserva um slot e
chama `_acquire`; somente depois que todas as chamadas Selenium/SHA daquele par
retornam é que executa `_collect_done` e `_promote`. Assim, uma comparação de 13
s que termina durante um fetch de aproximadamente 48 s só entra em staging no
retorno do fetch, produzindo a diferença observada entre `worker_task_duration`
e `pipeline_latency`.

A correção deve conservar um único owner do WebDriver, mas fazê-lo publicar
arquivos prontos em uma fila independente. O scheduler só atribui um slot após
o par estar pronto, e o consumidor de resultados/staging/ordered commit continua
rodando durante downloads. UI, ETA e perfil permanecem observadores ou
configuração inicial, nunca sinais de controle do pipeline.

## Validação após a correção

O pipeline corrigido executa aquisição em um executor de um único thread
(`webdriver-owner`) e mantém staging/promoção no coordenador. O teste de
regressão bloqueia deliberadamente o download da terceira versão por 300 ms e
confirma que o checkpoint da primeira comparação é promovido durante esse
bloqueio. Outro teste mede quatro comparações simultâneas e exige latência
`worker_finished -> staged` inferior a 500 ms.

No benchmark sintético reproduzível de 25 versões CQLPA123/1.000 linhas, 4
processos concluíram 24 comparações em 2,191 s (657,356 comparações/minuto), com
utilização observada de 71,56% a 85,77% entre os slots. O baseline histórico
equivalente documentado em `d7b8e7f` usou 30 versões e marcou 2,322 s e 749,360
comparações/minuto; portanto tempos/throughput não são diretamente comparáveis
como A/B por terem quantidade de versões e execução distintas. A integridade
oficial do ensaio corrigido foi idêntica ao serial.

O ambiente não contém CQLPA123 real, credenciais nem Edge, então não é possível
publicar números reais de `READ_XLSX ≈ 11 s` antes/depois. A suíte automatizada
cobre a propriedade arquitetural que elimina os aproximadamente 48 s espúrios:
o commit progride enquanto o fetch está ativo e a telemetria registra, por task,
`reserved_at`, `download_ready_at`, `worker_started_at`, `worker_finished_at`,
`staged_at` e `promoted_at`, além das diferenças entre cada etapa.

As métricas de processo são limitadas a cinco snapshots por segundo e o log de
memória a uma emissão por segundo. A interface consome exclusivamente eventos
imutáveis `SlotProgress`/`ParallelMetrics`; o ticker não consulta mais o modelo
temporal compartilhado. Medições alternadas com logging normal/reduzido no
container apresentaram ruído de processo acima de 3% (medianas 592,112 e
629,870 comparações/minuto), sem Edge e sem uma UI Tk real. Por isso esses
números são diagnóstico ambiental, não evidência de overhead causal; o A/B real
obrigatório deve ser repetido na máquina alvo com as mesmas 25 versões reais.
