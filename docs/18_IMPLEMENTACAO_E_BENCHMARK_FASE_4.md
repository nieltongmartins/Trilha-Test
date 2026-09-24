# Fase 4 — escala de 1–8, interface compacta e benchmark final

## Escopo implementado

`ParallelAuditService` aceita de 1 a 8 processos e preserva o downloader de
proprietário único, a janela `2 * slots`, o staging descartável e um único
`OrderedCommitCoordinator`. Não houve mudança nas regras de comparação, na
identidade técnica, no SQLite oficial nem no avanço pelo maior prefixo contíguo.

A interface materializa somente os slots escolhidos, em duas colunas com pesos
iguais. Cada cartão foi reduzido a versão, etapa, progresso, decorrido e ETA; o
ID técnico permanece na telemetria. O grid ocupa no máximo quatro linhas, o que
mantém o painel utilizável nas metas de 1366×768, 1600×900 e 1920×1080. A
validação desta entrega foi estrutural/headless: o ambiente não oferece display
gráfico para captura visual.

## Telemetria e proteção operacional

* `TIMING_MODEL` informa estágio, amostras, média, mediana e estimativa robusta,
  com throttle de dois segundos por estágio.
* `SLOT_PROGRESS` informa slot, sequência, estágio, decorrido, estimativa do
  estágio, percentual e ETA, com throttle de três segundos ou mudança de estado.
* `TELEMETRIA_MEMORIA` informa RSS do coordenador, total e mapa por PID dos
  workers, Edge, total, staging e ocupação da janela.
* São expostos tempo de starvation do downloader e utilização acumulada por
  slot. PID/thread do worker, PID do coordenador, sequência e ID técnico
  continuam presentes em `TELEMETRIA_SLOT`.
* Ao atingir 90% de memória ou menos de 512 MiB disponíveis, novas reservas são
  retidas; tarefas já iniciadas terminam e nenhuma versão é descartada.

## Metodologia

Foram geradas 30 versões XLSX idênticas entre as oito configurações, em banco e
staging novos para cada execução. CQLPA120 sintética usou 250 linhas/versão e
CQLPA123 sintética, 1.000 linhas/versão. O tempo inclui startup e warm-up — não
foi descartada uma janela inicial curta, para evitar escolher arbitrariamente
uma amostra estável em execuções de poucos segundos. Os JSON completos incluem
médias, p50/p95, commit, espera de promoção, ocupação, utilização, starvation,
retries, erros e equivalência oficial.

O container não contém as planilhas reais, credenciais SharePoint nem Edge.
Portanto estes são benchmarks controlados do pipeline XLSX/SQLite, **não** uma
alegação de benchmark real do WebDriver. `rss_edge` é zero por construção; a
decisão deve ser reconfirmada na CQLPA123 real antes de distribuição ampla.

## Resultado CQLPA120 controlado

| Slots | Tempo (s) | Comparações/min | Speedup | Eficiência | CPU média | RSS total MiB |
|---:|---:|---:|---:|---:|---:|---:|
| 1 | 1,934 | 899,573 | 1,000 | 1,000 | 19,96% | 55,06 |
| 2 | 1,203 | 1.445,876 | 1,608 | 0,804 | 28,54% | 80,75 |
| 3 | 1,198 | 1.452,714 | 1,614 | 0,538 | 32,26% | 105,99 |
| 4 | 1,369 | 1.270,698 | 1,413 | 0,353 | 37,11% | 130,98 |
| 5 | 1,044 | 1.666,822 | 1,852 | 0,370 | 42,60% | 155,77 |
| **6** | **0,852** | **2.041,194** | **2,270** | **0,378** | **45,59%** | **181,07** |
| 7 | 1,042 | 1.669,266 | 1,856 | 0,265 | 45,59% | 206,31 |
| 8 | 1,077 | 1.615,944 | 1,796 | 0,225 | 45,80% | 231,79 |

## Resultado CQLPA123 controlado (workload principal)

| Slots | Tempo (s) | Comparações/min | Speedup | Eficiência | CPU média | RSS total MiB |
|---:|---:|---:|---:|---:|---:|---:|
| 1 | 3,872 | 449,426 | 1,000 | 1,000 | 23,78% | 58,72 |
| 2 | 2,591 | 671,444 | 1,494 | 0,747 | 35,49% | 88,88 |
| **3** | **2,280** | **763,019** | **1,698** | **0,566** | **43,10%** | **117,08** |
| 4 | 2,322 | 749,360 | 1,668 | 0,417 | 42,77% | 145,59 |
| 5 | 2,560 | 679,599 | 1,512 | 0,302 | 42,22% | 174,24 |
| 6 | 2,770 | 628,228 | 1,398 | 0,233 | 41,96% | 202,89 |
| 7 | 3,003 | 579,351 | 1,289 | 0,184 | 40,05% | 232,34 |
| 8 | 3,011 | 577,957 | 1,286 | 0,161 | 38,75% | 260,66 |

## Integridade, falhas e controles

Todas as configurações produziram exatamente os mesmos technical IDs, pares,
alterações e checkpoint. O teste agressivo mantém as sequências rápidas staged
atrás da primeira lenta. O teste de restart descarta staging órfão e retoma do
checkpoint oficial sem duplicação. Pause com oito slots bloqueia reservas no
boundary e retoma; Stop com oito slots preserva o checkpoint. A suíte também
confirma que o modelo temporal compartilhado exclui pausa do relógio.

## Conclusão e recomendação

O ponto ótimo observado foi 6 slots no workload leve e 3 no pesado. Em CQLPA123,
4–8 elevaram RSS sem ganho; de 5 em diante houve regressão de throughput. Isso
indica contenção mista de parsing/processos/overhead no ambiente controlado, não
RAM exaurida, SQLite ou download real. O default passa a **3**, por ser o melhor
resultado do workload prioritário, e não 8. O usuário pode selecionar 1–8.

Riscos remanescentes: o download serial e RSS Edge só podem ser medidos no
ambiente SharePoint real; amostras curtas têm variância; CPU pico via `psutil`
depende da disponibilidade do pacote; e 8 processos elevam RSS quase
linearmente. Antes de promover o default organizacional, repetir 30–50 versões
reais da CQLPA123 e CQLPA120, com Edge ativo e máquina ociosa, usando o mesmo
intervalo e os JSON desta fase como esquema de coleta.
