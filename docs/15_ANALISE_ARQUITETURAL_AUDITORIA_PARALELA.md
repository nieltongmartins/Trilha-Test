# Experimento de alto throughput — análise arquitetural da auditoria paralela

**Repositório:** `Trilha-Test` (exclusivamente experimental)

**Fase:** 1 — análise, sem implementação do pipeline paralelo

**Data:** 2026-09-24

**Decisão:** aguardar aprovação antes de iniciar a Fase 2

> Esta proposta não se aplica ao repositório oficial. Ela não altera regras de
> comparação, identidade técnica, fórmulas, watermark, relatório ou histórico.
> Não há benchmark real de 1–5 slots nesta fase; qualquer número futuro deve ser
> medido com o mesmo intervalo da CQLPA123.

## Resumo executivo e decisão recomendada

O motor atual é corretamente serial: percorre pares adjacentes, conserva o
snapshot anterior, lê o atual, compara e persiste comparação e checkpoint numa
transação antes de iniciar o próximo par. O prefetch sobrepõe rede a CPU, mas não
torna leituras ou comparações simultâneas.

A experiência deve separar três identidades e três filas:

1. **Version task:** baixar, validar e produzir um snapshot imutável, identificado
   por `(run_id, spreadsheet_id, technical_version_id, sequence)`;
2. **Comparison task:** consumir os snapshots das sequências `i-1` e `i` e
   produzir um resultado staged para a sequência `i`;
3. **Promotion:** um único coordenador promove exclusivamente o maior prefixo
   contíguo staged, em transações oficiais ordenadas.

Recomenda-se começar com **um único ator de aquisição**, dono exclusivo do
WebDriver, e paralelizar parsing/comparação em **processos independentes**. Essa
escolha é mais segura que múltiplos Edge e precisa ser confirmada por benchmark
contra threads, pois serialização de snapshots e pressão de memória podem anular
o ganho do multiprocessing. A Fase 2 deve aceitar somente 1 ou 2 slots; 3–5
ficam bloqueados até equivalência integral e testes de falha passarem.

---

## 1. Arquitetura atual relevante

### Descoberta e identidade

`BrowserSharePointSource` usa uma sessão Edge visível, autenticação/MFA manual e
`fetch` GET same-origin executado pelo Selenium. A descoberta percorre pastas e
arquivos. A enumeração de versões:

- prefere consulta incremental inclusiva a partir do **ID técnico** do checkpoint;
- usa `VersionLabel` somente para validar a fronteira e apresentar a versão;
- pagina com `nextLink` ou `$skip`, rejeita repetição, duplicidade e ordem não
  monotônica;
- lê metadados antes/depois e rejeita a lista se o watermark mudar;
- inclui a versão do checkpoint para que o primeiro par pendente tenha antecessor.

`VersionInfo.id` é a identidade técnica. A posição na lista normalizada, depois
da validação e ordenação técnica, define `sequence`. O label nunca deve virar
chave do scheduler.

### Download, prefetch, SHA e ZIP

O Edge faz o `fetch`; Python transfere o `Blob` em chunks base64 de 512 KiB para
um `.part`, calcula SHA-256 incremental durante a escrita, valida tamanho e XLSX
ZIP e faz rename atômico. Depois, `AuditService` calcula novamente o SHA do
arquivo e pede à fonte para conferir esse digest contra o digest incremental.
Há retry limitado, fallback de URL histórica, recuperação da sessão Edge e
limpeza best-effort.

O prefetch mantém até dois blobs no contexto JavaScript. Ele começa após a
versão corrente chegar ao disco, enquanto Python calcula SHA, lê e compara. Os
slots têm token, URL, ID técnico, label e tamanho esperado; prefetch duplicado é
evitado dentro desse buffer. Apesar de downloads poderem prosseguir no Edge, as
chamadas Python ao mesmo WebDriver continuam originadas por um fluxo serial.

### Reader, snapshots e comparação

Existe um `ConsecutiveWorkbookReader` mutável. Ele guarda cache da versão
anterior, assinaturas de worksheets/rows e snapshot de shared strings para
reutilizar objetos somente depois de prova por SHA-256. Incompatibilidades
derrubam o cache e usam fallback openpyxl. O caminho independente
`read_workbook()` não depende de versão anterior.

Snapshots são mapeamentos em memória por aba/endereço, com valores tipados e
algumas estruturas de rows imutáveis. `compare_snapshots()` é puro em relação a
fonte e banco: recebe dois snapshots e devolve `CellChange`.

### Persistência e checkpoint

O SQLite possui uma conexão compartilhada aberta com
`check_same_thread=False`, porque hoje somente uma thread de auditoria a usa de
forma serial. Para cada par, uma transação insere `versao_processada`, todas as
`alteracao` e atualiza `checkpoint`; a constraint única do par evita duplicação.
O checkpoint só é anunciado à UI depois do commit. Uma falha faz rollback e
registra a execução como falha sem perder o último prefixo confirmado.

### Pause, stop, interface e erros

A interface Tk executa a auditoria em uma thread daemon e recebe eventos por
filas, sempre aplicados pela main thread. Hoje há uma barra global e uma barra da
versão corrente. Pause e stop são cooperativos **entre comparações já
confirmadas**. Pause cancela prefetch e espera; stop encerra a execução como
interrupção graciosa. Uma exceção em qualquer etapa encerra toda a auditoria.

## 2. Por que a execução atual é serial

O loop possui dependência explícita: `previous_snapshot = current_snapshot` ao
fim de cada iteração. A mesma instância de reader depende exatamente dessa ordem;
a mesma conexão SQLite persiste e avança o checkpoint; e a mesma fonte/driver
executa download, prefetch e recuperação. Assim, paralelizar o `for` atual seria
incorreto: criaria corrida no reader, WebDriver, conexão, prefetch e checkpoint.

O prefetch é apenas sobreposição I/O–CPU. Ele não materializa snapshots fora de
ordem nem desacopla comparação de promoção.

## 3. Unidade ideal de trabalho

### `VersionTask`

Campos mínimos imutáveis:

- `run_id`, `spreadsheet_id`, `sequence`;
- `technical_version_id` (chave), `version_label` (display);
- `predecessor_technical_id`;
- metadados originais de `VersionInfo`;
- número da tentativa e timestamps.

Resultado: `SnapshotHandle`, SHA-256 confirmado, tamanho e métricas. A tarefa
termina em `SNAPSHOT_READY`; ela não escreve histórico oficial.

### `ComparisonTask`

Chave: `(run_id, spreadsheet_id, predecessor_id, current_id)`. Torna-se elegível
quando os dois `SnapshotHandle` estão prontos. Resultado: lista/stream de mudanças,
contagem, hashes/identidades dos inputs e métricas. Termina em `STAGED`; não
avança checkpoint.

### `PromotionTask`

Não é executada por slot. É uma operação exclusiva do coordenador para a próxima
sequência após o checkpoint. Ela valida novamente as identidades e promove um
resultado staged em transação oficial.

Essa separação evita chamar “versão concluída” de “comparação comprometida”.
Estados devem ter significado preciso:

`PENDING -> RESERVED -> RUNNING -> SNAPSHOT_READY` (versão),
`COMPARISON_READY -> RUNNING -> STAGED` (par), e `COMMITTED/CHECKPOINTED`
(promoção atômica; externamente os dois últimos ocorrem juntos). `FAILED` inclui
tipo de tarefa, tentativa e causa.

## 4. Dependência entre versões

Construir um DAG linear após discovery. O nó de comparação `i` depende somente
dos snapshots `i-1` e `i`, não da comparação `i-1`. Portanto, se snapshots
5.121/5.122 estiverem prontos, o par pode ser comparado mesmo com 5.120 lento.
Apenas a promoção `i` depende da promoção `i-1`.

O snapshot inicial é o checkpoint (ou a primeira versão quando ainda não há
checkpoint); ele deve ser materializado, mas não conta como comparação. Cada
snapshot mantém refcount igual ao número de comparações ainda não staged que o
usam (no máximo duas). Ao zerar, é liberado.

## 5. Conclusão física fora de ordem

O scheduler mantém mapas por **ID técnico** e sequence. Workers publicam eventos
imutáveis numa fila central; nunca mutam diretamente o grafo. Um resultado só é
aceito se `run_id`, reservation token, ID e tentativa coincidirem com a reserva
ativa. Resultado tardio de worker cancelado é descartado.

Após qualquer `SNAPSHOT_READY`, o scheduler enfileira todos os pares recém
elegíveis. Após `STAGED`, o slot fica livre imediatamente para version/comparison
e o coordenador é notificado. “Aguardando promoção” é contador global do
resultado, não ocupação do slot.

## 6. Commit e checkpoint estritamente em ordem

Há exatamente um `OrderedCommitCoordinator`, na thread de coordenação, e somente
ele recebe a capacidade de escrita oficial. Seu algoritmo é:

1. ler do banco o ID técnico do checkpoint confirmado;
2. localizar a próxima sequence no plano imutável;
3. se o resultado do par não está `STAGED`, parar;
4. validar predecessor/current IDs, SHA, run e contagem;
5. `BEGIN IMMEDIATE` no banco oficial;
6. inserir `versao_processada` e `alteracao`;
7. atualizar o checkpoint para o ID/label atual;
8. atualizar contadores da execução;
9. `COMMIT`; só então publicar checkpoint e marcar `COMMITTED/CHECKPOINTED`;
10. repetir enquanto houver prefixo contíguo staged.

Falha em qualquer passo faz rollback do par atual. O coordenador para novas
promoções, preserva o checkpoint anterior e pode deixar outros trabalhos já em
curso chegarem a boundary seguro. Nunca deve haver uma operação independente de
“atualizar checkpoint”. Constraints oficiais continuam como defesa idempotente.

Embora seja possível promover vários pares numa única transação, recomenda-se
uma transação por par na Fase 2: reduz lock, mantém o boundary atual e torna a
retomada observável. O loop ainda pode drenar rapidamente vários pares prontos.

## 7. Modelo de staging

### Escolha para a Fase 2

Usar **SQLite temporário por execução**, separado do banco oficial, para
resultados de comparação e metadados; manter snapshots imutáveis somente sob
backpressure (memória ou artefato temporário medido). Tabelas devem usar prefixo
explícito `exp_` e conter `complete=0/1`, IDs técnicos, sequence, hashes dos
inputs e reservation token. Só marcar `complete=1` na mesma transação que grava
todas as mudanças staged.

O staging não é fonte de verdade e não é anexado à transação oficial. O
coordenador lê um resultado completo e abre uma transação **somente no banco
oficial**. Isso evita depender de atomicidade multiarquivo do SQLite.

### Política de crash escolhida

Escolher **B: descartar e recalcular staging no restart**. No startup, arquivos
de run sem marcador de encerramento são apagados; o plano é reconstruído a
partir do checkpoint oficial. Downloads `.part`, snapshots e resultados
incompletos nunca são aceitos. É a alternativa mais simples e segura. Recuperar
staging fica fora da primeira experiência porque exigiria validar formato,
versão do código, inputs e durabilidade de cada artefato.

## 8. Scheduler central e work stealing

Um único scheduler possui a máquina de estados e filas; workers apenas pedem e
devolvem trabalho. Reserva é atômica sob o lock do scheduler e gera token único.
Não há partição fixa por slot.

Prioridade sugerida:

1. comparação elegível mais antiga, para liberar snapshots e promoção;
2. version task mais antiga dentro da janela e necessária para fechar a lacuna;
3. demais version tasks elegíveis em ordem;
4. retry vencido, sem ultrapassar limite.

O scheduler não deve deixar um worker parado só porque a promoção está bloqueada.
Contudo, respeita três semáforos: `max_inflight_downloads=1` inicialmente,
`max_parse_workers=slots` e orçamento de snapshots/bytes. A janela à frente é
calculada em sequences, não em labels: começar como parâmetro de ensaio
`max_ahead = 2 * slots`, sem promovê-lo a default definitivo antes do benchmark.

Para impedir download duplicado, o registro central por
`(spreadsheet_id, technical_version_id)` possui uma única reserva e um único
future. Requisições subsequentes aguardam esse future/handle; jamais chamam a
fonte novamente.

## 9. Threads versus processos

### Recomendação

- **Thread/ator de aquisição:** única, pois Selenium e I/O são externos e o
  driver deve ter dono exclusivo.
- **Processos de parsing/comparação:** candidato principal, pois o parser percorre
  XML/células e contém trabalho Python CPU-bound; processos evitam o GIL e isolam
  estado/falha do reader.
- **Thread coordenadora + main Tk:** eventos e UI, sem trabalho pesado.

Não afirmar ganho antecipadamente. Retornar um snapshot de 1,55 milhão de células
por pickle pode duplicar memória e consumir IPC. A prova técnica da Fase 2 deve
comparar, no mesmo conjunto aquecido: (a) executor síncrono, (b) 2 threads com
readers independentes e (c) 2 processos. Se IPC dominar, testar processo que
produz artefato imutável local e comparação no mesmo processo/por caminhos, sem
alterar semântica. `spawn` deve ser usado no Windows; payloads devem ser objetos
de dados, nunca conexão SQLite, source ou WebDriver.

## 10. Estratégia Selenium

Adotar inicialmente a **Opção A**: uma sessão Edge e um ator que serializa **toda**
chamada ao WebDriver, inclusive discovery, status de prefetch, chunks, limpeza e
recovery. Nenhum worker chama `BrowserSharePointSource` diretamente.

O buffer JS atual pode continuar sobrepondo até dois fetches, mas toda interação
Python permanece pelo ator. Para o primeiro protótipo, é aceitável desabilitar o
prefetch especulativo e usar uma fila de aquisição serial, estabelecendo baseline
correto; em seguida reabilitá-lo sob o mesmo ator e medir. Múltiplos Edge por slot
só merecem experimento separado se aquisição virar gargalo, com medição de RAM,
CPU, login/MFA, throttling e estabilidade. Não faz parte da Fase 2.

Se o Edge fechar, o ator tenta a recuperação limitada já existente. Se falhar,
marca a version task, bloqueia a lacuna, cancela novas aquisições e solicita stop
seguro; processos que já receberam XLSX válido podem finalizar, mas nada além da
lacuna é promovido.

## 11. Estratégia do reader

Não compartilhar `ConsecutiveWorkbookReader`: seu cache `_previous` e shared
strings representam ordem local e seriam corrompidos por atribuição dinâmica.

Na Fase 2, usar **leitura completa independente por versão** (`read_workbook`) e
snapshot imutável. Isso sacrifica reutilização incremental para obter uma
baseline paralela correta e comparável. Um reader independente por worker ainda
não é suficiente para reutilização: work stealing não garante que um worker
receba versões consecutivas. Afinidade opcional pode ser estudada depois, sem
impedir roubo, e só se benchmark provar vantagem.

Cada snapshot deve carregar ID técnico, SHA do XLSX e schema version. O comparador
recebe apenas snapshots finalizados; nenhum cache mutável cruza processo.

## 12. Memória, temporários e backpressure

Não é possível estimar bytes confiáveis apenas de “1,55 milhão de células”:
objetos Python variam por tipo, strings, fórmulas e compartilhamento. Antes de
definir 3–5 slots, medir `peak RSS / células` em versões reais. Como aproximação
de capacidade, usar o **pico observado**, não uma constante teórica:

`budget_snapshot = floor((RAM_utilizavel - RSS_base - margem_Edge) / pico_snapshot)`.

Instrumentar a cada transição:

- RSS do coordenador, de cada processo e soma da árvore Edge;
- número e bytes estimados/serializados de snapshots;
- XLSX, `.part` e staging em disco;
- tamanho do buffer/prefetch; fila, janela e slots ativos;
- picos e motivo de backpressure.

Limites simultâneos: `max_ahead`, `max_snapshot_count`, orçamento RSS e espaço
livre em disco. Parar novas version tasks ao atingir qualquer limite, mas dar
prioridade a comparisons que liberem snapshots. Depois que os pares adjacentes
necessários ficam staged, decrementa-se refcount e fecha/remove-se o handle. O
snapshot do checkpoint só pode ser liberado após stage do primeiro par.

## 13. Pause e stop

### Pause

`RUNNING -> PAUSING -> PAUSED`. Ao pedir pausa, o scheduler deixa de reservar
novas tarefas e cancela apenas aquisição/prefetch ainda cancelável. Download em
rename/validação, escrita transacional de staging e promoção oficial terminam seu
boundary. Parsers em andamento terminam e publicam snapshot; comparisons em
andamento terminam o stage. Quando `active=0` e não há transação, todos os slots
ficam PAUSED. Staging permanece até resume (mas continua descartável em crash).

### Stop

`RUNNING/PAUSED -> STOPPING -> STOPPED`. Não há novas reservas. Jobs ainda não
iniciados são cancelados; aquisição cooperativamente cancelável é abortada e
limpa. Jobs CPU já despachados podem terminar dentro de timeout; seus resultados
só são aceitos se reservation token continuar válido. A promoção já iniciada
termina commit/rollback. Depois o pool é encerrado, staging/temporários são
apagados e a execução é finalizada com checkpoint oficial preservado. Stop não
gera `erro_processamento`.

## 14. Recuperação de falhas e retries

- **Falha de versão:** `FAILED`, com tentativa e backoff. Duas tentativas totais
  é um ponto inicial conservador, configurável apenas para erros classificados
  como transitórios. ZIP, SHA, identidade ou parsing determinístico não devem
  repetir indefinidamente.
- **Falha de comparação:** retry uma vez em processo novo; divergência repetida
  é fatal para a lacuna.
- **Worker morto:** invalida sua reserva, remove artefato incompleto e reencaminha
  se ainda houver retry. Os demais concluem seus boundaries.
- **Erro SQLite staging:** descarta transação staged; oficial não muda.
- **Erro SQLite oficial:** rollback, congela promoção e inicia stop seguro; não
  continua acumulando trabalho sem perspectiva de commit.
- **Crash/máquina reiniciada:** banco oficial determina checkpoint. Staging e
  temporários do run interrompido são descartados; pares não oficiais são
  recalculados. Constraints tornam uma tentativa repetida detectável, mas o
  plano deve começar estritamente depois do checkpoint.

Uma versão `FAILED` nunca é pulada para promoção. Trabalho independente já em
curso pode virar staged, sujeito aos limites, mas o estado global mostra a
lacuna e oferece retry/restart em vez de declarar sucesso.

## 15. Throughput, métricas e ETA

Registrar timestamps monotônicos de queue/reserve/start/end por fase e timestamp
de commit. Métricas mínimas:

- versões materializadas/min e comparações staged/min;
- comparações committed/min (throughput efetivo);
- duração e espera p50/p95 por download, SHA, parsing, dependency, comparison,
  staging e commit;
- duração por slot, ocupação, retries/erros;
- CPU, RSS Python total/por worker, RSS Edge, snapshots e disco.

A taxa recente deve ser uma janela móvel temporal (por exemplo, últimos 5
minutos ou 20 commits, o que contiver amostra suficiente), calculada pelo
conjunto: `rate = commits_na_janela / segundos_da_janela`. ETA principal:

`ETA = comparações_pendentes_oficiais / committed_rate_sustentada`.

Antes dos primeiros commits, exibir “calculando” e, apenas como diagnóstico,
`staged_rate`; não estimar pela soma das durações. Se uma lacuna interromper
commits, a taxa de commit deve decair com o tempo e a UI indicar “bloqueada na
versão X”, em vez de prometer ETA otimista. Speedup será
`wall_time_1 / wall_time_N` para intervalos idênticos e mesmo estado de cache.

## 16. Desenho da interface

Antes da execução, um `Combobox readonly` “Processamentos simultâneos” aceita 1
ou 2 na Fase 2 (1–5 somente após generalização); valor fica bloqueado durante a
auditoria. Recomenda-se default experimental **2**, não 3, para reduzir risco no
primeiro ensaio.

A área global mostra:

- barra de `committed / total`;
- “N comparações oficialmente confirmadas”; checkpoint por label **e ID técnico
  em detalhe/log**, pendentes, processando e staged aguardando ordem;
- slots ativos/configurados, throughput committed recente e ETA;
- RAM Python/Edge, snapshots ativos e backpressure quando acionado.

Cada slot tem `slot_id`, VersionLabel (display), ID técnico opcional em tooltip,
etapa normalizada, barra e cronômetro. Etapas: aguardando, download, validação
SHA, leitura XLSX, aguardando dependência, comparação, concluída e erro.
“Aguardando promoção” deve aparecer no resumo/staging, não prender o slot. Ao
atingir 100%, mostrar “Concluída — 12,4 s” por curto intervalo e aceitar já o
próximo evento; a UI não deve atrasar o scheduler para sustentar a animação.

Workers publicam `SlotProgress` imutável por fila. Somente a main thread Tk muda
widgets. Eventos carregam run/token e sequence; a UI ignora evento atrasado de
uma reserva anterior do mesmo slot.

## 17. Testes necessários

1. **Unitários do DAG/scheduler:** reserva exclusiva, token tardio, work stealing,
   prioridade, janela, refcount e transições inválidas.
2. **Ordem solicitada:** atrasos 30/2/3/1 s; 5.121–5.123 ficam staged, checkpoint
   permanece 5.119 e drena 5.120–5.123 após fechar a lacuna.
3. **Dependência:** comparação só aparece com ambos snapshots, mas independe da
   comparação precedente.
4. **Atomicidade:** injetar falha após versão, no meio das mudanças e antes do
   checkpoint; nenhum parcial ou avanço.
5. **Crash:** estado oficial em 5.120, 5.121/5.122 staged e 5.123 running; matar
   processo, reiniciar, descartar staging, recalcular sem duplicar e chegar ao
   mesmo checkpoint.
6. **Falhas isoladas:** parser morre, download/Edge fecha, ZIP/SHA inválido,
   staging cheio e banco oficial bloqueado; nunca pular lacuna.
7. **Pause/stop em cada boundary:** reserva, fetch, `.part`, SHA, parse, compare,
   stage e commit; ausência de thread/processo/arquivo órfão.
8. **WebDriver:** fake detecta chamadas concorrentes; o máximo deve ser 1.
9. **Deduplicação:** múltiplas dependências pedem mesmo ID, apenas um download.
10. **Backpressure:** primeira versão bloqueada e centenas posteriores; memória,
    snapshots e janela permanecem limitados.
11. **Equivalência A/B automatizada:** exportar bancos de 1 e N slots em ordem
    canônica e comparar IDs técnicos, pares, status, SHA, contagens, tipo/aba/
    endereço, valores anterior/novo e checkpoint. Ignorar somente IDs surrogate,
    timestamps e códigos de execução esperadamente diferentes.
12. **Reader:** snapshot independente equivale byte semanticamente ao reader
    atual/fallback em strings, tipos, datas e fórmulas.
13. **UI:** número de cards, selector bloqueado, eventos tardios, contadores e
    ETA sem acesso Tk fora da main thread.
14. **Benchmark real:** mesmas 20–50 versões CQLPA123, mesma máquina, cache frio e
    aquecido separados, ordem 1/2/3/4/5 randomizada/repetida; registrar todas as
    métricas pedidas e integridade A/B.

## 18. Riscos e mitigação

| Risco | Efeito | Mitigação / gate |
|---|---|---|
| Snapshot/IPC enorme | RAM maior e speedup negativo | orçamento por pico real; threads vs processos; artefato local |
| GIL nas threads | pouco paralelismo | benchmark pareado; processos como candidato |
| Disco/serialização em processos | parsing ganho, total pior | medir wall time e bytes IPC; não generalizar sem ganho |
| Reader incremental fora de ordem | resultado/caches incorretos | leitura independente na Fase 2 |
| WebDriver concorrente | corrupção/hang | ator único e teste de concorrência máxima 1 |
| SharePoint throttling | retries/instabilidade | um download por vez inicialmente; telemetria/backoff |
| SQLite `check_same_thread=False` | transações cruzadas | escritor oficial único; workers sem conexão oficial |
| Lacuna antiga | acúmulo ilimitado | janela, budgets e prioridade para fechar/liberar |
| Staging confundido com oficial | checkpoint falso | banco temporário separado e UI com contadores distintos |
| Resultado tardio após cancel | stage indevido | run/reservation token e máquina de estados central |
| Crash entre stage e commit | repetição | descartar staging; checkpoint oficial como verdade |
| Alteração de versão durante discovery | plano inconsistente | preservar watermark e rejeição atuais |
| Múltiplos Edge | RAM/MFA/throttling | fora da Fase 2; experimento isolado posterior |
| UI saturada por eventos | lentidão/flicker | coalescer último evento por slot na main thread |

## 19. Speedup teórico, sem inventar benchmark

Se uma versão gasta aproximadamente `D` em aquisição serial, `P` em parsing
paralelizável e `C` em comparação/stage, um limite otimista de pipeline é:

`throughput(N) <= 1 / max(D, (P + C) / N, commit_cost)`.

Isso já mostra por que o ganho não é linear: discovery/download continuam
serializados; commit, IPC e memória adicionam overhead; CPU, disco e Edge são
recursos compartilhados. Pela lei de Amdahl, se `s` é a fração estritamente
serial medida no baseline:

`speedup(N) <= 1 / (s + (1-s)/N)`.

Com os intervalos informados pelo solicitante (download ~0,8–1,3 s, leitura
~10–13 s, comparação ~0,6–0,8 s), parsing parece dominar e existe potencial
teórico para ganho substancial com 2–3 processos. Isso **não** autoriza estimar
um fator: GIL, IPC, pressão de RAM, contenção de CPU/disco e perda do reader
incremental ainda são desconhecidos. O ótimo pode ser 2, 3, 4 ou 5; também é
possível 5 ser pior. Somente `wall_time_1 / wall_time_N` real, acompanhado de
equivalência exata, decide.

## 20. Plano de implementação por etapas e gates

### Fase 1 — este documento

- arquitetura e invariantes revisadas;
- nenhuma implementação definitiva, schema oficial ou UI alterados;
- aprovação explícita necessária para prosseguir.

### Fase 2A — núcleo determinístico, ainda 1 slot

1. Introduzir tipos de task/event/state e DAG sob feature flag experimental.
2. Extrair `OrderedCommitCoordinator`, mantendo SQL/semântica atuais.
3. Implementar staging temporário descartável e limpeza de startup.
4. Implementar executor síncrono de 1 slot; provar equivalência com o motor atual.
5. Adicionar testes de ordem, atomicidade, crash e comparação A/B.

**Gate:** nenhum diff semântico e suíte/falhas passando.

### Fase 2B — exatamente 2 slots

1. Criar ator único de aquisição e deduplicação por ID técnico.
2. Adicionar executor de dois workers, reader independente e backpressure.
3. Benchmarkar threads versus processos; selecionar por throughput, RSS e
   estabilidade, não por pressuposto.
4. Adicionar dois cards, selector 1/2, métricas agregadas, pause/stop e retry.
5. Validar 1 vs 2 slots no mesmo intervalo real.

**Gate:** igualdade total dos bancos, checkpoint ordenado, zero concorrência no
driver e ganho estável sem exceder orçamento de memória.

### Fase 3 — generalização 1–5

Generalizar configuração/cards/pool sem mudar invariantes. Executar testes de
lacuna e falha com cada N. Não permitir mudança de N durante uma execução.

**Gate:** equivalência de cada N contra 1 slot e limites de memória/disco.

### Fase 4 — benchmark e decisão

Executar 1–5 slots nas mesmas 20–50 versões reais, separar frio/aquecido,
repetir e registrar wall time, versões/comparações por minuto, fases, espera,
CPU/RSS/disco, erros e retries. Calcular speedup e eficiência, escolher o ponto
ótimo observado. Se nenhum N>1 cumprir integridade e ganho, desativar o
experimento sem migrá-lo ao projeto oficial.

## Invariantes de aceite

1. ID técnico, nunca VersionLabel, identifica e reserva trabalho.
2. Um WebDriver tem exatamente um dono e nenhuma chamada concorrente.
3. Nenhuma instância mutável de reader cruza workers.
4. Staged não significa committed/checkpointed.
5. Somente o coordenador escreve histórico oficial e checkpoint.
6. Comparação oficial + mudanças + checkpoint são uma transação.
7. O checkpoint representa sempre o maior prefixo contíguo confirmado.
8. Crash pode perder trabalho especulativo, nunca evidência oficial confirmada.
9. Pause/stop operam em boundaries; não matam transação.
10. Paralelismo só é aprovado com equivalência A/B exata.
