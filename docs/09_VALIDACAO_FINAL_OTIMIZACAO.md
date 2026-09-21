# Validação final da otimização

Data: 2026-09-21. Base validada: `0150389` (`Instrument and harden
SharePoint transfers`, equivalente ao commit aprovado após integração). Ambiente:
Linux, Python local, SQLite local, sem Edge e sem credenciais SharePoint. As medidas
abaixo são **SINTÉTICAS**, salvo indicação explícita; não foram misturadas com uma
medição real de rede.

## Veredito

**APROVADO PARA VALIDAÇÃO REAL**.

A suíte, o soak de 5.000 comparações, a retomada transacional, a integridade do
SQLite, o leitor tipado e o relatório streaming não revelaram regressão de
integridade. A aprovação significa que o pacote está pronto para o ensaio
operacional em um SharePoint autorizado; não significa que latência, Edge ou
throttling reais tenham sido validados neste contêiner.

Durante a revisão foi demonstrado um risco de encerramento: a thread daemon de
trabalho compartilha banco e fonte com a interface, enquanto o `finally` de
`main` fecha ambos. Foi aplicada a única correção desta rodada: enquanto `_busy`
for verdadeiro, o fechamento pela janela é recusado com mensagem clara. Depois
da operação, o fechamento permanece normal. Não foi introduzido cancelamento
forçado.

## Cobertura funcional e de falhas

A suíte completa tem 135 testes. Os testes de auditoria comprovam rollback da
unidade versão/alterações/checkpoint, ausência de órfãos, preservação do último
checkpoint confirmado e retomada exatamente no par incompleto. A versão já
confirmada não é refeita. Há injeção depois dos inserts e antes do checkpoint,
além de falha de download antes da primeira confirmação.

A aquisição SharePoint simulada cobre falha do `fetch`, erro de prefetch com
fallback, offset e tamanho de chunk inválidos, base64 inválido, ZIP inválido ou
incompleto, ausência de `workbook.xml`, divergência de digest e limpeza de
`.part`. SHA-256 tem teste binário determinístico. O leitor cobre shared strings,
inline/rich strings, fórmulas normais e compartilhadas, datas/epochs, booleanos,
erros, números, strings, vazios e fallback openpyxl, sempre verificando valor e
tipo Python.

A persistência cobre inserts de versão, alterações e checkpoint na mesma
transação SQLite. Restrições de unicidade e FKs são testadas. A publicação do
relatório é atômica: falha de validação mantém o relatório anterior e remove o
artefato parcial. Particionamento, limite físico do Excel, ordem global, totais,
ADD/MOD/DEL, aba Integridade e reabertura read-only são cobertos.

Os pontos “fechamento durante download/prefetch/leitura/comparação/persistência/
relatório” usam agora uma única política verificável: o fechamento é bloqueado
durante qualquer operação de fundo. Assim não existe fechamento concorrente de
Edge, SQLite ou temporários. Os testes de workspace cobrem limpeza normal,
limpeza após exceção e limpeza na próxima inicialização. WinError 32 não pode ser
reproduzido no Linux; a política evita deliberadamente o caminho concorrente que
o produz no Windows.

## Soak test sintético

Foram processadas 5.001 versões (5.000 pares), alternando dois XLSX válidos de
100 células e 5.557 bytes. Cada par produziu uma alteração e percorreu aquisição,
SHA-256, leitura, comparação, transação e checkpoint.

| Métrica | Resultado |
|---|---:|
| Comparações / alterações | 5.000 / 5.000 |
| Tempo total | 10,963 s |
| Tempo médio por par | 2,193 ms |
| Aquisições | 5.001 |
| RSS inicial / final | 28,18 / 30,61 MiB |
| Variação de RSS | +2,43 MiB |
| SQLite inicial / final | 72 / 1.308 KiB |
| Temporários observados | 3 constantes (2 XLSX + banco) |
| Checkpoint final | `5000` |
| `integrity_check` | `ok` |
| Violações FK / alterações órfãs | 0 / 0 |

Foram feitas 21 amostras, a cada 250 pares. `gc.get_count()` oscilou entre as
gerações em vez de crescer monotonicamente; o RSS cresceu apenas 8,6% e não
acompanhou o volume de 5.000 snapshots. Os snapshots anteriores foram liberados,
mantendo somente o par corrente. O workspace cresceu pelo SQLite confirmado,
não por temporários. O Edge e cache de endereços não existem neste cenário; o
snapshot manteve 100 endereços.

## Benchmarks antes/depois

| Etapa | Antes | Depois | Natureza / observação |
|---|---:|---:|---|
| Enumeração | não registrado | não medido | **REAL necessário**; paginação/retry cobertos por teste |
| Download normal | não registrado | não medido | **REAL necessário** |
| Prefetch | inexistente | não medido | **REAL necessário**; consumo, fallback e digest cobertos |
| Espera residual | não registrado | não medido | **REAL necessário** |
| Leitor XLSX | ~3,235 s | 0,0430 s / 3.000 células | **SINTÉTICO**, amostra pequena; 58,06% vs leitor anterior no mesmo arquivo |
| Comparador | ~1,538 s | incluído em 2,193 ms/par | **SINTÉTICO** de 100 células; não é comparação direta de volume |
| Banco | não isolado | incluído em 2,193 ms/par | **SINTÉTICO** |
| Relatório 100 mil | ~134,93 s / ~358 MiB | 37,41 s / 34,38 MiB RSS | **SINTÉTICO**, medição física já registrada |
| Memória da auditoria | não registrado | 28,18→30,61 MiB | **SINTÉTICO**, 5.000 pares |
| Total por versão | não registrado | 2,193 ms | **SINTÉTICO**, XLSX pequeno e local |

O microbenchmark do leitor confirmou igualdade exata do snapshot tipado e mediu
102,62 ms no caminho anterior contra 43,04 ms no atual (uma repetição após
warm-up). Essa amostra serve como smoke benchmark; o valor oficial anterior de
3,235 s veio de outro volume e não deve ser usado para calcular ganho percentual.

O ensaio físico de 1.000.000 de alterações foi tentado: a carga no SQLite chegou
ao fim em 8,74 s, porém o executor encerrou o processo durante a exportação, antes
de produzir resultado validável. Por isso ele não é apresentado como sucesso.
O maior relatório físico concluído e reaberto nesta máquina permanece o ensaio
de 100.000 alterações (37,41 s, 34,38 MiB RSS, 3,21 MiB XLSX). Testes de
particionamento cobrem volumes acima de dois milhões sem materializar o arquivo.

## Integridade e regressões

Ao fim do soak, `PRAGMA integrity_check` retornou `ok` e
`PRAGMA foreign_key_check` retornou vazio. Não houve alteração órfã, duplicidade
de par nem avanço indevido do checkpoint. A suíte completa voltou a obter 135
sucessos; `compileall` e `git diff --check` também passaram.

Não foi encontrada regressão funcional. A única deficiência confirmada foi o
encerramento concorrente, corrigido pelo bloqueio explícito. Não houve mudança em
algoritmos de comparação, leitor, persistência, relatório ou transferência nesta
rodada.

## Alterações acumuladas e ganhos

As rodadas anteriores introduziram persistência/checkpoint transacionais,
comparação incremental com apenas dois snapshots, leitor XLSX XML em passagem
única com equivalência tipada, relatório write-only/batched com publicação
atômica, cache de enumeração, download SharePoint em chunks validáveis, SHA-256
do binário adquirido, prefetch de uma versão e telemetria de tempos. Os ganhos
quantificados são a queda do relatório de 134,93 s para 37,41 s e do pico de
aproximadamente 358 MiB para 34,38 MiB, além de 58,06% no microbenchmark pareado
do leitor.

## Gargalos, riscos e recomendações

1. Rede, autenticação, throttling, tamanho de blobs e RSS do Edge dependem do
   tenant e continuam sem medição real.
2. Base64 e a cópia Edge→Python continuam tendo custo proporcional ao XLSX;
   prefetch reduz espera, não bytes transferidos.
3. O relatório de um milhão de linhas exige uma janela operacional superior à
   permitida por este executor e espaço para banco, `.part.xlsx` e arquivo final.
4. Não force o encerramento pelo Gerenciador de Tarefas. Aguarde a operação; a
   interface agora impede o fechamento normal enquanto ela estiver ativa.
5. Na validação real, registrar separadamente enumeração, download, espera de
   prefetch, RSS Python/Edge e total por versão; repetir com ao menos 20 versões
   reais após warm-up.
6. Manter banco e temporários em disco local, reservar espaço livre superior a
   duas vezes o maior relatório e conservar backup antes de cargas extensas.
7. Se cancelamento durante auditoria se tornar requisito, implementar token de
   cancelamento cooperativo entre pares; não encerrar a thread à força.
