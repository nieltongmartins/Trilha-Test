# EXP-WS-001 — reuso criptográfico de worksheets

Data: 2026-09-23. Decisão: **MANTER**.

## Hipótese e arquitetura

O repositório já continha aplicação, suíte funcional, corpus sintético, leitor
XLSX oficial e protótipo diagnóstico; portanto não foi criada uma referência do
zero. O leitor oficial continua sendo o oráculo. O experimento promove ao fluxo
da auditoria o reuso de `SheetSnapshot` entre versões consecutivas, sem alterar
o formato lógico `dict[aba, dict[célula, valor]]`.

Uma aba somente recebe o mesmo objeto Python da versão anterior quando coincidem:

* SHA-256 dos bytes descompactados da worksheet;
* SHA-256 de `sharedStrings.xml` e `styles.xml` completos;
* epoch efetivo (`date1904` ou 1900);
* SHA-256 de `workbook.xml.rels` completo;
* target do membro da worksheet.

Membro ausente é representado por `missing` e nunca equivale ao SHA-256 de um
membro vazio. CRC-32 e tamanho não participam da autorização. Qualquer mudança
na assinatura refaz parsing e conversão da aba. Fórmula não suportada aciona o
mesmo fallback openpyxl do oráculo e limpa integralmente o cache. O cache guarda
somente a versão imediatamente anterior.

O comparador encerra a análise de uma aba quando os mapas são o mesmo objeto.
No pipeline otimizado essa identidade nasce da prova acima. Mesmo para chamadas
externas, o atalho permanece semanticamente seguro: um objeto não pode diferir
de si próprio.

## Medição sintética controlada

Corpus local gerado com openpyxl: uma worksheet, 20.000 linhas por 5 colunas,
100.000 células numéricas, duas cópias byte a byte idênticas. Foram feitas seis
leituras por caminho após a criação do arquivo; números incluem abertura ZIP,
descompressão e SHA-256. Ambiente Linux local, sem Edge, rede ou SQLite.

| Caminho | Amostras (ms) | Mediana | Média | p95 observado | Máximo |
|---|---|---:|---:|---:|---:|
| leitor oficial | 890,88; 814,27; 790,98; 745,43; 774,00; 775,56 | 783,27 | 798,52 | 890,88 | 890,88 |
| reuso seguro | 63,65; 22,99; 23,91; 22,90; 22,80; 22,76 | 22,94 | 29,83 | 63,65 | 63,65 |

O ganho de mediana foi **97,07%**. A primeira amostra do caminho incremental é
um outlier de aquecimento conservado na tabela. O leitor reportou uma aba
reutilizada e zero células parseadas. Memória não foi promovida como resultado:
o ensaio ocorreu no mesmo processo e arenas do Python impedem atribuição segura
por RSS; estruturalmente não há cópia do mapa de 100.000 células.

## Equivalência, riscos e limites

A suíte compara o resultado incremental ao `read_workbook`, incluindo valores e
tipos já cobertos pelo corpus existente. Um caso consecutivo idêntico comprova
reuso das três abas e zero células parseadas; um par alterado comprova
invalidação. A regressão completa obteve 190 testes aprovados.

O ganho depende de worksheets idênticas entre versões. Se `sharedStrings` ou
`styles` mudar, a política conservadora invalida todas as abas, mesmo que a
mudança não seja referenciada por uma delas. Isso reduz hit rate, nunca
integridade. O benchmark é sintético e não permite afirmar tempo operacional de
2–3 segundos: fetch, Edge→Python, base64 e arquivos reais ainda precisam da
coleta autorizada descrita em `08_PROFILING_LEITOR_XLSX.md`.

Não houve alteração no banco. Versão, alterações e checkpoint continuam na
mesma transação. Não houve paralelismo nem comandos Selenium concorrentes.

## Próximos experimentos

1. Executar o profiler em pelo menos 20 versões reais, mantendo os JSON fora do
   Git, e registrar mediana, p90, p95, máximo, RSS e taxa de reuso por aba.
2. Medir Edge→Python com 256 KiB, 512 KiB, 1 MiB e 2 MiB antes de mudar chunks.
3. Somente se o hit rate real for baixo, avaliar parser alternativo com fallback
   e equivalência tipada integral.

