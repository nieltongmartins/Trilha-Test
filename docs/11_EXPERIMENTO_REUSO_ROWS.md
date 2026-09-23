# EXP-ROW-002 — snapshots persistentes por row

Data: 2026-09-23. Decisão: **REFINAR**, mantendo a implementação no laboratório.

## Diagnóstico e hipótese

Os dados operacionais fornecidos para CQLPA123 mostram cerca de 1.120.098
células convertidas sempre que qualquer byte de uma das duas grandes worksheets
muda, embora o comparador encontre poucas alterações. O experimento preserva o
atalho por worksheet do EXP-WS-001 e adiciona uma unidade persistente por
`<row>`: bytes XML idênticos produzem o mesmo objeto de row somente após SHA-256.
Não se usa CRC, tamanho, posição ZIP ou número de células como prova.

Não havia arquivos reais, credenciais SharePoint nem os bancos A/B no ambiente
de desenvolvimento. Por isso **não há resultado real, p90/p95 operacional,
melhor tempo SharePoint ou equivalência de banco a declarar nesta execução**.
Essas medições não devem ser inferidas do ensaio sintético abaixo.

## Arquitetura implementada

Cada worksheet incremental é uma `RowSheetSnapshot`, uma visão imutável sobre
mapas de células por row. Uma versão alterada varre os segmentos XML de row,
calcula SHA-256 dos bytes completos e compartilha o mapa anterior quando
coincidem chave estrutural e digest. Só rows diferentes passam por ElementTree,
conversão tipada e alocação de células. O comparador reconhece mapas de row com
identidade compartilhada e não os percorre.

A autorização depende também da igualdade exata das dependências já usadas pelo
leitor: `sharedStrings.xml`, `styles.xml`, epoch, relacionamentos e target. Uma
mudança em qualquer dependência invalida todas as rows. Fórmulas compartilhadas
mantêm parsing integral porque mestre e seguidores atravessam rows; array e
dataTable continuam no fallback openpyxl. Formatos XML não conservadores também
caem no leitor seguro. Uma amostra inicial com menos de 10% de rows reutilizadas
abandona cedo o índice incremental para limitar o pior caso.

O snapshot público continua oferecendo a interface `Mapping` (`[]`, `get`,
`items`, igualdade com `dict`), sem materializar novamente um dicionário plano.
O cache continua limitado à versão imediatamente anterior.

## Benchmark sintético reproduzível

Máquina Linux do ambiente de desenvolvimento; uma aba com 10.000 rows x 5
colunas (50.000 células), uma célula numérica alterada, três amostras aquecidas.
O tempo inclui ZIP, dependências, hashing, diff estrutural, parsing e snapshot.

| Caminho | amostras (s) | mediana | redução da mediana |
|---|---|---:|---:|
| leitor oficial | 0,6260; 0,6086; 0,5426 | 0,6086 s | — |
| incremental por row | 0,1636; 0,1119; 0,1460 | 0,1460 s | 76,0% |

Na última amostra: 10.000 rows totais, 9.999 reutilizadas, uma parseada,
49.995 células reutilizadas e cinco parseadas. O diff estrutural levou 0,122 s,
o parsing da row 0,0002 s e o snapshot 0,0017 s. O resultado foi comparado ao
leitor oficial e a alteração produzida pelo comparador foi idêntica.

Este corpus é menor e menos variado que CQLPA123. Percentis com apenas três
amostras seriam enganosos; p90/p95 não são publicados. RSS também não é usado
como medida de memória porque arenas do Python e execuções no mesmo processo
impedem atribuição confiável. Estruturalmente, 9.999 mapas foram compartilhados
e somente um novo mapa foi alocado.

Um segundo ensaio alterou todas as 50.000 células. As medianas de três
amostras foram 0,4207 s no leitor oficial e 0,4338 s no incremental, regressão
de 3,1%. O abandono antecipado ocorreu depois de 256 rows; por isso 51.280
células foram efetivamente convertidas (as 1.280 da amostra conservadora mais
as 50.000 do parsing integral). O resultado atende o limite experimental de
5–10% neste corpus, mas ainda precisa ser confirmado nos arquivos reais.

## Equivalência e cenários

A suíte automatizada cobre igualdade total, modificação, adição, remoção,
False versus zero, data1904, numFmt/data, fórmula normal e compartilhada,
array/dataTable, shared/inline/rich strings, aba adicionada/removida e Unicode.
Os novos testes verificam especificamente 99/100 rows compartilhadas com uma
célula alterada e invalidação total quando `styles.xml` muda. ZIP corrompido,
membro ausente e XML inválido continuam propagando falha, sem aceitar snapshot
parcial.

A suíte completa obteve 195 testes aprovados. Não foi executado teste A/B de
banco com versões reais; o código de transação, checkpoint, hash global,
prefetch e fonte SharePoint não foi alterado.

## Limitações e próxima decisão

* O primeiro arquivo precisa construir o índice de rows e pode custar mais que
  o leitor plano; o ganho é esperado nas versões seguintes com alta retenção.
* Shared formulas desativam reuso parcial da worksheet inteira.
* Mudanças globais de styles/shared strings invalidam conservadoramente todas as
  rows, mesmo se poucos índices forem usados.
* O scanner incremental aceita rows SpreadsheetML sem prefixo e preserva no
  fragmento todas as declarações de namespace da raiz da worksheet. Rows
  prefixadas ou estruturas que ainda não possam ser isoladas com segurança
  causam fallback explícito, com motivo na telemetria.
* A medição real de 30–100 versões e o banco A/B dependem dos XLSX e credenciais
  que não estão no repositório.

Recomendação: **continuar/refinar**, condicionada ao ensaio real. Manter somente
se a equivalência oficial e banco A/B forem 100%, a mediana real cair ao menos
50% e o pior caso ficar dentro de 10%. Caso qualquer condição falhe, abandonar
o caminho por row sem alterar o leitor oficial.

## Correção de compatibilidade com namespaces reais

Arquivos reais do SharePoint expuseram `ParseError: unbound prefix` porque a
primeira implementação parseava cada `<row>` em um wrapper que declarava apenas
o namespace default. A worksheet real declarava na raiz prefixos usados dentro
da row (`r`, `mc`, `x14`, `x14ac`, `xr` e extensões do produtor), portanto o
fragmento isolado perdia parte do contexto XML.

O leitor agora captura com o próprio parser XML todas as declarações `xmlns` da
raiz e as aplica, sem alterar o fragmento, somente ao wrapper temporário usado
para interpretação. O SHA-256 continua sendo calculado sobre os bytes originais
da row. Prefixos não são removidos, renomeados ou ignorados. Erros de parsing do
fragmento são convertidos em incompatibilidade incremental: o cache é limpo, o
leitor openpyxl oficial assume o workbook e `fallback=True` registra também o
motivo.

O teste de regressão combina namespace default, `r`, `mc`, `x14`, `x14ac`,
`xr`, um namespace desconhecido válido, atributos prefixados e elementos
prefixados na mesma worksheet. Ele comprova equivalência integral com o leitor
oficial, reutilização da row idêntica e parsing exclusivo da row alterada.

Depois da correção, uma repetição do benchmark de 50.000 células obteve
medianas de 0,4494 s no leitor oficial e 0,1067 s no incremental (redução de
76,26%). Foram novamente reutilizadas 9.999 rows/49.995 células e parseadas uma
row/cinco células. A captura de namespaces da raiz não apresentou regressão
mensurável nesse ensaio sintético.
