# Profiling isolado do leitor XLSX

## Escopo e segurança

`tools/benchmark_xlsx_reader.py` é uma ferramenta diagnóstica independente. Ela
não instancia `AuditService`, não abre SQLite, não acessa SharePoint, não gera
relatório e não lê nem grava checkpoint. O leitor oficial agora visita os
filhos diretos de cada célula uma única vez; a implementação anterior com
`find()` repetido permanece privada e acessível somente ao benchmark e testes.

O benchmark executa primeiro `read_workbook` e usa o snapshot resultante como
oráculo. Além das três variantes instrumentadas, mede o caminho real anterior
e o `read_workbook` otimizado após um warm-up:

1. `instrumented_current`: mesmas funções de conversão e buscas `find()` do
   loop atual, mas sobre XML previamente descompactado para separar o custo do
   ZIP do custo de `ElementTree.iterparse`;
2. `direct_children`: protótipo que visita uma vez os filhos `<f>`, `<v>` e
   `<is>` da célula;
3. `direct_children_gc_disabled`: o mesmo protótipo com o GC suspenso somente
   durante as worksheets e obrigatoriamente restaurado em `finally`.
4. `consecutive_sheet_reuse`: protótipo sequencial, sem threads/processos, que
   mantém `SheetSnapshot` por membro ZIP e só o reutiliza após equivalência
   criptográfica de todas as entradas que participam da interpretação.

As variantes não são importadas pelo produto. O processo termina com código 1
se qualquer snapshot divergir, inclusive quando valores iguais tiverem tipos
Python diferentes.

## Execução

Para um diretório contendo as versões reais da mesma planilha:

```bash
python tools/benchmark_xlsx_reader.py /caminho/versoes \
  --repeat 5 --output /tmp/profile-xlsx.json
```

É recomendável fechar outros processos intensivos, usar ao menos 20 versões,
executar em disco local e repetir a coleta após uma rodada de aquecimento. O
JSON de saída deve ficar fora do repositório e não deve ser anexado à auditoria.

## Métricas

Por arquivo, a saída contém tamanho ZIP, RSS antes/depois, tempos individuais e
mediana do leitor oficial. `official_before_after` contém média, mediana, mínimo,
máximo, amostras, células, RSS e equivalência tipada dos dois caminhos reais.
Por variante contém RSS, GC, total, equivalência,
ganho, worksheets, células armazenadas, XML descompactado, shared strings e
entradas do snapshot.

O bloco `consecutive_sheet_reuse` mede, para a sequência ordenada, tempo do
leitor oficial, tempo do protótipo, leitura ZIP, hashing, parsing/conversão,
fechamento, RSS, pico de alocações Python via `tracemalloc`, células processadas,
células efetivamente alteradas e abas
reaproveitadas. Cada aba informa CRC-32 e tamanhos do diretório ZIP para
diagnóstico, além do SHA-256 usado na prova. **CRC e tamanho nunca autorizam
reuso**: colisões de CRC são possíveis e timestamp não é consultado.

Cada worksheet registra bytes comprimidos/descompactados, células vistas e
armazenadas, fórmulas, traduções de fórmulas compartilhadas, contagem por tipo e
tempos de descompressão, `iterparse`, coordenada, `_cell_value` equivalente,
inserção no `dict`, fórmula, data, booleano, erro e demais conversões.

Os cronômetros internos são **inclusivos**: `iterparse_hot_loop` inclui o loop,
conversões e inserções; `cell_value` inclui suas subclasses. Portanto eles não
devem ser somados. O tempo oficial é a referência de desempenho. Variantes
instrumentadas servem para atribuição relativa e comparação entre protótipos,
pois a própria instrumentação tem custo.

O RSS em Linux vem de `/proc/self/statm`; em plataformas sem uma fonte barata e
portável o campo é `null`. A memória é medida no mesmo processo e pode reter
arenas do Python entre variantes. Para máxima precisão de pico, execute um
arquivo por processo e complemente com a ferramenta de memória do sistema.

## Critério de equivalência

Uma otimização só pode permanecer quando todas as versões reportarem
`snapshot_exactly_equal: true`. A revisão também deve cobrir abas, coordenadas,
valores/tipos, fórmulas comuns e compartilhadas, datas nos dois epochs,
booleanos, erros, vazios explícitos, shared strings e inline strings. Uma futura
alteração do leitor exigirá os mesmos testes unitários, regressão completa e
benchmark repetível nos arquivos reais.

## Modelo de dependências e proposta

O snapshot já é logicamente `dict[título, dict[coordenada, valor]]`; portanto o
protótipo conserva objetos por aba, monta apenas o dicionário externo da nova
versão e descarta referências a abas removidas ao final da leitura. Não há
cópia das células de uma aba reaproveitada. Renomear uma aba troca somente a
chave externa, desde que o mesmo target e suas dependências continuem iguais.

A assinatura conservadora contém:

* SHA-256 dos bytes descompactados da worksheet (logo fórmulas, fórmulas
  compartilhadas, tipos, estilos referenciados e valores precisam ser
  byte-a-byte idênticos);
* SHA-256 de `sharedStrings.xml` e `styles.xml` completos — incluindo `numFmt`,
  `cellXfs` e qualquer estrutura de estilo ainda não interpretada;
* epoch efetivo (`date1904`) obtido de `workbook.xml`;
* SHA-256 completo de `workbook.xml.rels` e o target da worksheet.

Essa escolha invalida mais do que o mínimo: por exemplo, adicionar uma relação
de outra aba força releitura de todas. É intencional nesta fase. Ausência de um
membro também faz parte da assinatura e não equivale a arquivo vazio. Mudança
em qualquer dependência implica parsing. A lista de shared strings e os estilos
de data da versão atual são sempre reconstruídos, evitando que um índice seja
interpretado com a tabela anterior.

Relações próprias da worksheet, hyperlinks, drawings, tabelas, nomes definidos,
`calcChain` e resultados em cache não alteram o snapshot do leitor atual: ele
preserva a expressão de `<f>` e lê apenas o valor lógico da célula. Antes de
ampliar o contrato do snapshot, essas estruturas devem entrar na assinatura ou
desabilitar o reuso. Recursos não suportados (`array`/`dataTable`) continuam
exigindo o fallback oficial; o protótipo não deve ser promovido até espelhar
esse fallback por arquivo inteiro.

## Streaming e decisão de adoção

Comparar durante `iterparse` é tecnicamente possível: cada célula convertida
pode ser confrontada imediatamente com a coordenada anterior, enquanto se
constrói somente o novo `SheetSnapshot`. Isso reduz uma passagem de comparação,
mas não elimina o novo snapshot, necessário para a versão seguinte e para
detectar coordenadas removidas. A implementação foi deliberadamente adiada:
ela mistura leitura e regras de diff e aumenta o risco sobre fórmulas, tipos e
ordenação de eventos.

O protótipo não está ligado ao pipeline, checkpoint ou banco. Uma promoção só
deve ocorrer depois de benchmark no corpus real mostrar ganho material e testes
cobrirem: uma célula, aba inteira, shared strings, styles, numFmt, fórmula,
data, adição, remoção e renomeação de aba. O resultado tipado deve ser idêntico
ao leitor oficial em todas as versões; qualquer dúvida ou membro inesperado
deve resultar em releitura, nunca em reuso otimista.
