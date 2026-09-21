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
