# EXP-ROW-003 — diagnóstico de falsos negativos do SHA de rows

Data: 2026-09-23. Decisão atual: **medir antes de alterar a assinatura**.

## Resultado honesto desta execução

O checkout experimental não contém versões reais da CQLPA123 (nem qualquer
arquivo XLSX). Portanto ainda não é possível atribuir uma causa principal real,
publicar a porcentagem pedida para 20–50 pares ou recomendar a troca definitiva
do SHA-256. Fazer isso com fixtures sintéticas seria apresentar hipótese como
medição operacional.

Foi adicionada somente a instrumentação offline
`tools/diagnose_row_hashes.py`. Ela torna reproduzível o ensaio assim que o
corpus autorizado estiver disponível e não modifica leitor, cache, SharePoint,
prefetch, banco, checkpoint, interface, relatório ou persistência. Um pequeno
ensaio sintético de validação produziu intencionalmente uma row de cada classe
A/B/C; ele valida o instrumento e a proteção de dados, **não diagnostica a
CQLPA123**.

## Protocolo para os 20–50 pares reais

Executar em uma cópia local, ordenada por versão:

```bash
python tools/diagnose_row_hashes.py /caminho/CQLPA123 \
  --max-pairs 50 --sample-limit 5 --output diagnostico-rows.json
```

O comando ordena naturalmente os nomes, forma pares adjacentes e limita o
corpus a 50 pares. Antes da execução deve-se confirmar que essa ordem coincide
com a ordem técnica das versões; se os nomes não a representarem, devem ser
passados explicitamente em ordem. O JSON informa `pair_count`, permitindo
recusar conclusões quando houver menos de 20 pares.

Para cada worksheet presente nos dois arquivos, o diagnóstico:

1. extrai os mesmos bytes `<row>` usados pelo cache atual e calcula SHA-256;
2. lê ambos os arquivos pelo leitor oficial em produção;
3. agrupa o snapshot oficial por número de row e compara valor e **tipo Python**
   (por exemplo, `False` não é igual a `0`);
4. classifica A (SHA diferente/lógica igual), B (SHA diferente/lógica
   diferente) e C (SHA igual/lógica igual);
5. inspeciona a árvore XML somente para nomear diferenças estruturais;
6. mede hashing, classificação estrutural, leitura do oráculo, comparação
   semântica, tempo total e pico de alocação Python (`tracemalloc`).

`raw_hash_false_negative_percent` tem denominador `sha_different` e numerador
`different_logically_equal`. `rows_total` inclui somente ocorrências de row com
a mesma chave (número e ordinal) nos dois lados. Adições/remoções são separadas
em `rows_added_or_removed`. Assim o percentual não mistura falta de
correspondência com falso negativo de igualdade.

## Privacidade e amostragem

As amostras automáticas A/B/C registram somente:

* aba, número de row e hashes;
* tamanhos XML e contagens de células/elementos;
* presença de fórmula e estilo;
* nomes de atributos alterados e categorias estruturais.

Valores, textos inline, texto de fórmulas e valores de atributos nunca entram
no JSON. Esses dados existem somente na memória do oráculo durante o processo.
O teste automatizado inclui marcadores confidenciais e comprova que eles não
aparecem no relatório.

## O que a classificação consegue demonstrar

As categorias são cumulativas, e sua frequência conta rows, não eventos XML:

| Categoria | Evidência observada | Interpretação inicial |
|---|---|---|
| `serializacao_lexica` | árvores expandidas idênticas, bytes diferentes | ordem de atributos, prefixo/declaracão de namespace ou forma lexical; candidata a irrelevância, ainda sujeita ao contexto |
| `whitespace` | árvore igual após remover somente whitespace estrutural | potencialmente irrelevante fora de texto de célula |
| `atributos_row` | atributo do elemento row mudou | exige análise por nome; `hidden`, `ht`, `s` e `customFormat` não podem ser descartados por suposição |
| `atributos_celula` / `estilo_id` / `tipo_celula` | metadado de célula mudou | potencialmente relevante à conversão ou auditoria |
| `atributos_formula` / `formula` | fórmula ou metadado de fórmula mudou | semanticamente relevante; fórmula compartilhada pode depender de outra row |
| `valor_xml_ou_shared_string_id` | texto de `<v>` mudou | pode ser valor real ou apenas renumeração de shared string; depende do arquivo global |
| `texto_inline` | conteúdo inline mudou | normalmente relevante |
| `elementos_ou_ordem` | sequência de tags mudou | precisa de análise específica; ordem de células e extensões não têm a mesma semântica |
| `extensao_ou_namespace_externo` | existe elemento fora do namespace principal | desconhecido até provar a semântica da extensão |
| `estrutura_nao_classificada` | diferença restante | bloqueia normalização até investigação manual segura |

A classificação não chama automaticamente nenhuma categoria de “segura”. Ela
serve para localizar a causa dominante no corpus e então escrever testes de
equivalência direcionados.

## Dependências globais e limites do oráculo

Por par, o relatório separa igualdade de bytes de `sharedStrings.xml`,
`styles.xml`, relationships e metadata de workbook, além de comparar `numFmts`,
`cellXfs` e epoch 1900/1904. Uma row com bytes iguais **não é reutilizável** só
por isso quando muda uma dependência que possa reinterpretá-la:

* índice `t="s"` depende de `sharedStrings`;
* `s` depende de `cellXfs` e `numFmt`, especialmente para datas;
* números interpretados como datas dependem também do epoch;
* fórmulas compartilhadas dependem de mestre/seguidores possivelmente em
  outras rows;
* relationships/target identificam qual worksheet está sendo comparada;
* metadata de workbook deve ser avaliada por campo, nunca ignorada em bloco.

O “semântico” deste experimento significa exatamente o snapshot consumido pela
auditoria atual: coordenada, tipo Python, valor tipado ou fórmula. Estilos
puramente visuais não fazem parte desse snapshot. O instrumento registra a
presença/alteração estrutural de estilos, mas não deve ser usado para alegar
equivalência visual completa do Excel.

## Critério de decisão após a execução real

Somar os 20–50 pares e publicar, sem arredondar contagens:

```text
rows_total=
sha_equal=
sha_different=
different_logically_equal=
different_logically_changed=
raw_hash_false_negative_percent=
```

Também devem ser publicadas frequência das categorias, frequência de mudanças
globais e distribuição por par (para evitar que um único arquivo domine a
média). Se a maioria das rows com SHA diferente for logicamente diferente, o
cache atual está fazendo a escolha correta. Se a maioria for igual, existe
oportunidade, mas a autorização para reuso ainda dependerá das entradas globais
efetivamente referenciadas.

## Alternativas a avaliar — não implementadas

| Estratégia | Segurança | Custo/complexidade | Impacto esperado | Riscos e invalidações |
|---|---|---|---|---|
| Assinatura da representação tipada | alta se produzida pelo mesmo conversor | alto: em geral paga o parsing que se quer evitar | elimina falsos negativos lexicais | shared strings, styles/numFmt, epoch e fórmulas; pouco ganho se materializar valores |
| Hash por célula relevante | alta com tokens e dependências explícitas | médio/alto; scanner adicional e índice por célula | bom quando poucas células mudam dentro de muitas rows | células vazias, ordem/duplicatas, shared formulas, extensão desconhecida |
| Hash de tokens XML relevantes | potencialmente alta após prova por categoria | médio; parser streaming sem alocar objetos finais | pode ignorar ruído de row e preservar ganho | lista incompleta de tokens cria falso positivo, que é inaceitável |
| Canonicalização XML (C14N) | resolve apenas diferenças lexicais | médio e provavelmente caro em 60 mil rows | proporcional à frequência de ordem/prefixos/whitespace | não resolve IDs globais nem autoriza ignorar atributos; regex é proibida |
| Assinatura semântica canônica híbrida | melhor potencial: tokens locais + digest das dependências referenciadas | alta complexidade de projeto e testes | evita parsing completo quando a causa dominante é ruído ou renumeração | índice de estilos/strings, fórmulas cruzando rows e formatos desconhecidos |

Recomendação provisória: se o corpus confirmar muitos falsos negativos, testar
primeiro um **scanner de tokens relevantes com dependências referenciadas**, em
paralelo ao SHA bruto e sem conceder reuso. Ele tende a custar menos que montar
a representação tipada completa. C14N pode ser uma baseline de medição, mas não
é, isoladamente, uma assinatura semântica.

## Próximo experimento e barreiras de segurança

1. Executar o comando sobre 20–50 pares reais autorizados e arquivar somente o
   JSON sanitizado.
2. Revisar as dez categorias/atributos dominantes e selecionar apenas aquelas
   cuja irrelevância ao snapshot oficial possa ser provada.
3. Implementar uma assinatura candidata em modo sombra: ela apenas registra
   “reutilizaria”, enquanto o leitor continua fazendo parsing oficial.
4. Exigir zero falso positivo contra o oráculo, inclusive nos pares com mudança
   de shared strings, styles, numFmt, epoch, relationships e fórmulas.
5. Comparar custo da assinatura com o parsing evitável (mediana e p95), além do
   pico de memória. Abandonar se a assinatura custar tanto quanto converter as
   células.
6. Somente depois disso considerar uma mudança de cache, ainda restrita ao
   repositório experimental.

Não foi feita nesta tarefa nenhuma normalização nem alteração definitiva do
mecanismo de reutilização.
