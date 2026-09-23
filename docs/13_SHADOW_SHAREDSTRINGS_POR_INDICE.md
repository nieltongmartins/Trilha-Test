# Shadow mode: dependência granular de `sharedStrings` por índice

## Regra anterior e excesso de invalidação

A assinatura de cada worksheet é composta pelo SHA-256 dos bytes da própria
worksheet, pelo SHA-256 completo de `sharedStrings.xml` e `styles.xml`, pelo
epoch, pelo SHA-256 dos relacionamentos e pelo target. O reuso por row só monta
o índice de rows anteriores quando **toda** a cauda dessa assinatura é igual.
Consequentemente, acrescentar um único `<si>` impedia o índice anterior de ser
consultado e todas as rows eram parseadas, mesmo quando seus bytes e todos os
índices referenciados continuavam iguais.

## Arquitetura experimental (não ativa)

`SharedStringsSnapshot` mantém somente a versão anterior e a atual. Ele contém:

* a presença explícita do membro (ausente é diferente de vazio);
* o valor textual usado pelo leitor;
* por posição, SHA-256 da serialização da árvore XML completa de `<si>`;
* `None` para uma entrada cuja estrutura não pertence ao conjunto conhecido.

A árvore inclui texto simples, runs e propriedades de rich text, `xml:space`,
Unicode, elementos fonéticos e atributos. Entidades são resolvidas pelo parser
XML. Mudanças de aparência de rich text são, intencionalmente, mudanças da
assinatura mesmo quando o texto concatenado é igual. Elementos/extensões
ignorados pelo leitor rápido não recebem prova: isso produz parsing normal, não
reuso. Não se usa regex para interpretar shared strings.

Cada `_CachedRow` registra o conjunto ordenado dos índices das células `t="s"`
realmente interpretadas. Células com fórmula não registram o `<v>` em cache como
dependência, porque o leitor retorna a fórmula. Índices negativos, não inteiros
ou inexistentes atravessam a barreira segura de recurso não suportado.

No segundo arquivo, somente uma row com o mesmo SHA-256 dos bytes, mesmas demais
dependências e igualdade de assinatura em **todos** os índices usados é marcada
como segura pelo shadow. Ausência, remoção, troca, reordenação, estrutura
ignorada ou divergência exata contra a row recém-parseada invalida a candidata.
A posição é parte da prova: procurar a mesma string em outro índice é proibido.
Styles, epoch, relacionamentos, target e fórmulas compartilhadas mantêm as
barreiras anteriores.

O resultado oficial continua deliberadamente conservador nesta etapa. Mesmo
uma candidata segura é parseada, e seu mapa completo (coordenadas, tipos,
valores, fórmulas, datas, booleanos, `None`, zero e strings) é comparado ao mapa
anterior. Só a telemetria informa o reuso potencial; identidade de objeto e o
comparador não foram alterados.

## Telemetria

Foram adicionadas contagens de totais anterior/atual, índices iguais, alterados,
novos e removidos; candidates/safe/invalidated; índices consultados, mudados e
novos; rows dependentes, potencialmente reutilizáveis e invalidadas por índice,
dependência global ou estrutura desconhecida; e tempos de hash e diff. Elas são
emitidas na linha `PERF` junto das métricas existentes.

## Validação e custo observado

A suíte sintética cobre identidade, append de uma e várias entradas, alteração,
remoção, reordenação, rich text, whitespace/`xml:space`, Unicode, fonética,
estrutura desconhecida, ausência, corrupção, índice inexistente, rows com
índice estável/alterado e igualdade com o leitor oficial. Testes preexistentes
continuam cobrindo styles, epoch, fórmulas, fallback e igualdade integral.

Neste ambiente não estão disponíveis os XLSX CQLPA123 5.119/5.120/5.121 nem
5.219/5.220/5.221; portanto não se inventaram percentuais ou benchmarks reais.
Um ensaio isolado com 30.000 entradas simples mediu aproximadamente 1,2–1,3 s
sem instrumentação de memória. Com `tracemalloc` (que aumenta muito o custo), o
pico observado foi 11,6 MiB e a memória retida 5,0 MiB. O custo por índice
inclui digest hexadecimal e referência em tupla; as dependências por row
crescem conforme o número de índices distintos efetivamente usados.

## Decisão de ativação e riscos

Não é seguro ativar o reuso nesta etapa: o modo é apenas sombra, como solicitado.
Antes de uma flag experimental, é necessário obter equivalência 100% nos pares
reais, medir o custo dentro do pipeline e investigar qualquer OOXML encontrado
fora do conjunto conhecido. O principal risco residual é falso negativo por
serializações semanticamente equivalentes mas diferentes; ele reduz performance,
sem comprometer auditoria. Novas construções OOXML permanecem conservadoras.
