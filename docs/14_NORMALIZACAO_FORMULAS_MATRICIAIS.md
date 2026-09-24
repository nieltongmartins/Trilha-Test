# Normalização de fórmulas matriciais

## Caminho investigado

Na versão instalada do openpyxl (3.1.5), uma célula de fórmula matricial lida
com `data_only=False` expõe `cell.value` como `ArrayFormula`, com o texto em
`text` e o intervalo em `ref`. `DataTableFormula` é o outro tipo especial de
fórmula da mesma versão e expõe `ref`, `ca`, `dt2D`, `dtr`, `r1`, `r2`, `del1`
e `del2`.

O leitor XML rápido rejeita deliberadamente fórmulas `array` e `dataTable` e
aciona `_read_openpyxl`. Antes desta correção, `_read_openpyxl` copiava
`cell.value` diretamente para o snapshot. Assim, o objeto chegava ao
comparador, onde duas instâncias distintas não eram iguais, e depois a
`AuditService._serialize`, que aplicava `str(value)`. O texto com endereço de
memória era então gravado em `alteracao.valor_anterior`/`valor_novo`; o
relatório apenas exportava fielmente esse texto do SQLite. Portanto, o defeito
ocorria antes do relatório, na criação do snapshot e na serialização.

## Correção

`normalize_formula_value` é a normalização central. Ela preserva sem alteração
fórmulas comuns (strings), números, datas, strings, booleanos, vazio e `None`.
Os tipos especiais são codificados como JSON compacto, Unicode e com chaves
ordenadas, precedido pelo tipo:

```text
ARRAYFORMULA|{"formula":"=SUM(A1:A2)","ref":"B1:B2"}
DATATABLEFORMULA|{"ca":false,"del1":false,"del2":false,"dt2D":false,"dtr":false,"r1":null,"r2":null,"ref":"B1:B2"}
```

O leitor normaliza antes de construir o snapshot. O comparador também aplica a
mesma função como barreira defensiva para snapshots fornecidos diretamente, e
a persistência a reaplica antes de converter valores escalares em texto. Logo,
snapshot, comparação, SQLite e relatório compartilham a mesma representação.

## Histórico SQLite

Foi pesquisado o repositório por bancos SQLite existentes e por ocorrências de
`ArrayFormula object at`; não havia banco de auditoria versionado para
inspecionar nem registro histórico encontrado. A correção não migra, apaga ou
reescreve dados existentes. Caso essa representação seja encontrada em um
banco de produção, sua eventual migração deve ser tratada separadamente.
