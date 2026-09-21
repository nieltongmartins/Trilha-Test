# Benchmark do relatório streaming

Medição em 2026-09-21 no contêiner de desenvolvimento, com SQLite local, 100.000
alterações `MOD`, uma versão e uma execução. O pico é o RSS máximo do processo
(`resource.getrusage`), portanto inclui Python, SQLite e openpyxl. O arquivo foi
reaberto em `read_only=True` e todas as validações de integridade passaram.

| Implementação | Alterações | Tempo | Pico de memória | XLSX | Validação |
|---|---:|---:|---:|---:|---|
| anterior (referência fornecida) | 100.000 | ~134,93 s | ~358,39 MiB (pico Python) | não registrado | não aplicável |
| streaming | 100.000 | 37,41 s | 34,38 MiB RSS | 3,21 MiB | VALIDADO |

Isso representa, nesta máquina, redução aproximada de 72% no tempo e de 90% na
métrica de pico reportada (as métricas de memória não são rigorosamente idênticas).
O cenário físico de um milhão de linhas não foi executado nesta rodada: a projeção
do teste de 100 mil excederia o tempo disponível. Os limites até mais de dois
milhões são cobertos por testes parametrizados de particionamento, combinados com
testes reais de streaming, reabertura, contagens, ordem e publicação atômica.
