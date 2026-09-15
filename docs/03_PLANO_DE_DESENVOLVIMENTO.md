# PLANO DE DESENVOLVIMENTO
## Auditor de Planilhas Excel — SharePoint Online

**Documento:** 03_PLANO_DE_DESENVOLVIMENTO.md  
**Versão:** 1.0  
**Status:** Oficial  
**Data:** 15/09/2026  

---

# 1. OBJETIVO

Este documento define o plano oficial de desenvolvimento da V1 do
Auditor de Planilhas Excel armazenadas no SharePoint Online.

O desenvolvimento deverá ser:

- curto;
- incremental;
- testável;
- rastreável;
- funcional;
- orientado a entregas reais.

O objetivo não é construir uma plataforma excessivamente complexa.

O objetivo é entregar uma ferramenta simples e confiável que:

1. selecione uma planilha;
2. consulte seu histórico de versões;
3. compare versões consecutivas;
4. identifique alterações;
5. armazene a trilha;
6. mantenha checkpoint;
7. continue auditorias anteriores;
8. gere relatório consolidado.

---

# 2. DOCUMENTOS OFICIAIS

Antes de implementar uma fase, consultar:

1. `docs/01_ESPECIFICACAO_FUNCIONAL.md`
2. `docs/02_ARQUITETURA.md`
3. `docs/03_PLANO_DE_DESENVOLVIMENTO.md`
4. `docs/04_HISTORICO_IMPLEMENTACOES.md`

Responsabilidades:

`01_ESPECIFICACAO_FUNCIONAL.md`
- define O QUE deve ser construído.

`02_ARQUITETURA.md`
- define COMO a solução deverá ser estruturada.

`03_PLANO_DE_DESENVOLVIMENTO.md`
- define EM QUE ORDEM será construída.

`04_HISTORICO_IMPLEMENTACOES.md`
- registra O QUE efetivamente foi implementado.

---

# 3. REGRA PRINCIPAL DE DESENVOLVIMENTO

A V1 será dividida em 6 fases.

Cada fase poderá possuir:

MÍNIMO:
1 commit, quando suficiente.

PREFERENCIAL:
1 ou 2 commits.

MÁXIMO ABSOLUTO:
3 commits.

É proibido fragmentar artificialmente uma implementação apenas para
produzir mais commits.

Se uma fase aparentemente exigir mais de 3 commits, o agente deverá
interromper e avaliar:

- se o escopo cresceu;
- se houve complexidade não prevista;
- se alguma funcionalidade pode ser simplificada;
- se existe necessidade de revisão do plano.

Não ultrapassar 3 commits silenciosamente.

---

# 4. REGRA DE AVANÇO

O agente deverá executar apenas a fase autorizada.

Ao concluir uma fase:

1. executar os testes;
2. verificar os critérios de aceite;
3. atualizar `04_HISTORICO_IMPLEMENTACOES.md`;
4. apresentar resumo ao responsável;
5. informar commits realizados;
6. informar testes executados;
7. informar pendências;
8. informar o próximo passo.

Depois:

PARAR.

Não iniciar automaticamente a fase seguinte.

A próxima fase depende de autorização do responsável pelo projeto.

---

# 5. VISÃO GERAL DAS FASES

A V1 será construída em:

FASE 1 — Fundação e Banco de Auditoria

FASE 2 — Motor Excel e Comparação

FASE 3 — Auditor Local Incremental

FASE 4 — Microsoft Graph / SharePoint

FASE 5 — Interface e Relatório

FASE 6 — Robustez e Preparação para Produção

Fluxo:

F1
 ↓
F2
 ↓
F3
 ↓
F4
 ↓
F5
 ↓
F6
 ↓
V1

---

# 6. META DE COMMITS

Estimativa preferencial:

F1: 1–2 commits
F2: 1–2 commits
F3: 1–2 commits
F4: 1–3 commits
F5: 1–3 commits
F6: 1–3 commits

Estimativa total:

aproximadamente 6 a 15 commits.

Máximo teórico da V1:

18 commits.

O objetivo não é atingir o máximo.

Menos commits são preferíveis quando mantêm:

- clareza;
- testabilidade;
- rastreabilidade.

---

# 7. FASE 1 — FUNDAÇÃO E BANCO DE AUDITORIA

## 7.1 Objetivo

Criar a fundação executável da aplicação e o banco oficial da trilha.

Ao final da fase deverá existir uma aplicação Python inicial capaz de:

- iniciar;
- carregar configurações;
- configurar logging;
- criar o banco;
- criar as tabelas;
- executar testes básicos do banco.

Nenhuma integração com SharePoint deverá ser implementada nesta fase.

---

## 7.2 Entregáveis

Criar/ajustar:

`main.py`

`requirements.txt`

`.gitignore`

`.env.example`

`README.md`

Estrutura:

`app/`

`tests/`

`data/`

`logs/`

Implementar inicialmente:

`app/config.py`

`app/database.py`

`app/models.py`

`app/exceptions.py`

`app/logging_config.py`

A divisão definitiva entre `database.py` e `models.py` poderá ser
simplificada se não houver benefício em mantê-los separados.

---

## 7.3 Banco

Criar as estruturas necessárias para:

PLANILHA

CHECKPOINT

VERSAO_PROCESSADA

ALTERACAO

EXECUCAO_AUDITORIA

ERRO_PROCESSAMENTO

---

## 7.4 Requisitos do banco

Implementar:

- chaves primárias;
- relacionamentos;
- índices necessários;
- unicidade;
- proteção básica contra duplicidade;
- integridade referencial;
- timestamps relevantes.

O banco deverá ser criado automaticamente quando inexistente.

---

## 7.5 Testes da Fase 1

Validar:

- aplicação inicia;
- banco é criado;
- tabelas são criadas;
- inicialização repetida não destrói dados;
- constraints principais funcionam;
- conexão é encerrada corretamente.

---

## 7.6 Critérios de aceite

A Fase 1 será considerada concluída quando:

[ ] projeto Python estiver executável;

[ ] não existir dependência de Node.js;

[ ] banco SQLite puder ser criado;

[ ] todas as tabelas oficiais existirem;

[ ] relacionamento básico estiver funcional;

[ ] proteção contra duplicidade essencial estiver implementada;

[ ] testes da camada de persistência passarem;

[ ] documentação do histórico estiver atualizada.

---

## 7.7 Commits

Preferência:

Commit 1:
Fundação do projeto e configuração.

Commit 2:
Banco, modelo e testes.

Máximo:

3 commits.

---

# 8. FASE 2 — MOTOR EXCEL E COMPARAÇÃO

## 8.1 Objetivo

Implementar o coração da auditoria:

VERSÃO N → VERSÃO N+1

Sem SharePoint.

Sem interface gráfica.

Sem relatório final.

---

## 8.2 Entregáveis

Implementar:

`app/excel/reader.py`

`app/excel/comparator.py`

Testes correspondentes.

---

## 8.3 Reader

O leitor deverá:

- abrir `.xlsx`;
- operar somente em leitura lógica;
- preservar fórmulas;
- percorrer abas;
- gerar snapshot determinístico;
- tratar corretamente valores.

---

## 8.4 Comparator

O comparador deverá detectar:

ADD

DEL

MOD

Deverá comparar:

- abas;
- endereços;
- valores;
- fórmulas.

---

## 8.5 Cenários obrigatórios

Criar arquivos controlados de teste.

Exemplo:

`tests/fixtures/CQL028/`

Versões simuladas:

`0.84.xlsx`

`0.85.xlsx`

`0.86.xlsx`

`0.87.xlsx`

Os arquivos deverão conter alterações conhecidas.

---

## 8.6 Testes obrigatórios

Validar:

### MOD

Antes:

A1 = 10

Depois:

A1 = 15

Resultado:

MOD

---

### ADD

Antes:

B2 = vazio

Depois:

B2 = "OK"

Resultado:

ADD

---

### DEL

Antes:

C3 = "Pendente"

Depois:

C3 = vazio

Resultado:

DEL

---

### Zero

Antes:

D4 = vazio

Depois:

D4 = 0

Resultado:

ADD

Zero não poderá ser confundido com vazio.

---

### Booleano

False não poderá ser confundido com vazio.

---

### Fórmula

Antes:

=SUM(A1:A10)

Depois:

=SUM(A1:A20)

Resultado:

MOD

---

### Múltiplas abas

Alterações deverão preservar corretamente o nome da aba.

---

### Sem alterações

Snapshots iguais deverão retornar:

0 alterações.

---

## 8.7 Determinismo

Executar a mesma comparação múltiplas vezes deverá produzir o mesmo
resultado e a mesma ordenação lógica.

---

## 8.8 Critérios de aceite

[ ] Reader funcional;

[ ] fórmulas preservadas;

[ ] ADD correto;

[ ] DEL correto;

[ ] MOD correto;

[ ] zero tratado corretamente;

[ ] False tratado corretamente;

[ ] múltiplas abas funcionam;

[ ] snapshots iguais retornam zero diferenças;

[ ] resultado determinístico;

[ ] testes automatizados passam.

---

## 8.9 Commits

Preferência:

Commit 1:
Reader + snapshots.

Commit 2:
Comparator + testes.

Máximo:

3 commits.

---

# 9. FASE 3 — AUDITOR LOCAL INCREMENTAL

## 9.1 Objetivo

Construir o fluxo completo da auditoria utilizando arquivos locais.

Esta fase deverá provar o comportamento final antes de introduzir
SharePoint.

---

## 9.2 Fonte local

Implementar:

`app/sources/base.py`

`app/sources/local.py`

A fonte local deverá simular:

- planilhas;
- IDs;
- versões;
- metadados;
- arquivos históricos.

---

## 9.3 Serviço de auditoria

Implementar:

`app/audit_service.py`

Responsável por:

- selecionar planilha;
- consultar checkpoint;
- listar versões;
- determinar pendências;
- preservar versão-base;
- comparar sequencialmente;
- persistir resultados;
- atualizar checkpoint;
- registrar execução;
- registrar falhas.

---

# 10. PRIMEIRO CENÁRIO LOCAL

Exemplo:

CQL028

Versões:

0.84
0.85
0.86
0.87
0.88
0.89
0.90
0.91
0.92
0.93
0.94
0.95
0.96
0.97
0.98
0.99

Primeira execução:

0.84 → 0.85
...
0.98 → 0.99

Resultado:

checkpoint = 0.99

---

# 11. TESTE DE REEXECUÇÃO

Executar novamente sem adicionar versão.

Resultado obrigatório:

novas versões = 0

novas comparações = 0

novas alterações = 0

checkpoint continua:

0.99

Nenhum registro deverá ser duplicado.

---

# 12. TESTE INCREMENTAL

Adicionar versões simuladas:

1.00
1.01
1.02

Executar novamente.

Resultado:

0.99 → 1.00

1.00 → 1.01

1.01 → 1.02

Checkpoint final:

1.02

As comparações anteriores não deverão ser reprocessadas.

---

# 13. TESTE DE FALHA

Simular falha em:

1.01 → 1.02

Esperado:

comparações anteriores permanecem válidas;

checkpoint permanece na última comparação concluída;

execução registra falha;

próxima execução consegue retomar.

---

# 14. TRANSAÇÃO

Cada comparação deverá persistir de forma segura.

Fluxo:

BEGIN

versão processada

alterações

checkpoint

COMMIT

Em falha:

ROLLBACK

---

# 15. CRITÉRIOS DE ACEITE DA FASE 3

[ ] fonte local funciona;

[ ] primeira auditoria funciona;

[ ] histórico completo é consolidado;

[ ] checkpoint funciona;

[ ] reexecução não duplica;

[ ] auditoria incremental funciona;

[ ] versão-base é preservada;

[ ] falha não avança checkpoint incorretamente;

[ ] versão sem alterações é registrada;

[ ] histórico de execução é criado;

[ ] testes passam.

Ao concluir esta fase, o núcleo da aplicação deverá funcionar sem
SharePoint.

---

# 16. COMMITS DA FASE 3

Preferência:

Commit 1:
Fonte local + AuditService.

Commit 2:
Checkpoint, transações e testes incrementais.

Máximo:

3 commits.

---

# 17. FASE 4 — MICROSOFT GRAPH / SHAREPOINT

## 17.1 Objetivo

Substituir a fonte local pela fonte SharePoint real sem modificar o
motor de auditoria.

Esta é a principal prova da arquitetura.

---

## 17.2 Implementar

`app/sources/sharepoint.py`

e componentes auxiliares estritamente necessários.

---

# 18. AUTENTICAÇÃO

Utilizar mecanismo suportado pela Microsoft.

Requisitos:

- nenhuma senha em código;
- nenhum token no Git;
- nenhum segredo em logs;
- menor privilégio;
- operações somente leitura.

A estratégia definitiva dependerá do ambiente corporativo disponível.

---

# 19. PROVA DE LEITURA

Antes de processar auditoria real, comprovar:

[ ] acesso ao site;

[ ] acesso à biblioteca;

[ ] listagem de arquivos;

[ ] DriveItem ID;

[ ] nome;

[ ] caminho;

[ ] listagem de versões;

[ ] identificadores de versões;

[ ] data/hora;

[ ] autor;

[ ] comentário, quando disponível;

[ ] conteúdo histórico recuperável.

---

# 20. VERSÕES SECUNDÁRIAS

Este é um CRITÉRIO CRÍTICO.

No ambiente real existem versões semelhantes a:

0.84
0.85
...
0.98
0.99

A integração deverá comprovar quais dessas versões podem ser:

1. enumeradas;
2. identificadas;
3. baixadas;
4. comparadas.

Não presumir funcionamento.

---

# 21. REGRA DE BLOQUEIO

Se Microsoft Graph/SharePoint não permitir recuperar alguma categoria
necessária de versão:

PARAR.

Registrar:

- endpoint testado;
- comportamento observado;
- requisito afetado;
- impacto;
- alternativas.

Não alterar SharePoint.

Não utilizar workaround destrutivo.

Não avançar silenciosamente.

---

# 22. TESTE CONTROLADO

Utilizar inicialmente UMA planilha real.

Preferencialmente uma planilha controlada para teste.

Somente após validação poderá ser utilizada uma planilha operacional
para conferência.

---

# 23. TESTE DA CQL028

Quando autorizado e tecnicamente possível, utilizar CQL028 como cenário
de validação real.

Objetivo conceitual:

obter versões disponíveis

e produzir:

0.84 → 0.85
...
0.98 → 0.99

ou o intervalo efetivamente disponibilizado pela API.

Comparar resultado com alterações conhecidas.

---

# 24. CRITÉRIOS DE ACEITE DA FASE 4

[ ] autenticação segura;

[ ] leitura do SharePoint;

[ ] nenhuma escrita;

[ ] DriveItem ID obtido;

[ ] versões enumeradas;

[ ] conteúdo histórico recuperado;

[ ] comportamento das versões secundárias documentado;

[ ] metadados disponíveis capturados;

[ ] AuditService funciona com SharePointSource;

[ ] motor Excel não precisou ser reescrito;

[ ] teste controlado concluído.

---

# 25. COMMITS DA FASE 4

Preferência:

Commit 1:
Autenticação + descoberta/listagem.

Commit 2:
Histórico + download de versões.

Commit 3:
Integração com AuditService + testes.

Máximo:

3 commits.

---

# 26. FASE 5 — INTERFACE E RELATÓRIO

## 26.1 Objetivo

Transformar o núcleo validado em ferramenta utilizável pelo usuário.

---

# 27. INTERFACE

A interface deverá ser Python.

Preferência inicial:

Tkinter/ttk.

Não utilizar:

Node.js
npm
React
Angular
Vue
Vite

A interface deverá priorizar funcionalidade.

---

# 28. FLUXO DA INTERFACE

A tela deverá permitir:

1. listar planilhas;
2. selecionar planilha;
3. visualizar status;
4. visualizar última versão auditada;
5. visualizar última versão disponível;
6. visualizar versões pendentes;
7. iniciar auditoria;
8. acompanhar resultado;
9. gerar relatório;
10. localizar/abrir relatório gerado.

---

# 29. PRIMEIRA AUDITORIA

Se não houver checkpoint:

exibir ação:

AUDITAR HISTÓRICO

---

# 30. AUDITORIA INCREMENTAL

Se houver checkpoint:

exibir ação:

CONTINUAR AUDITORIA

Exemplo:

Última auditada: 0.99

Última disponível: 1.20

Pendentes: conforme histórico real retornado pela fonte.

---

# 31. RELATÓRIO

Implementar:

`app/report_service.py`

Gerar:

`data/reports/CQL028_Trilha_Auditoria.xlsx`

O relatório deverá ser criado a partir do banco.

---

# 32. ABAS DO RELATÓRIO

## RESUMO

- planilha;
- DriveItem ID;
- primeira versão;
- última versão;
- última execução;
- versões processadas;
- total de alterações;
- ADD;
- DEL;
- MOD.

## VERSOES

- versão anterior;
- versão atual;
- data/hora;
- autor;
- comentário;
- status;
- alterações.

## TRILHA

- ID;
- versão anterior;
- versão atual;
- data/hora;
- autor;
- comentário;
- aba;
- célula;
- tipo;
- anterior;
- novo.

---

# 33. REGENERAÇÃO

Se o relatório for excluído:

o sistema deverá conseguir gerá-lo novamente.

Se o relatório for alterado manualmente:

isso não deverá alterar o banco.

O banco permanece como fonte oficial.

---

# 34. RESPONSIVIDADE

Auditorias demoradas não deverão congelar permanentemente a interface.

Implementar mecanismo simples de processamento em background somente
quando necessário.

Não introduzir arquitetura complexa.

---

# 35. CRITÉRIOS DE ACEITE DA FASE 5

[ ] interface inicia;

[ ] planilhas são listadas;

[ ] seleção funciona;

[ ] status é exibido;

[ ] auditoria pode ser iniciada;

[ ] resultado é apresentado;

[ ] relatório é gerado;

[ ] relatório contém três abas;

[ ] filtros funcionam;

[ ] relatório pode ser regenerado;

[ ] nenhuma dependência Node.js existe.

---

# 36. COMMITS DA FASE 5

Preferência:

Commit 1:
Interface principal.

Commit 2:
Relatório.

Commit 3:
Integração e ajustes finais da experiência.

Máximo:

3 commits.

---

# 37. FASE 6 — ROBUSTEZ E PREPARAÇÃO PARA PRODUÇÃO

## 37.1 Objetivo

Preparar a aplicação validada para utilização operacional.

Não adicionar funcionalidades grandes nesta fase.

---

# 38. TESTES FINAIS

Executar:

- testes unitários;
- testes de integração;
- reexecução;
- auditoria incremental;
- falha e retomada;
- geração de relatório;
- múltiplas planilhas controladas.

---

# 39. PERFORMANCE

Medir:

- tempo por versão;
- tempo por planilha;
- quantidade de células;
- memória;
- tamanho do banco;
- tamanho dos relatórios.

Otimizar somente gargalos observados.

---

# 40. LOGS

Revisar:

- logs técnicos;
- mensagens de erro;
- ausência de segredos;
- rastreabilidade das execuções.

---

# 41. ARQUIVOS TEMPORÁRIOS

Validar:

- criação;
- utilização;
- remoção;
- recuperação após falha.

Nenhum arquivo temporário deverá ser confundido com evidência oficial.

---

# 42. INTEGRIDADE

Avaliar/implementar SHA-256 das versões processadas quando tecnicamente
adequado.

Registrar claramente sua finalidade.

---

# 43. BACKUP

Documentar estratégia mínima de backup do banco.

O banco de auditoria deverá ser considerado ativo crítico.

---

# 44. EMPACOTAMENTO

Avaliar distribuição simples para Windows.

Se necessário, poderá ser utilizado mecanismo Python apropriado para
geração de executável.

A solução final não poderá exigir Node.js.

Evitar exigir que o usuário operacional execute comandos técnicos se
for viável disponibilizar executável.

---

# 45. CRITÉRIOS DE ACEITE DA FASE 6

[ ] testes críticos passam;

[ ] auditoria incremental está estável;

[ ] falhas não corrompem checkpoint;

[ ] logs adequados;

[ ] temporários controlados;

[ ] relatório confiável;

[ ] desempenho aceitável no cenário testado;

[ ] documentação atualizada;

[ ] instruções de execução disponíveis;

[ ] nenhuma operação de escrita no SharePoint;

[ ] V1 pronta para homologação.

---

# 46. COMMITS DA FASE 6

Preferência:

Commit 1:
Robustez + testes.

Commit 2:
Performance + integridade.

Commit 3:
Empacotamento + documentação de operação.

Máximo:

3 commits.

---

# 47. O QUE NÃO DEVE ACONTECER

Durante qualquer fase, evitar:

- implementar funcionalidades de fases futuras sem necessidade;
- criar dezenas de arquivos sem justificativa;
- adicionar frameworks pesados;
- criar microsserviços;
- adicionar Node.js;
- criar frontend web complexo;
- alterar SharePoint;
- utilizar nome do arquivo como identidade;
- avançar checkpoint antes do commit dos dados;
- reprocessar histórico sem necessidade;
- utilizar relatório Excel como banco oficial;
- esconder limitações técnicas.

---

# 48. CONTROLE DE ALTERAÇÕES DE ESCOPO

Se surgir novo requisito:

1. identificar;
2. avaliar impacto;
3. classificar como:
   - necessário para V1;
   - melhoria futura;
4. atualizar documentação se necessário;
5. obter autorização;
6. implementar na fase apropriada.

Não aumentar silenciosamente o escopo.

---

# 49. MELHORIAS FUTURAS

Após conclusão da V1 poderão ser avaliadas, separadamente:

- SQL Server;
- processamento em lote;
- auditoria automática agendada;
- Power BI;
- dashboards;
- filtros avançados;
- exportação adicional;
- política avançada de retenção;
- múltiplos ambientes SharePoint;
- notificações;
- administração centralizada.

Essas funcionalidades NÃO fazem parte automaticamente da V1.

---

# 50. DEFINIÇÃO DE PRONTO DA V1

A V1 estará pronta quando o seguinte cenário funcionar:

1. usuário inicia a ferramenta;

2. ferramenta acessa SharePoint somente em leitura;

3. usuário seleciona CQL028.xlsx;

4. ferramenta identifica a planilha pelo DriveItem ID;

5. consulta versões;

6. identifica ausência ou existência de checkpoint;

7. na primeira execução processa histórico disponível;

8. compara N → N+1;

9. identifica ADD, DEL e MOD;

10. armazena trilha no banco;

11. registra versões processadas;

12. registra execução;

13. estabelece checkpoint;

14. usuário encerra a ferramenta;

15. novas versões surgem no SharePoint;

16. usuário abre novamente a ferramenta;

17. seleciona a mesma planilha;

18. sistema reconhece o histórico existente;

19. identifica novas versões;

20. continua exatamente do checkpoint;

21. não duplica registros anteriores;

22. atualiza o banco consolidado;

23. gera `CQL028_Trilha_Auditoria.xlsx`;

24. relatório apresenta histórico completo;

25. nenhuma versão do SharePoint foi modificada.

Esse é o principal critério funcional de sucesso do projeto.

---

# 51. PROCEDIMENTO DO CODEX EM CADA SESSÃO

Ao iniciar uma sessão de implementação, o Codex deverá:

1. ler os quatro documentos oficiais;

2. identificar a fase atual;

3. consultar `04_HISTORICO_IMPLEMENTACOES.md`;

4. verificar o estado real do código;

5. confirmar a primeira tarefa pendente;

6. implementar somente o necessário para a fase autorizada;

7. executar testes;

8. corrigir falhas relacionadas ao trabalho realizado;

9. atualizar `04_HISTORICO_IMPLEMENTACOES.md`;

10. criar commit(s) coerentes;

11. apresentar relatório final da sessão.

---

# 52. RELATÓRIO DO CODEX

Ao terminar uma fase, informar:

FASE:
[identificação]

STATUS:
CONCLUÍDA / BLOQUEADA

IMPLEMENTADO:
[resumo]

ARQUIVOS CRIADOS:
[...]

ARQUIVOS ALTERADOS:
[...]

TESTES:
[...]

RESULTADOS:
[...]

COMMITS:
[...]

PENDÊNCIAS:
[...]

RISCOS:
[...]

PRÓXIMO PASSO:
[...]

Depois:

PARAR E AGUARDAR AUTORIZAÇÃO.

---

# 53. BLOQUEIO TÉCNICO

Se ocorrer bloqueio técnico, não improvisar alteração arquitetural
significativa.

Registrar:

BLOQUEIO

Fase:
[...]

Problema:
[...]

Causa:
[...]

Impacto:
[...]

Alternativas:
1. [...]
2. [...]

Recomendação:
[...]

Aguardar decisão quando a escolha alterar requisito, segurança ou
arquitetura.

---

# 54. STATUS INICIAL DO PROJETO

No momento da criação deste documento:

FASE 1 — NÃO INICIADA
FASE 2 — NÃO INICIADA
FASE 3 — NÃO INICIADA
FASE 4 — NÃO INICIADA
FASE 5 — NÃO INICIADA
FASE 6 — NÃO INICIADA

Progresso da V1:

0%

---

# 55. PRÓXIMA AÇÃO OFICIAL

Após criação dos quatro documentos oficiais:

iniciar:

FASE 1 — FUNDAÇÃO E BANCO DE AUDITORIA

Não iniciar Fase 2 simultaneamente.

A Fase 1 deverá ser concluída, testada e registrada antes da autorização
para a Fase 2.

---

FIM DO DOCUMENTO