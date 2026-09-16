# PLANO DE DESENVOLVIMENTO
## Auditor de Planilhas Excel — SharePoint Online

**Documento:** 03_PLANO_DE_DESENVOLVIMENTO.md  
**Versão:** 1.1
**Status:** Oficial
**Data:** 15/09/2026
**Última revisão:** 15/09/2026

---

# ENCERRAMENTO DA F5 E INÍCIO AUTORIZADO DA F6 — 16/09/2026

A validação real da F5 foi aceita no Windows corporativo com SharePoint Online,
Selenium/Edge e o SQLite existente migrado sem recriação. A F5 está concluída. Foi
autorizado iniciar a F6, executando estritamente uma tarefa por vez na ordem deste
plano. A primeira tarefa é **testes unitários**, primeiro item de Testes Finais.

O cenário de capacidade a validar na F6 passa a considerar aproximadamente 2.000
planilhas, algumas com mais de 3.000 versões, e potencialmente milhões de alterações,
sempre em um banco canônico compartilhado. Isso não autoriza trocar o SQLite sem
evidência de testes. Site e escopos SharePoint permanecem configurações variáveis e não
devem ser fixados no código.

---

# INÍCIO AUTORIZADO DA F5 — 16/09/2026

A solicitação expressa do responsável para identificar e executar a próxima fase
autorizou o início da F5 após a conclusão da F4. Interface Tkinter e relatório
consolidado foram implementados no primeiro commit da fase. A F5 permanece em
andamento até a validação manual da interface em ambiente gráfico; F6 não está
autorizada.

# ENCERRAMENTO DO PLANO — F4 CONCLUÍDA

A evidência corporativa fornecida encerrou as duas validações que ainda estavam pendentes:
a versão atual foi identificada por `UIVersion/UIVersionLabel` e adquirida por `/$value`;
e o `UniqueId` foi preservado após renomeação e movimentação manuais de arquivo
descartável. A implementação passou a compor histórico + atual, validar downloads reais
do Edge e usar `(site, contexto REST, UniqueId)` como identidade.

O Git real desta branch contém três commits F4 anteriores alcançáveis: `1a7421a`,
`ee013b3` e `ae6b9a2`. `df7943f`, citado por registro antigo de PR fechado sem merge, não
é objeto presente. Este encerramento usa um dos dois commits adicionais já autorizados
pela exceção controlada da F4. Todos os critérios da seção 25 foram atendidos por código,
testes automatizados e evidência corporativa. F5 continua não autorizada/não iniciada.

# ATUALIZAÇÃO INTERMEDIÁRIA DO PLANO — REGISTRO HISTÓRICO

O histórico Git real da branch deve prevalecer sobre hashes meramente citados em
documentação. Nesta branch, `1a7421a` contém a implementação Graph e o registro do
bloqueio; `ee013b3` contém a revisão documental. O hash `df7943f` citado anteriormente
não existe no repositório e não deve ser contado. Assim, esta entrega usa um único
commit adicional da F4, dentro do máximo ordinário de três commits efetivos, sem
invocar a exceção documental baseada na premissa incorreta de três commits anteriores.

A tarefa autorizada é implementar e testar `BrowserSharePointSource`: descoberta
recursiva, identidade `UniqueId`, versões históricas e metadados, download pelo ID,
validação XLSX, integração incremental e controles GET-only. O teste corporativo já
comprovou `/_api/web` (sem o espaço), `/Versions`, `CreatedBy`, `/Versions(ID)/$value`
e aquisição de 0.1–0.98. A versão atual e a estabilidade operacional do `UniqueId`
após cenários de movimentação/renomeação ainda requerem validação controlada. F4
permanece em andamento por essas pendências. F5 não está autorizada.

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

Antes de implementar uma fase, consultar na ordem definida pela governança:

1. `docs/06_GOVERNANCA.md`
2. `docs/01_ESPECIFICACAO_FUNCIONAL.md`
3. `docs/02_ARQUITETURA.md`
4. `docs/03_PLANO_DE_DESENVOLVIMENTO.md`
5. `docs/04_HISTORICO_IMPLEMENTACOES.md`

O arquivo:

`docs/05_PROMPT_OFICIAL.md`

define as instruções operacionais utilizadas nas sessões com o Codex.

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

FASE 4 — Aquisição de Versões SharePoint

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

EXCEÇÃO CONTROLADA DA FASE 4

A Fase 4 atingiu originalmente o limite de 3 commits durante a tentativa
de integração com Microsoft Graph e foi corretamente interrompida após
a identificação de bloqueio corporativo.

A revisão arquitetural decorrente desse bloqueio constitui alteração
formal de requisito autorizada pelo responsável pelo projeto.

Excepcionalmente, poderão ser autorizados até 2 commits adicionais
exclusivamente para:

1. adequação da Fase 4 ao mecanismo de aquisição aprovado;
2. implementação e validação da solução de aquisição escolhida.

Essa exceção não autoriza avanço para a Fase 5.

Qualquer necessidade além desses commits deverá provocar nova
interrupção e avaliação.

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

Gerar arquivos controlados de teste programaticamente com `openpyxl`, usando
`pytest/tmp_path` ou mecanismo temporário equivalente. Os arquivos devem ser
descartados após os testes e não devem ser versionados como fixtures binárias.

Versões simuladas durante a execução:

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

# 17. FASE 4 — AQUISIÇÃO DE VERSÕES SHAREPOINT

## 17.1 Objetivo

Integrar o núcleo de auditoria já validado a uma fonte SharePoint real
sem modificar as regras do motor de auditoria.

A Fase 4 deverá identificar, validar e implementar um mecanismo de
aquisição de versões históricas que seja:

- suportado tecnicamente;
- autorizado no ambiente corporativo;
- exclusivamente de leitura;
- compatível com versões principais e secundárias necessárias;
- capaz de fornecer conteúdo suficiente para a trilha de auditoria;
- compatível com processamento incremental;
- adequado ao uso futuro em escala.

Microsoft Graph permanece como mecanismo possível quando disponível e
autorizado, mas não constitui dependência obrigatória da V1.

---

## 17.2 Situação identificada

Durante a execução original da Fase 4, a integração Microsoft Graph
não pôde ser concluída porque o ambiente corporativo não disponibiliza
ao projeto os dados e autorizações necessários para a autenticação
originalmente planejada.

A organização não disponibilizou ao projeto os identificadores,
credenciais e registro de aplicação necessários para esse modelo de
integração.

Essa restrição deverá ser tratada como limitação real do ambiente e não
deverá ser contornada.

---

## 17.3 Evidência técnica já obtida

Foi validado com a planilha real CQL028.xlsx que uma versão histórica
pode ser acessada pelo usuário autenticado através de endereço
SharePoint contendo `_vti_history`.

No teste realizado, o recurso correspondente à versão 0.97 abriu
corretamente a versão histórica esperada através da sessão autenticada
do usuário.

Foi realizado também teste HTTP controlado através de Python sem
credenciais, cookies ou tokens.

Resultado:

HTTP 403

Portanto, o acesso pelo navegador/Excel autenticado não comprova
capacidade de acesso programático através do mesmo endereço.

---

# 18. INVESTIGAÇÃO DO MECANISMO DE AQUISIÇÃO

Antes de implementar uma nova integração, deverão ser avaliados
mecanismos suportados e autorizados que possam fornecer acesso
programático às versões históricas utilizando as capacidades
disponíveis no ambiente corporativo.

Cada alternativa deverá ser avaliada quanto a:

- autenticação;
- autorização;
- operação somente leitura;
- listagem/descoberta das planilhas;
- identidade técnica;
- descoberta das versões;
- recuperação do conteúdo histórico;
- versões principais;
- versões secundárias;
- metadados;
- ordenação;
- escalabilidade;
- segurança;
- compatibilidade com o ambiente Windows corporativo.

A investigação deverá ser objetiva e limitada ao necessário para
selecionar uma solução viável.

---

# 19. REGRAS DE SEGURANÇA DA INVESTIGAÇÃO

É proibido:

- extrair cookies do navegador;
- capturar tokens de sessões existentes;
- copiar credenciais internas do Microsoft Office;
- armazenar senha corporativa;
- utilizar autenticação obsoleta ou insegura;
- contornar políticas do Microsoft Entra;
- elevar permissões;
- modificar o SharePoint;
- restaurar versões;
- excluir versões;
- criar versões para facilitar a aquisição.

A existência de acesso através do navegador ou Microsoft Excel não
autoriza automaticamente sua reutilização programática.

---

# 20. CONTRATO DA FONTE

O mecanismo escolhido deverá permanecer atrás da abstração de fonte
definida na arquitetura.

O AuditService não deverá conhecer detalhes específicos de:

- Microsoft Graph;
- autenticação;
- `_vti_history`;
- URLs SharePoint;
- mecanismo alternativo de aquisição.

A fonte deverá entregar ao núcleo informações normalizadas equivalentes
às utilizadas pela LocalSource.

A substituição da fonte não deverá exigir reescrita do ExcelReader,
Comparator, regras de checkpoint ou persistência.

---

# 21. PROVA DE LEITURA

Para o mecanismo candidato deverão ser comprovados, quando tecnicamente
disponíveis:

[ ] acesso autorizado ao SharePoint;

[ ] identificação da planilha;

[ ] identidade técnica estável;

[ ] nome e caminho;

[ ] descoberta/listagem das versões;

[ ] identificadores das versões;

[ ] ordenação confiável;

[ ] data/hora;

[ ] autor da versão;

[ ] comentário, quando disponível;

[ ] conteúdo histórico recuperável;

[ ] operação exclusivamente em leitura.

A indisponibilidade de determinado metadado deverá ser registrada
explicitamente e não preenchida através de inferência.

---

# 22. VERSÕES SECUNDÁRIAS

Este permanece como CRITÉRIO CRÍTICO.

No ambiente real existem versões como:

0.84
0.85
...
0.98
0.99

A solução deverá comprovar quais dessas versões podem ser:

1. descobertas;
2. identificadas;
3. adquiridas;
4. ordenadas;
5. comparadas.

Não presumir funcionamento com base apenas na interface do SharePoint.

---

# 23. CONTINGÊNCIA

Caso nenhuma solução automatizada suportada e autorizada seja encontrada,
a Fase 4 deverá PARAR novamente.

Nesse caso deverão ser documentados:

- mecanismos avaliados;
- resultados;
- limitações;
- impacto operacional;
- alternativas restantes.

Somente após decisão explícita do responsável pelo projeto poderá ser
adotado mecanismo de importação assistida ou manual.

A importação manual de centenas de versões NÃO constitui, neste momento,
a solução oficial da V1.

---

# 24. TESTE CONTROLADO

A validação deverá utilizar inicialmente UMA planilha real.

Quando autorizado e tecnicamente possível:

CQL028.xlsx

A versão 0.97, cujo acesso histórico já foi comprovado pelo usuário,
poderá ser utilizada como uma das referências do teste.

O objetivo final permanece obter versões consecutivas suficientes para
comprovar:

N → N+1

e validar o fluxo completo através da SharePointSource.

Nenhuma operação de escrita poderá ocorrer.

---

# 25. CRITÉRIOS DE ACEITE DA FASE 4

A Fase 4 será considerada concluída quando:

[ ] existir mecanismo de aquisição definido e documentado;

[ ] o mecanismo for suportado e autorizado no ambiente;

[ ] acesso ao SharePoint ocorrer exclusivamente em leitura;

[ ] identidade técnica estável da planilha estiver definida;

[ ] versões necessárias puderem ser descobertas/adquiridas;

[ ] ordenação das versões estiver validada;

[ ] comportamento das versões secundárias estiver documentado;

[ ] conteúdo histórico necessário puder ser recuperado;

[ ] metadados disponíveis forem capturados;

[ ] AuditService funcionar com SharePointSource;

[ ] motor Excel não precisar ser reescrito;

[ ] checkpoint e idempotência permanecerem funcionais;

[ ] teste controlado com planilha real for concluído;

[ ] nenhuma operação de escrita no SharePoint ocorrer.

Caso esses critérios não possam ser atendidos por restrição corporativa,
a fase deverá permanecer BLOQUEADA até decisão formal sobre contingência.

---

# 25.1 COMMITS DA FASE 4

A Fase 4 já consumiu os 3 commits originalmente previstos antes da
identificação e registro definitivo do bloqueio corporativo.

Mediante autorização expressa do responsável pelo projeto, ficam
permitidos excepcionalmente até 2 commits adicionais exclusivamente
para concluir a Fase 4 revisada.

Não iniciar Fase 5 dentro desses commits.

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
- identidade técnica da planilha;
- DriveItem ID, quando disponível;
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

Executar, nesta ordem e como tarefas independentes:

- [x] testes unitários;
- [x] testes de integração;
- [x] reexecução;
- [x] auditoria incremental;
- [ ] falha e retomada;
- [ ] geração de relatório;
- [ ] múltiplas planilhas controladas.

Os quatro primeiros itens foram executados, cada um após autorização própria. Os itens
seguintes continuam pendentes e não estão implicitamente autorizados pela conclusão do
anterior.

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
- utilizar exclusivamente o nome do arquivo como identidade técnica;
- contornar mecanismos corporativos de autenticação ou autorização;
- assumir que acesso pelo navegador implica acesso programático;
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

2. ferramenta utiliza mecanismo SharePoint suportado e autorizado,
   exclusivamente em leitura;

3. usuário seleciona CQL028.xlsx;

4. ferramenta reconhece a planilha através de identidade técnica estável;

5. consulta ou adquire automaticamente as versões históricas necessárias;

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

1. ler a documentação obrigatória conforme ordem definida em
   `docs/06_GOVERNANCA.md`;

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

# 54. STATUS ATUAL DO PROJETO

FASE 1 — CONCLUÍDA
FASE 2 — CONCLUÍDA
FASE 3 — CONCLUÍDA
FASE 4 — CONCLUÍDA
FASE 5 — CONCLUÍDA
FASE 6 — EM ANDAMENTO (QUATRO TAREFAS CONCLUÍDAS)

Situação atual:

A F5 foi aceita em ambiente corporativo real. A F6 foi iniciada por autorização
expressa e concluiu as quatro primeiras tarefas do backlog: testes unitários, testes de
integração automatizados, reexecução e auditoria incremental.

Os requisitos de capacidade e operação com várias planilhas em um banco canônico estão
registrados para validação posterior na ordem do backlog da F6.

Progresso da V1:

5 de 6 fases concluídas; F6 em andamento.

---

# 55. PRÓXIMA AÇÃO OFICIAL

Próxima atividade: falha e retomada da F6. Não executar sem nova autorização.
O limite ordinário de três commits da F6 foi atingido; qualquer novo commit exige
autorização expressa ou procedimento previsto pela governança.

Não implementar importação manual como solução oficial sem nova decisão.

Não executar a próxima tarefa da F6 sem autorização expressa.

Não iniciar a F6 sem autorização expressa.

FIM DO DOCUMENTO
