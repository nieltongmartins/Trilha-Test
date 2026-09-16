# GOVERNANÇA DO PROJETO
## Auditor de Planilhas Excel — SharePoint Online

**Documento:** 06_GOVERNANCA.md
**Versão:** 1.1
**Status:** Oficial
**Data:** 15/09/2026
**Última revisão:** 16/09/2026

---

# DECISÃO DE ENCERRAMENTO DA V1 — 16/09/2026

O responsável pelo projeto concedeu autorização superior e explícita para identificar,
executar, validar e concluir todas as tarefas oficiais restantes nas fases abertas, sem
nova autorização entre tarefas. A decisão também autorizou os commits tecnicamente
necessários, sem reescrita de histórico ou fragmentação artificial. Esta decisão atual
tem precedência sobre as antigas regras de autorização tarefa a tarefa.

O inventário confirmou F1–F5 concluídas, backup da F6 documentalmente concluído e
empacotamento como única implementação pendente. Após empacotamento, 52 testes
automatizados aprovados, validação integrada, PRAGMAs e reconciliação dos seis documentos,
a F6 e a V1 foram encerradas. Nenhuma invariante de SharePoint read-only, SQLite
canônico, identidade, adjacência, transação, checkpoint ou segurança foi alterada.

**Estado oficial:** seis de seis fases concluídas; nenhuma tarefa oficial pendente;
ELABORAÇÃO DA VERSÃO ATUAL CONCLUÍDA PARA UTILIZAÇÃO E TESTES REAIS.

# DECISÃO FINAL DE GOVERNANÇA DA F4

As evidências corporativas de versão atual e estabilidade do `UniqueId` removem as pendências anteriores. O provider oficial V1 é Selenium/Edge + SharePoint REST read-only, com autenticação interativa e sem captura/transferência da sessão. A identidade é `(site, contexto REST, UniqueId)`; nome e caminho são mutáveis.

O Git real contém `1a7421a`, `ee013b3` e `ae6b9a2` como trabalhos F4 alcançáveis; `df7943f` não existe no repositório e corresponde a registro antigo de tentativa não incorporada. Como o limite ordinário de três foi alcançado, aplica-se a exceção já formalizada de até dois commits adicionais, usando somente um commit coerente para este encerramento. Isso não autoriza nem inicia F5.


# DECISÃO DE GOVERNANÇA APLICÁVEL À CONTINUAÇÃO DA F4

O responsável autorizou o provider Selenium/Edge + SharePoint REST como mecanismo da
V1. Isso não autoriza reutilização externa da sessão: toda comunicação autenticada deve
permanecer no navegador, somente com GET same-origin sob `/_api/`; captura de cookies,
tokens, senha, PRT ou credenciais permanece proibida. Descoberta deve ser dinâmica e
recursiva; nome/caminho não substituem identidade técnica; ID de versão e VersionLabel
devem permanecer distintos; falha intermediária interrompe a cadeia e o checkpoint só
avança após persistência. Graph pode permanecer isolado e opcional.

Para contagem de commits, somente objetos efetivamente alcançáveis na branch contam.
Hashes apenas documentados ou pertencentes a PR fechado sem merge não consomem o
limite. O registro intermediário identificava `1a7421a` e `ee013b3`; a inspeção final também
confirmou `ae6b9a2` alcançável. `df7943f` não existe. Aplica-se a exceção controlada
descrita na decisão final acima, sem fragmentação. Esta decisão não autoriza F5.

---

# 1. OBJETIVO

Este documento estabelece as regras de governança do projeto
Auditor de Planilhas Excel — SharePoint Online.

Sua finalidade é garantir que o desenvolvimento permaneça:

- alinhado aos requisitos;
- simples;
- controlado;
- seguro;
- rastreável;
- incremental;
- tecnicamente consistente.

Este documento não substitui a especificação funcional, arquitetura,
plano de desenvolvimento ou histórico de implementações.

Sua função é definir:

- autoridade;
- hierarquia documental;
- regras obrigatórias;
- limites de autonomia do agente de desenvolvimento;
- tratamento de conflitos;
- tratamento de bloqueios;
- controle de escopo;
- critérios para alteração das decisões oficiais.

---

# 2. DOCUMENTAÇÃO OFICIAL

A documentação oficial do projeto é composta por:

1. `01_ESPECIFICACAO_FUNCIONAL.md`
2. `02_ARQUITETURA.md`
3. `03_PLANO_DE_DESENVOLVIMENTO.md`
4. `04_HISTORICO_IMPLEMENTACOES.md`
5. `05_PROMPT_OFICIAL.md`
6. `06_GOVERNANCA.md`

Não criar novos documentos de governança, planejamento ou controle
sem necessidade técnica clara ou autorização do responsável pelo
projeto.

---

# 3. FUNÇÃO DE CADA DOCUMENTO

## 01_ESPECIFICACAO_FUNCIONAL.md

Define:

O QUE o sistema deve fazer.

Contém:

- requisitos;
- regras de negócio;
- comportamentos esperados;
- limitações;
- critérios funcionais de aceite.

---

## 02_ARQUITETURA.md

Define:

COMO o sistema deverá ser estruturado.

Contém:

- componentes;
- responsabilidades;
- fluxo de dados;
- persistência;
- integração;
- decisões arquiteturais.

---

## 03_PLANO_DE_DESENVOLVIMENTO.md

Define:

EM QUE ORDEM o sistema será implementado.

Contém:

- fases;
- objetivos;
- entregáveis;
- testes;
- critérios de aceite;
- limites de commits.

---

## 04_HISTORICO_IMPLEMENTACOES.md

Define:

O QUE efetivamente aconteceu durante o desenvolvimento.

Contém:

- implementações realizadas;
- testes executados;
- commits;
- decisões;
- bloqueios;
- limitações confirmadas;
- estado atual do projeto.

Este documento não deve registrar como concluído algo que ainda não
existe no código.

---

## 05_PROMPT_OFICIAL.md

Define:

COMO o agente de desenvolvimento deve iniciar e conduzir suas sessões.

O prompt não deverá criar requisitos conflitantes com a documentação
oficial.

Sua função principal é direcionar o agente para consultar e respeitar
os documentos oficiais.

---

## 06_GOVERNANCA.md

Define:

QUAIS REGRAS CONTROLAM O DESENVOLVIMENTO.

Este documento estabelece:

- hierarquia;
- limites;
- autoridade;
- resolução de conflitos;
- controle de mudanças.

---

# 4. AUTORIDADE DO PROJETO

A autoridade máxima sobre:

- requisitos;
- escopo;
- prioridades;
- arquitetura;
- mudanças relevantes;
- inclusão ou exclusão de funcionalidades;

pertence ao responsável pelo projeto.

O agente de desenvolvimento poderá:

- analisar;
- recomendar;
- apontar riscos;
- sugerir alternativas;
- implementar decisões autorizadas.

O agente não poderá assumir autoridade para alterar unilateralmente
requisitos ou escopo.

---

# 5. HIERARQUIA DE DECISÃO

Em caso de conflito, utilizar a seguinte ordem:

1. decisão expressa e atual do responsável pelo projeto;

2. `06_GOVERNANCA.md`;

3. `01_ESPECIFICACAO_FUNCIONAL.md`;

4. `02_ARQUITETURA.md`;

5. `03_PLANO_DE_DESENVOLVIMENTO.md`;

6. `04_HISTORICO_IMPLEMENTACOES.md`;

7. `05_PROMPT_OFICIAL.md`;

8. implementação existente.

O código existente não possui autoridade para invalidar silenciosamente
uma regra documental.

Se código e documentação divergirem, a divergência deverá ser analisada.

---

# 6. CONFLITO ENTRE DOCUMENTOS

Quando dois documentos apresentarem instruções incompatíveis:

NÃO escolher silenciosamente uma interpretação.

O agente deverá:

1. identificar o conflito;
2. citar os documentos envolvidos;
3. explicar o impacto;
4. verificar a hierarquia definida neste documento;
5. aplicar a regra superior quando a resolução for inequívoca;
6. solicitar decisão quando a resolução alterar requisito, segurança
   ou comportamento importante.

---

# 7. PRINCÍPIOS INEGOCIÁVEIS

As seguintes regras são consideradas fundamentais para a V1:

## 7.1 SharePoint somente leitura

A aplicação não poderá modificar o SharePoint.

São proibidas operações como:

- alterar arquivo;
- sobrescrever arquivo;
- excluir arquivo;
- excluir versão;
- criar versão;
- alterar metadados;
- alterar comentários;
- check-in;
- check-out;
- qualquer operação de escrita não explicitamente autorizada em uma
  futura mudança formal de requisito.

Conveniência técnica não justifica violar esta regra.

---

## 7.2 Ausência de Node.js

A aplicação não deverá possuir dependência obrigatória de:

- Node.js;
- npm;
- React;
- Angular;
- Vue;
- Vite;
- ferramentas equivalentes que introduzam dependência do ecossistema
  Node para execução da aplicação.

A solução deverá permanecer baseada em Python.

---

## 7.3 Banco como fonte oficial

O banco de auditoria será a fonte oficial da trilha consolidada.

Arquivos Excel de relatório são produtos derivados.

O relatório não deverá ser utilizado como substituto do banco para:

- checkpoint;
- controle de versões;
- identificação de registros processados;
- idempotência.

---

## 7.4 Identidade da planilha

O nome do arquivo não será utilizado como identidade técnica exclusiva
da planilha.

A aplicação deverá utilizar identidade técnica estável e inequívoca
fornecida ou estabelecida através da fonte SharePoint adotada.

Quando disponíveis, deverão ser preservados os identificadores
fornecidos pelo SharePoint/Microsoft Graph, incluindo DriveItem ID e os
identificadores de contexto necessários.

Caso o mecanismo autorizado de aquisição não disponibilize DriveItem ID,
deverá ser definida e documentada identidade técnica alternativa antes
do uso operacional.

Nenhuma alteração dessa estratégia poderá comprometer checkpoint,
idempotência ou continuidade da trilha existente.

---

## 7.5 Auditoria incremental

Auditorias posteriores deverão continuar do checkpoint.

Histórico já consolidado não deverá ser recalculado sem motivo
explicitamente definido.

---

## 7.6 Idempotência

Reexecutar uma auditoria sem novas versões não poderá duplicar:

- comparações;
- versões processadas;
- alterações;
- evidências.

---

## 7.7 Integridade do checkpoint

O checkpoint somente poderá avançar depois que os dados correspondentes
forem persistidos com sucesso.

Falha de processamento não poderá produzir checkpoint falso.

---

# 8. REGRA DE FASES

A V1 possui seis fases oficiais:

F1 — Fundação e Banco de Auditoria

F2 — Motor Excel e Comparação

F3 — Auditor Local Incremental

F4 — Aquisição de Versões SharePoint

F5 — Interface e Relatório

F6 — Robustez e Preparação para Produção

O agente deverá trabalhar somente na fase autorizada.

O mecanismo concreto utilizado na F4 não constitui uma fase independente.

Microsoft Graph ou outro mecanismo suportado e autorizado deverá
permanecer encapsulado na camada de aquisição SharePoint.

---

# 9. LIMITE DE COMMITS

Cada fase poderá possuir no máximo:

3 commits.

Preferência:

1 ou 2 commits.

O limite existe para manter o desenvolvimento:

- objetivo;
- curto;
- compreensível;
- controlado.

Não criar microcommits artificialmente.

Se uma fase aparentar exigir mais de 3 commits, o agente deverá parar e
avaliar o motivo antes de continuar.

Exceções ao limite somente poderão ocorrer mediante decisão expressa do
responsável pelo projeto, com:

- motivo;
- fase afetada;
- quantidade adicional autorizada;
- finalidade;
- registro no plano e/ou histórico correspondente.

Uma exceção não altera permanentemente o limite padrão das demais fases.

---

# 10. PROIBIÇÃO DE AVANÇO AUTOMÁTICO

Concluir uma fase NÃO autoriza iniciar a próxima.

Ao concluir uma fase, o agente deverá:

1. executar testes;
2. atualizar o histórico;
3. registrar commits;
4. apresentar resultado;
5. informar pendências;
6. indicar a próxima fase;
7. PARAR.

A próxima fase somente será iniciada após autorização do responsável.

---

# 11. CONTROLE DE ESCOPO

O agente não deverá implementar automaticamente funcionalidades que:

- não estejam previstas;
- sejam apenas sugestões;
- sejam melhorias futuras;
- pertençam a fases posteriores;
- aumentem significativamente a complexidade.

Quando identificar uma possível melhoria:

registrar ou apresentar como:

MELHORIA FUTURA

e não implementar sem autorização.

---

# 12. SIMPLICIDADE

Este projeto deverá permanecer simples.

Evitar:

- overengineering;
- microsserviços;
- filas distribuídas sem necessidade;
- infraestrutura complexa;
- frameworks pesados;
- abstrações prematuras;
- camadas sem finalidade prática;
- padrões arquiteturais aplicados apenas por convenção.

Uma solução menor e correta é preferível a uma solução sofisticada sem
benefício demonstrável.

---

# 13. DEPENDÊNCIAS

Antes de adicionar nova dependência, verificar:

1. é realmente necessária?
2. Python padrão já resolve?
3. é mantida?
4. introduz risco?
5. dificulta distribuição?
6. cria dependência externa desnecessária?

O `requirements.txt` deverá permanecer pequeno.

---

# 14. IMPLEMENTAÇÃO ANTECIPADA

É proibido implementar uma fase futura apenas porque determinado arquivo
já está sendo alterado.

Exemplo:

Durante F2, não implementar interface da F5 apenas por conveniência.

Exceções somente quando uma pequena preparação técnica for
indispensável para evitar retrabalho evidente e não introduzir
funcionalidade futura.

---

# 15. TESTES

Nenhuma fase deverá ser marcada como concluída sem os testes previstos
para seus critérios de aceite.

O agente não poderá registrar:

"testado"

quando o teste não tiver sido executado.

O histórico deverá informar comandos e resultados reais sempre que
aplicável.

---

# 16. FALHA EM TESTES

Quando um teste obrigatório falhar:

a fase permanece:

EM ANDAMENTO

ou:

BLOQUEADA.

Não marcar a fase como concluída apenas porque a maior parte da
implementação funciona.

---

# 17. BLOQUEIOS

Um bloqueio relevante deverá ser comunicado.

Formato mínimo:

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

O agente não deverá esconder bloqueios através de workaround que altere
requisitos.

---

# 18. AQUISIÇÃO SHAREPOINT

Comportamentos do SharePoint, Microsoft Graph ou de qualquer outro
mecanismo de aquisição não deverão ser inventados.

Durante a F4 deverão ser validados no ambiente real, conforme as
capacidades do mecanismo candidato:

- autenticação;
- autorização;
- acesso ao SharePoint;
- identidade técnica da planilha;
- site/biblioteca, quando aplicável;
- histórico;
- versões;
- ordenação;
- metadados;
- conteúdo histórico;
- versões principais;
- versões secundárias;
- operação exclusivamente em leitura.

Microsoft Graph poderá ser utilizado quando disponível e autorizado,
mas não constitui mecanismo obrigatório ou exclusivo da V1.

Outro mecanismo somente poderá ser adotado quando for suportado,
autorizado e compatível com os princípios de segurança deste projeto.

---

# 19. VERSÕES SECUNDÁRIAS

Versões observadas na interface do SharePoint, como:

0.84
0.85
0.86
...
0.99

somente poderão ser consideradas tecnicamente suportadas pela aplicação
depois que a integração comprovar que podem ser recuperadas de forma
adequada.

A presença na interface web não autoriza o agente a presumir que a
versão possa ser adquirida programaticamente pelo mecanismo escolhido.

O comportamento deverá ser comprovado tecnicamente.

---

# 20. LIMITAÇÃO DE INTEGRAÇÃO

Caso uma limitação do SharePoint, Microsoft Graph, autenticação
corporativa ou mecanismo de aquisição impeça requisito previsto:

não alterar silenciosamente o requisito.

O agente deverá:

1. documentar comportamento observado;
2. registrar endpoint/método relevante;
3. explicar impacto;
4. propor alternativas;
5. aguardar decisão quando necessário.

---

# 21. SEGURANÇA

Nunca incluir no repositório:

- senha;
- token;
- client secret real;
- chave privada;
- credencial corporativa;
- informação sensível desnecessária.

Arquivos de exemplo poderão conter somente placeholders.

Também é proibido utilizar como mecanismo de contorno:

- extração de cookies de navegador;
- captura de tokens de sessões autenticadas;
- extração de credenciais do Microsoft Office;
- armazenamento de senha corporativa;
- reutilização não autorizada de sessão autenticada;
- mecanismo destinado a contornar políticas do Microsoft Entra ou
  controles corporativos equivalentes.

A existência de acesso legítimo do usuário através do navegador ou
Microsoft Excel não implica automaticamente autorização para reutilizar
essa sessão programaticamente.

---

# 22. LOGS

Logs não poderão expor:

- senha;
- token;
- segredo;
- credenciais.

Logs deverão ser úteis para diagnóstico sem comprometer segurança.

---

# 23. ALTERAÇÃO DA ARQUITETURA

Pequenos ajustes internos poderão ser realizados quando:

- preservarem requisitos;
- reduzirem complexidade;
- não alterarem comportamento;
- não afetarem segurança.

Mudanças estruturais relevantes deverão ser apresentadas antes da
implementação quando afetarem:

- persistência;
- identidade;
- segurança;
- integração;
- checkpoint;
- idempotência;
- fluxo principal.

---

# 24. ALTERAÇÃO DE REQUISITOS

Somente o responsável pelo projeto poderá autorizar mudança funcional
relevante.

Após autorização, atualizar primeiro ou conjuntamente a documentação
correspondente.

O código não deverá se tornar a única fonte de uma nova regra de
negócio.

---

# 25. HISTÓRICO DE IMPLEMENTAÇÕES

`04_HISTORICO_IMPLEMENTACOES.md` deverá refletir o estado real.

Não:

- inventar commits;
- inventar testes;
- antecipar conclusão;
- registrar funcionalidade planejada como implementada.

---

# 26. COMMITS

Cada commit deverá possuir objetivo compreensível.

Evitar commits genéricos como:

"updates"

"changes"

"fix stuff"

Preferir mensagens relacionadas ao resultado entregue.

Exemplo:

`Implementa persistência inicial da trilha de auditoria`

ou:

`Implementa comparação determinística de versões Excel`

---

# 27. ESTADO DO REPOSITÓRIO

Antes de alterar código, o agente deverá inspecionar o estado atual do
repositório.

Não presumir que o repositório está igual ao estado descrito em uma
conversa anterior.

Código existente deverá ser considerado antes de:

- criar arquivo duplicado;
- substituir implementação;
- remover componente;
- alterar estrutura.

---

# 28. CÓDIGO EXISTENTE VERSUS DOCUMENTAÇÃO

Se o código existente contrariar a documentação:

não assumir automaticamente que o código está correto.

Verificar:

1. histórico;
2. documentação;
3. commits;
4. intenção da implementação.

Corrigir ou solicitar decisão conforme a hierarquia de governança.

---

# 29. DOCUMENTAÇÃO ADICIONAL

Não criar automaticamente:

- catálogo de prompts;
- diário separado;
- mapa da documentação;
- agente separado;
- backlog adicional;
- documentos de arquitetura paralelos;
- checklists redundantes.

Os seis documentos atuais são suficientes para a V1.

Novo documento somente deverá existir quando resolver necessidade real.

---

# 30. PROMPT OFICIAL

`05_PROMPT_OFICIAL.md` deverá orientar o Codex a consultar, no mínimo:

- `06_GOVERNANCA.md`;
- `01_ESPECIFICACAO_FUNCIONAL.md`;
- `02_ARQUITETURA.md`;
- `03_PLANO_DE_DESENVOLVIMENTO.md`;
- `04_HISTORICO_IMPLEMENTACOES.md`.

O prompt não deverá duplicar toda a documentação.

Sua função é iniciar corretamente a sessão e aplicar as regras
documentadas.

---

# 31. ORDEM RECOMENDADA DE LEITURA PELO CODEX

Em cada nova sessão:

1. `06_GOVERNANCA.md`
2. `01_ESPECIFICACAO_FUNCIONAL.md`
3. `02_ARQUITETURA.md`
4. `03_PLANO_DE_DESENVOLVIMENTO.md`
5. `04_HISTORICO_IMPLEMENTACOES.md`

O `05_PROMPT_OFICIAL.md` é utilizado como instrução inicial da sessão.

---

# 32. PROCEDIMENTO DE INÍCIO DE SESSÃO

O agente deverá:

1. ler a documentação obrigatória;
2. inspecionar o repositório;
3. verificar `git status`;
4. identificar a fase atual;
5. identificar o último trabalho registrado;
6. verificar bloqueios;
7. determinar a próxima tarefa permitida;
8. executar somente o escopo autorizado.

---

# 33. PROCEDIMENTO DE ENCERRAMENTO

Antes de encerrar uma fase ou sessão relevante:

1. executar testes aplicáveis;
2. verificar alterações realizadas;
3. atualizar histórico;
4. registrar decisões importantes;
5. verificar quantidade de commits da fase;
6. informar resultado;
7. informar pendências;
8. informar próximo passo.

Não avançar automaticamente.

---

# 34. PRIORIDADE DE DECISÃO TÉCNICA

Quando houver mais de uma solução tecnicamente válida, priorizar:

1. integridade da auditoria;
2. segurança;
3. correção;
4. rastreabilidade;
5. idempotência;
6. simplicidade;
7. testabilidade;
8. desempenho;
9. aparência.

---

# 35. DEFINIÇÃO DE SUCESSO

O sucesso do projeto não será medido pela quantidade de:

- código;
- arquivos;
- frameworks;
- commits;
- funcionalidades extras.

Será medido pela capacidade de realizar corretamente o fluxo:

Selecionar planilha
        ↓
Estabelecer identidade técnica
        ↓
Adquirir versões do SharePoint
        ↓
Determinar checkpoint
        ↓
Comparar versões pendentes
        ↓
Registrar alterações
        ↓
Atualizar checkpoint
        ↓
Preservar histórico
        ↓
Gerar relatório consolidado
        ↓
Reexecutar sem duplicação

mantendo o SharePoint somente em leitura.

---

# 36. ALTERAÇÃO DESTA GOVERNANÇA

Este documento poderá ser alterado quando:

- surgir necessidade real;
- houver mudança de requisito;
- uma regra se mostrar inadequada;
- o responsável pelo projeto decidir alterar o processo.

Alterações deverão ser explícitas.

O agente não poderá modificar unilateralmente uma regra de governança
para facilitar sua própria implementação.

---

# 37. ESTADO INICIAL

Na criação deste documento:

Documentação oficial:

[✓] 01_ESPECIFICACAO_FUNCIONAL.md

[✓] 02_ARQUITETURA.md

[✓] 03_PLANO_DE_DESENVOLVIMENTO.md

[✓] 04_HISTORICO_IMPLEMENTACOES.md

[✓] 05_PROMPT_OFICIAL.md

[✓] 06_GOVERNANCA.md

Fase de implementação:

F1 — NÃO INICIADA

Próxima ação:

Iniciar F1 — Fundação e Banco de Auditoria somente após autorização do
responsável pelo projeto.

---

FIM DO DOCUMENTO