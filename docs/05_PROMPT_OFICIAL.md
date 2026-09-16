# PROJETO: AUDITOR DE PLANILHAS SHAREPOINT


# ESTADO OFICIAL APÓS A ACEITAÇÃO DA F5

F5 está concluída após validação real no Windows corporativo com SharePoint Online,
Selenium + Edge e migração do SQLite existente. Interface, aquisição e processamento de
versões, persistência, checkpoint e relatório com as três abas foram aceitos. O
SharePoint permanece read-only; cookies, tokens, senhas, PRT, bypass e transferência de
sessão continuam proibidos. A F6 foi iniciada com autorização expressa e deve avançar
uma tarefa por vez, conforme a ordem do plano. Testes unitários, testes de integração
automatizados, reexecução e auditoria incremental foram concluídos; a próxima tarefa é
falha e retomada e aguarda autorização. Os três commits ordinários da F6 foram consumidos.



======================================================================
DIRETRIZ ATUAL DA F4 — SHAREPOINT REST NO EDGE
======================================================================

Na F4, usar `BrowserSharePointSource` com Selenium/Edge visível, autenticação manual e
SharePoint REST executado na sessão do navegador. Não solicitar nem extrair cookies,
tokens, senha, PRT ou credenciais; não repassar a sessão a `requests`/`urllib`. Permitir
somente GET same-origin em `/_api/`. Descobrir `.xlsx` recursivamente em escopos
configurados, usar o `UniqueId` retornado como identidade técnica contextualizada e o
server-relative path apenas como localização. Enumerar histórico com `/Versions`, usar
sempre o ID retornado para download e preservar label/metadados separadamente. Validar
Open XML antes de comparar e nunca saltar uma versão com falha. Manter a baseline do
checkpoint e a idempotência. Graph é opcional/futuro. Não assumir que `/Versions` inclui
a versão atual. Não iniciar F5.
Você atuará como agente de desenvolvimento deste projeto.

Sua responsabilidade é implementar uma aplicação simples, funcional,
segura e rastreável para geração incremental de trilhas de auditoria de
planilhas Excel armazenadas no SharePoint Online.

Não transforme este projeto em uma arquitetura desnecessariamente complexa.
Priorize simplicidade, segurança, testabilidade e funcionamento real.


======================================================================
1. OBJETIVO DO SISTEMA
======================================================================

Desenvolver uma aplicação em Python capaz de:

1. Acessar planilhas Excel armazenadas no SharePoint Online.

2. Permitir selecionar uma planilha para auditoria.

3. Identificar a planilha por identidade técnica estável e inequívoca.

   Quando disponível através do mecanismo de aquisição adotado, utilizar
   preferencialmente o DriveItem ID e os identificadores de contexto
   necessários.

   O nome do arquivo NÃO deverá ser utilizado como identidade técnica
   exclusiva.

4. Consultar ou adquirir automaticamente o histórico de versões
   disponíveis da planilha,

5. Na primeira auditoria:
   - obter o histórico disponível;
   - comparar versões consecutivas;
   - comparar N → N+1;
   - identificar alterações célula a célula.

6. Nas auditorias posteriores:
   - consultar a última versão já auditada;
   - consultar as versões atualmente disponíveis no SharePoint;
   - continuar exatamente do checkpoint;
   - processar somente versões novas.

Exemplo:

Primeira execução:

0.84 → 0.85
0.85 → 0.86
...
0.98 → 0.99

Checkpoint = 0.99

Execução posterior, com SharePoint em 1.20:

0.99 → 1.00
1.00 → 1.01
...
1.19 → 1.20

Novo checkpoint = 1.20.

Se futuramente chegar a 5.99, continuar do checkpoint existente
até 5.99 sem reprocessar o histórico já consolidado.


======================================================================
2. REGRA FUNDAMENTAL DE SEGURANÇA
======================================================================

O SharePoint é SOMENTE FONTE DE LEITURA.

A aplicação NÃO poderá:

- alterar arquivos;
- sobrescrever arquivos;
- excluir arquivos;
- criar versões;
- excluir versões;
- executar check-in;
- executar check-out;
- alterar metadados;
- alterar comentários;
- modificar qualquer conteúdo do SharePoint.

A integração deve utilizar somente as permissões mínimas necessárias
para leitura.

Nenhuma implementação poderá violar esta regra.


======================================================================
3. TRILHA DE AUDITORIA
======================================================================

Para cada comparação N → N+1, identificar alterações nas células.

Tipos:

ADD = célula anteriormente vazia/inexistente recebeu conteúdo.

DEL = conteúdo existente foi removido.

MOD = conteúdo existente foi alterado.

Registrar pelo menos:

- identificador global da ação;
- identidade técnica da planilha;
- DriveItem ID, quando disponível;
- nome da planilha no momento da auditoria;
- versão anterior;
- versão atual;
- data/hora da versão;
- autor da versão;
- comentário da versão, quando disponível;
- aba;
- endereço da célula;
- tipo da alteração;
- valor anterior;
- valor novo.

IMPORTANTE:

Autor e data/hora pertencem à VERSÃO salva.

Nunca afirmar que o autor da versão é necessariamente o autor individual
da alteração da célula.

A aplicação não pode inventar granularidade que o SharePoint não fornece.


======================================================================
4. VERSÕES SEM ALTERAÇÕES
======================================================================

Toda versão processada deve ser registrada, mesmo quando a comparação
não produzir alterações de células.

O sistema deve ser capaz de distinguir:

"versão processada sem diferenças"

de:

"versão nunca processada".


======================================================================
5. CHECKPOINT E IDEMPOTÊNCIA
======================================================================

O checkpoint será individual por planilha e associado à sua identidade
técnica estável.

Quando DriveItem ID estiver disponível, ele deverá ser preservado como
parte da identidade/metadados técnicos conforme definido na arquitetura.

O sistema deverá garantir:

- processamento incremental;
- ausência de duplicação;
- retomada segura;
- reexecução sem duplicar registros;
- atualização do checkpoint somente após processamento bem-sucedido.

ATENÇÃO:

Se o checkpoint for 0.99 e a nova versão for 1.00, a versão 0.99 ainda
precisa ser recuperada como base para comparação:

0.99 → 1.00.

Não filtrar a versão-base necessária para a primeira comparação nova.


======================================================================
6. CONSOLIDAÇÃO E ARMAZENAMENTO
======================================================================

O banco de dados será a fonte oficial da trilha de auditoria.

Inicialmente utilizar SQLite.

A arquitetura deve permitir migração futura para SQL Server sem
reescrever o motor de comparação.

Os relatórios Excel são REPRESENTAÇÕES da trilha armazenada no banco,
não a fonte primária da evidência.

Uma alteração manual ou perda do relatório Excel não pode destruir
a trilha oficial.

Estrutura conceitual mínima:

PLANILHA
VERSAO_PROCESSADA
ALTERACAO
CHECKPOINT
EXECUCAO_AUDITORIA
ERRO_PROCESSAMENTO

Definir constraints/índices necessários para impedir duplicidade.


======================================================================
7. RELATÓRIO
======================================================================

Permitir gerar relatório individual por planilha.

Exemplo:

CQL028_Trilha_Auditoria.xlsx

O relatório deverá conter inicialmente:

ABA: RESUMO
- planilha;
- DriveItem ID;
- primeira versão auditada;
- última versão auditada;
- quantidade de versões;
- quantidade total de alterações;
- ADD;
- MOD;
- DEL;
- última execução.

ABA: VERSOES
- versão;
- data/hora;
- autor;
- comentário;
- status;
- quantidade de alterações.

ABA: TRILHA
- ID;
- versão anterior;
- versão atual;
- data/hora;
- autor da versão;
- comentário;
- aba;
- célula;
- tipo;
- valor anterior;
- valor novo.

O relatório deve possuir filtros e apresentação simples para consulta.

O usuário poderá gerar novamente o relatório a qualquer momento a partir
do banco consolidado.


======================================================================
8. INTEGRIDADE
======================================================================

Registrar as execuções da auditoria.

Cada execução deve possuir pelo menos:

- ID da execução;
- início;
- término;
- status;
- planilha;
- versão inicial;
- versão final;
- quantidade de versões processadas;
- quantidade de alterações;
- erros.

Avaliar utilização de SHA-256 para reforçar a verificação de integridade
das evidências processadas, sem criar complexidade desnecessária.

Não implementar mecanismos criptográficos sem finalidade claramente
documentada.


======================================================================
9. TECNOLOGIAS
======================================================================

Obrigatório:

Obrigatório:

- Python 3.x
- openpyxl
- SQLite inicialmente
- integração suportada e autorizada com SharePoint Online
- pytest para testes

Microsoft Graph API permanece como mecanismo de integração suportado
pela arquitetura quando estiver disponível e autorizado no ambiente,
mas não constitui dependência obrigatória ou exclusiva da V1.

Dependências específicas de aquisição SharePoint somente deverão ser
adicionadas após validação do mecanismo efetivamente adotado.
A aplicação NÃO DEVE possuir dependência de:


- Node.js
- npm
- React
- Angular
- Vue
- Vite

Não introduzir frontend JavaScript complexo.

Se interface gráfica for necessária, utilizar tecnologia compatível
com execução Python e manter a solução simples.

Priorizar bibliotecas Python maduras e com manutenção ativa.

É proibido introduzir mecanismo destinado a contornar políticas
corporativas de autenticação ou autorização.


======================================================================
10. INTERFACE
======================================================================

A V1 deve possuir uma interface simples e funcional.

Fluxo esperado:

1. Abrir aplicação.
2. Conectar/configurar SharePoint.
3. Visualizar planilhas elegíveis.
4. Selecionar uma planilha.
5. Visualizar:
   - nome;
   - status da auditoria;
   - última versão auditada;
   - última versão disponível;
   - quantidade de versões pendentes.
6. Executar:
   "Auditar histórico completo"
   ou
   "Continuar auditoria".
7. Visualizar resultado.
8. Gerar/abrir relatório Excel.

Não priorizar aparência sobre funcionalidade.


======================================================================
11. ARQUITETURA
======================================================================

Manter separação entre:

- acesso ao SharePoint;
- leitura Excel;
- comparação;
- persistência;
- serviço de auditoria;
- geração de relatório;
- interface.

O motor de auditoria NÃO deverá depender diretamente do Microsoft Graph
nem de qualquer mecanismo específico de aquisição SharePoint.

Deve ser possível utilizar uma fonte local simulada para testes.

Exemplo conceitual:

             Aplicação
                 |
          Serviço Auditoria
          /       |       \
         /        |        \
    VersionSource      Comparator    Banco
   Local/
   SharePoint
                 |
              Relatório

A SharePointSource deverá encapsular o mecanismo concreto de aquisição.

Microsoft Graph poderá ser utilizado quando autorizado.

Outro mecanismo somente poderá ser utilizado quando for suportado,
autorizado, exclusivamente de leitura e validado tecnicamente.


======================================================================
12. TESTES
======================================================================

NÍVEL 1 — LOCAL

Utilizar arquivos Excel simulando versões.

Validar obrigatoriamente:

- ADD;
- DEL;
- MOD;
- fórmulas;
- múltiplas abas;
- versões sem alteração;
- checkpoint;
- reexecução;
- ausência de duplicação;
- retomada após checkpoint.

NÍVEL 2 — AQUISIÇÃO SHAREPOINT CONTROLADA

Utilizar uma única planilha real.

Validar:

- identidade técnica estável;
- DriveItem ID, quando disponível;
- descoberta/listagem das versões;
- aquisição/leitura das versões históricas;
- metadados disponíveis;
- ordenação das versões;
- versões principais/secundárias necessárias;
- comparação;
- checkpoint;
- nenhuma alteração no SharePoint.

NÍVEL 3 — ESCALA

Somente depois dos níveis anteriores.

Testar múltiplas planilhas e volume maior de versões.


======================================================================
13. DOCUMENTAÇÃO DO PROJETO
======================================================================

Manter somente documentação necessária.

Criar:

docs/
  01_ESPECIFICACAO_FUNCIONAL.md
  02_ARQUITETURA.md
  03_PLANO_DE_DESENVOLVIMENTO.md
  04_HISTORICO_IMPLEMENTACOES.md
  05_PROMPT_OFICIAL.md
  06_GOVERNANCA.md

Antes de iniciar qualquer implementação, consultar obrigatoriamente,
nesta ordem:

1. docs/06_GOVERNANCA.md
2. docs/01_ESPECIFICACAO_FUNCIONAL.md
3. docs/02_ARQUITETURA.md
4. docs/03_PLANO_DE_DESENVOLVIMENTO.md
5. docs/04_HISTORICO_IMPLEMENTACOES.md

Este arquivo, 05_PROMPT_OFICIAL.md, contém as instruções operacionais
para a sessão do agente de desenvolvimento.

O documento 06_GOVERNANCA.md possui precedência sobre os demais
documentos conforme a hierarquia definida nele.

Não criar documentos adicionais sem necessidade técnica clara ou
solicitação do responsável pelo projeto.

01_ESPECIFICACAO_FUNCIONAL.md:
requisitos, regras de negócio, limitações e critérios de aceite.

02_ARQUITETURA.md:
componentes, banco, fluxo de dados e decisões técnicas.

03_PLANO_DE_DESENVOLVIMENTO.md:
fases, tarefas e status.

04_HISTORICO_IMPLEMENTACOES.md:
registro resumido do que efetivamente foi implementado/testado.


======================================================================
14. ESTRATÉGIA DE DESENVOLVIMENTO
======================================================================

O desenvolvimento deve ser CURTO E INCREMENTAL.

REGRA OBRIGATÓRIA:

CADA FASE PODERÁ POSSUIR NO MÁXIMO 3 COMMITS.

Não fragmentar artificialmente uma implementação apenas para gerar
mais commits.

Preferência:

1 ou 2 commits por fase.

Utilizar 3 somente quando tecnicamente justificável.

Não criar dezenas de microtarefas para funcionalidades simples.


======================================================================
15. FASES INICIAIS
======================================================================

Planejar a V1 em aproximadamente 6 fases:

FASE 1 — Fundação + banco
FASE 2 — Motor Excel + comparação
FASE 3 — Auditor local incremental
FASE 4 — Aquisição de Versões SharePoint
FASE 5 — Interface + relatório consolidado
FASE 6 — Robustez, testes reais e preparação para produção

A Fase 4 deverá utilizar Microsoft Graph quando disponível e autorizado
ou outro mecanismo suportado e autorizado conforme definido na
arquitetura.

A seleção de mecanismo alternativo não autoriza contorno de controles
corporativos de segurança.

Cada fase deve entregar algo executável/testável.

Não iniciar uma nova fase enquanto os critérios de aceite da fase atual
não forem atendidos.


======================================================================
16. PRINCÍPIOS DE IMPLEMENTAÇÃO
======================================================================

Prioridades, nesta ordem:

1. Integridade da trilha.
2. Segurança de leitura do SharePoint.
3. Correção da comparação.
4. Idempotência.
5. Simplicidade.
6. Performance.
7. Interface/aparência.

Evitar:

- overengineering;
- abstrações prematuras;
- microsserviços;
- infraestrutura desnecessária;
- dependências desnecessárias;
- frameworks pesados;
- Node.js;
- funcionalidades não solicitadas.


======================================================================
17. REGRA PARA DÚVIDAS E LIMITAÇÕES DE INTEGRAÇÃO
======================================================================

Não inventar comportamento do SharePoint, Microsoft Graph ou qualquer
mecanismo de aquisição.

Se determinada capacidade depender da fonte, especialmente:

- autenticação;
- identidade técnica;
- descoberta de versões;
- recuperação de versões históricas;
- versões principais/secundárias;
- metadados;
- ordenação;

confirmar tecnicamente antes de implementar.

Se existir limitação que afete algum requisito:

1. interromper a implementação afetada;
2. registrar o requisito;
3. registrar o comportamento observado;
4. registrar o impacto;
5. apresentar alternativas suportadas;
6. aguardar decisão quando houver mudança de arquitetura, segurança
   ou escopo.

Não implementar workaround destrutivo ou mecanismo destinado a
contornar autenticação/autorização corporativa.

Não extrair cookies, tokens, senhas ou credenciais de sessões existentes.

Acesso através do navegador ou Microsoft Excel não deverá ser tratado
automaticamente como autorização para acesso programático.


======================================================================
18. PROCEDIMENTO OBRIGATÓRIO EM CADA SESSÃO
======================================================================

NÃO comece implementando funcionalidades automaticamente.

Antes de qualquer alteração no código:

1. Leia integralmente, nesta ordem:
   - docs/06_GOVERNANCA.md
   - docs/01_ESPECIFICACAO_FUNCIONAL.md
   - docs/02_ARQUITETURA.md
   - docs/03_PLANO_DE_DESENVOLVIMENTO.md
   - docs/04_HISTORICO_IMPLEMENTACOES.md

2. Inspecione o estado atual do repositório.

3. Verifique:
   - git status;
   - branch atual;
   - commits recentes;
   - estrutura existente;
   - alterações não commitadas.

4. Não apague ou substitua código existente sem justificativa.

5. Identifique no PLANO_DE_DESENVOLVIMENTO a fase atualmente autorizada.

6. Consulte o HISTORICO_IMPLEMENTACOES para verificar:
   - fases concluídas;
   - fase atual;
   - commits realizados;
   - decisões;
   - bloqueios;
   - limitações;
   - pendências;
   - próximo passo registrado.

7. Execute SOMENTE a fase/tarefa expressamente autorizada.

8. Não implemente funcionalidades de fases posteriores.

9. Quando a tarefa autorizada for apenas investigação técnica:
   - não transformar investigação em implementação automaticamente;
   - realizar somente testes seguros e necessários;
   - registrar resultados observados;
   - não alterar arquitetura silenciosamente.

10. Executar os testes aplicáveis ao trabalho realizado.

11. Atualizar:
    docs/04_HISTORICO_IMPLEMENTACOES.md

    registrando somente fatos reais:
    - implementação realizada;
    - investigação realizada;
    - arquivos criados/alterados;
    - testes efetivamente executados;
    - resultados;
    - decisões;
    - bloqueios;
    - limitações;
    - commits reais.

12. Respeitar o limite de commits definido pelo plano e as exceções
    formalmente autorizadas.

13. Ao finalizar, apresentar:

    - fase;
    - status;
    - tarefa executada;
    - resumo;
    - arquivos criados;
    - arquivos alterados;
    - testes executados;
    - resultados;
    - commits;
    - problemas;
    - bloqueios;
    - pendências;
    - riscos;
    - próximo passo recomendado.

14. NÃO iniciar a próxima fase.

Depois:

PARE E AGUARDE AUTORIZAÇÃO DO RESPONSÁVEL PELO PROJETO.
