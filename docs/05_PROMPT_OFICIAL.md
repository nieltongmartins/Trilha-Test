# PROJETO: AUDITOR DE PLANILHAS SHAREPOINT

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

3. Identificar a planilha tecnicamente pelo DriveItem ID do SharePoint.
   O nome do arquivo NÃO é sua identidade técnica.

4. Consultar o histórico de versões disponíveis da planilha,
   incluindo versões principais e secundárias recuperáveis.

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
- DriveItem ID da planilha;
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

O checkpoint será individual por planilha e associado ao DriveItem ID.

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

- Python 3.x
- openpyxl
- SQLite inicialmente
- Microsoft Graph API para SharePoint
- pytest para testes

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

O motor de auditoria NÃO deve depender diretamente do Microsoft Graph.

Deve ser possível utilizar uma fonte local simulada para testes.

Exemplo conceitual:

             Aplicação
                 |
          Serviço Auditoria
          /       |       \
         /        |        \
    Fonte      Comparator    Banco
   Local/
   Graph
                 |
              Relatório


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

NÍVEL 2 — SHAREPOINT CONTROLADO

Utilizar uma única planilha real.

Validar:

- DriveItem ID;
- listagem das versões;
- download/leitura de versões;
- metadados;
- versões principais/secundárias disponíveis;
- comparação;
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
FASE 4 — Integração Microsoft Graph / SharePoint
FASE 5 — Interface + relatório consolidado
FASE 6 — Robustez, testes reais e preparação para produção

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
17. REGRA PARA DÚVIDAS
======================================================================

Não inventar comportamento do SharePoint ou Microsoft Graph.

Se determinada capacidade depender da API, especialmente recuperação
de versões históricas principais/secundárias, confirmar tecnicamente
antes de implementar.

Se existir limitação da Microsoft Graph que afete algum requisito,
interromper aquela implementação e informar:

1. requisito afetado;
2. limitação encontrada;
3. impacto;
4. alternativas possíveis.

Não implementar workaround destrutivo.


======================================================================
18. SUA PRIMEIRA TAREFA
======================================================================

NÃO comece implementando toda a aplicação.

Antes de qualquer alteração no código:

1. Leia integralmente:
   - docs/06_GOVERNANCA.md
   - docs/01_ESPECIFICACAO_FUNCIONAL.md
   - docs/02_ARQUITETURA.md
   - docs/03_PLANO_DE_DESENVOLVIMENTO.md
   - docs/04_HISTORICO_IMPLEMENTACOES.md

2. Inspecione o estado atual do repositório.

3. Verifique o git status e a estrutura existente.

4. Não apague ou substitua código existente sem justificativa.

5. Identifique no PLANO_DE_DESENVOLVIMENTO a fase atualmente
   autorizada.

6. Consulte o HISTORICO_IMPLEMENTACOES para verificar:
   - fases concluídas;
   - fase atual;
   - commits realizados;
   - bloqueios;
   - pendências;
   - próximo passo registrado.

7. Nesta primeira execução de desenvolvimento, trabalhar SOMENTE na:

   FASE 1 — FUNDAÇÃO E BANCO DE AUDITORIA.

8. Implementar somente os entregáveis e critérios de aceite previstos
   para a Fase 1.

9. Executar os testes previstos para a fase.

10. Atualizar:
    docs/04_HISTORICO_IMPLEMENTACOES.md

    registrando somente fatos reais:
    - implementação realizada;
    - arquivos criados/alterados;
    - testes efetivamente executados;
    - resultados;
    - decisões;
    - bloqueios;
    - commits reais.

11. Respeitar o limite máximo de 3 commits para a fase,
    preferencialmente utilizando 1 ou 2.

12. NÃO iniciar a Fase 2.

Ao terminar, apresente:

- status da Fase 1;
- resumo do que foi implementado;
- arquivos criados;
- arquivos alterados;
- estrutura resultante;
- testes executados;
- resultado dos testes;
- commits realizados;
- problemas ou bloqueios;
- pendências;
- próximo passo recomendado.

Depois:

PARE E AGUARDE AUTORIZAÇÃO DO RESPONSÁVEL PELO PROJETO.