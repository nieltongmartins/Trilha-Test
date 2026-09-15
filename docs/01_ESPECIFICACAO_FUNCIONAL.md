# ESPECIFICAÇÃO FUNCIONAL
## Auditor de Planilhas Excel — SharePoint Online

**Documento:** 01_ESPECIFICACAO_FUNCIONAL.md  
**Versão:** 1.0  
**Status:** Oficial  
**Data:** 15/09/2026  

---

# 1. OBJETIVO

Este documento define os requisitos funcionais, regras de negócio,
limitações e critérios de aceite da aplicação Auditor de Planilhas Excel.

A aplicação tem como finalidade gerar e manter uma trilha de auditoria
incremental de planilhas Excel armazenadas no SharePoint Online.

O sistema deverá atuar como auditor externo das planilhas.

O SharePoint será exclusivamente uma fonte de dados em modo leitura.

A aplicação não poderá alterar arquivos, versões, metadados ou qualquer
outro conteúdo existente no SharePoint.

---

# 2. OBJETIVO PRINCIPAL DO SISTEMA

A aplicação deverá permitir:

- acessar planilhas Excel armazenadas no SharePoint Online;
- identificar cada planilha por seu identificador técnico imutável;
- consultar o histórico de versões disponível;
- recuperar versões históricas acessíveis;
- comparar versões consecutivas;
- identificar alterações célula a célula;
- armazenar as alterações encontradas;
- manter uma trilha consolidada por planilha;
- continuar auditorias anteriores de forma incremental;
- gerar relatórios de auditoria;
- garantir que reexecuções não gerem duplicidades.

---

# 3. UNIDADE DE AUDITORIA

A unidade de auditoria será uma planilha Excel individual.

Exemplo:

CQL028.xlsx

Cada planilha possuirá sua própria:

- identificação;
- relação de versões processadas;
- trilha de alterações;
- checkpoint;
- histórico de execuções;
- relatório.

---

# 4. IDENTIFICAÇÃO DA PLANILHA

O nome do arquivo NÃO será utilizado como chave técnica.

A identidade da planilha será baseada no identificador fornecido pelo
SharePoint/Microsoft Graph, preferencialmente o DriveItem ID.

Exemplo conceitual:

Nome:

CQL028.xlsx

Identidade técnica:

DriveItem ID = 01ABCDEF...

Caso o arquivo seja renomeado, a aplicação deverá continuar
reconhecendo-o como a mesma planilha enquanto sua identidade técnica
permanecer a mesma.

O nome deverá ser armazenado apenas como informação descritiva.

---

# 5. SHAREPOINT COMO FONTE DE DADOS

O SharePoint Online será considerado a fonte oficial das versões das
planilhas.

A aplicação deverá operar exclusivamente em modo leitura.

É proibido à aplicação:

- alterar planilhas;
- salvar planilhas no SharePoint;
- sobrescrever arquivos;
- excluir arquivos;
- excluir versões;
- criar versões;
- executar check-in;
- executar check-out;
- alterar comentários;
- alterar metadados;
- alterar permissões;
- modificar qualquer conteúdo do SharePoint.

As permissões concedidas à aplicação deverão seguir o princípio do
menor privilégio.

---

# 6. HISTÓRICO DE VERSÕES

A aplicação deverá consultar as versões disponíveis para cada planilha
selecionada.

Exemplo:

0.84
0.85
0.86
0.87
...
0.98
0.99

O objetivo é considerar tanto versões principais quanto secundárias,
desde que possam ser enumeradas e recuperadas através das APIs
disponíveis no ambiente.

A capacidade real de recuperação dessas versões deverá ser validada
durante a integração com o Microsoft Graph/SharePoint.

A aplicação não deverá presumir que uma versão pode ser recuperada
apenas porque aparece na interface web do SharePoint.

---

# 7. PRIMEIRA AUDITORIA

Quando uma planilha nunca tiver sido auditada, o sistema deverá permitir
executar uma auditoria do histórico disponível.

Exemplo:

Versões disponíveis:

0.84
0.85
0.86
...
0.98
0.99

As comparações deverão ocorrer sequencialmente:

0.84 → 0.85
0.85 → 0.86
0.86 → 0.87
...
0.98 → 0.99

Ao término:

Última versão auditada = 0.99

Essa informação será utilizada como checkpoint.

---

# 8. AUDITORIA INCREMENTAL

Quando uma planilha já possuir trilha consolidada, a aplicação deverá
continuar do último checkpoint.

Exemplo:

Última versão auditada:

0.99

Versão atual disponível:

1.20

O sistema deverá processar:

0.99 → 1.00
1.00 → 1.01
1.01 → 1.02
...
1.19 → 1.20

Ao final:

Última versão auditada = 1.20

Se futuramente a planilha chegar à versão 5.99, a auditoria deverá
continuar do checkpoint existente até a última versão disponível.

O histórico anteriormente consolidado NÃO deverá ser processado
novamente durante uma execução incremental normal.

---

# 9. REGRA DO CHECKPOINT

O checkpoint será armazenado individualmente por planilha.

O checkpoint deverá estar associado à identidade técnica da planilha,
e não ao nome do arquivo.

O checkpoint somente poderá avançar quando a comparação correspondente
for concluída e persistida com sucesso.

Exemplo:

Checkpoint:

0.99

Nova versão:

1.00

A aplicação deverá manter acesso à versão 0.99 para realizar:

0.99 → 1.00

Somente depois da conclusão bem-sucedida dessa comparação o checkpoint
poderá ser atualizado para 1.00.

---

# 10. IDEMPOTÊNCIA

A aplicação deverá ser idempotente.

Isso significa que executar novamente a auditoria sem novas versões não
poderá gerar novas alterações ou duplicar registros existentes.

Exemplo:

Primeira execução:

Checkpoint = 1.20

Segunda execução:

Última versão SharePoint = 1.20

Resultado esperado:

Novas versões = 0
Novas comparações = 0
Novas alterações = 0

O banco deverá possuir proteções adicionais contra duplicidade.

---

# 11. COMPARAÇÃO DE VERSÕES

A comparação será realizada entre versões consecutivas da mesma
planilha.

O motor deverá comparar:

- abas;
- células;
- valores;
- fórmulas, quando aplicável.

A comparação básica será:

VERSÃO N → VERSÃO N+1

A aplicação deverá identificar todas as diferenças relevantes
encontradas entre os dois estados.

---

# 12. TIPOS DE ALTERAÇÃO

Serão utilizados inicialmente três tipos de alteração.

## 12.1 ADD

Uma célula anteriormente vazia ou inexistente recebeu conteúdo.

Exemplo:

Antes:

A1 = vazio

Depois:

A1 = "Aprovado"

Resultado:

ADD

---

## 12.2 DEL

Uma célula anteriormente preenchida passou a estar vazia.

Exemplo:

Antes:

B10 = 150

Depois:

B10 = vazio

Resultado:

DEL

---

## 12.3 MOD

Uma célula possuía conteúdo e seu conteúdo foi alterado.

Exemplo:

Antes:

G22 = 150

Depois:

G22 = 180

Resultado:

MOD

---

# 13. LINHAS E COLUNAS

A inserção ou remoção de linhas e colunas poderá resultar em múltiplas
alterações de células.

A V1 não deverá utilizar heurísticas para afirmar que determinada linha
foi inserida ou removida se essa informação não puder ser determinada
de forma confiável.

O objetivo principal é registrar o estado anterior e posterior das
células.

---

# 14. ABAS

O motor deverá considerar todas as abas elegíveis da pasta de trabalho.

Deverão ser detectadas diferenças resultantes de:

- células modificadas;
- células adicionadas;
- células removidas;
- abas adicionadas;
- abas removidas.

A representação de criação ou remoção de abas deverá ser definida pelo
motor de auditoria de forma determinística.

---

# 15. FÓRMULAS

A aplicação deverá preservar e comparar fórmulas.

O carregamento das planilhas deverá utilizar configuração adequada para
que a fórmula possa ser analisada como conteúdo da célula.

Exemplo:

Antes:

=SUM(A1:A10)

Depois:

=SUM(A1:A20)

Deverá ser registrado como alteração.

A V1 não deverá tentar reproduzir o mecanismo de cálculo do Microsoft
Excel.

---

# 16. AUTOR DA ALTERAÇÃO

O SharePoint fornece informações relacionadas à versão salva.

Portanto, o sistema poderá registrar:

- autor da versão;
- data/hora da versão.

O sistema NÃO deverá afirmar que esse usuário foi necessariamente o
autor individual da edição de cada célula.

No relatório deverá ser utilizada terminologia como:

"Autor da versão"

e não:

"Autor da célula".

---

# 17. DATA E HORA

A data/hora disponível será associada à versão salva.

A aplicação não deverá afirmar que essa data/hora corresponde ao
instante exato em que determinada célula foi editada.

Essa limitação deverá permanecer documentada.

---

# 18. COMENTÁRIO DA VERSÃO

Quando disponível através da integração, o comentário da versão deverá
ser armazenado.

Exemplo:

Versão:

0.99

Comentário:

"Inserção de resultados até o dia 14/09"

O comentário será considerado metadado complementar da evidência.

---

# 19. VERSÕES SEM DIFERENÇAS

Uma versão processada deverá ser registrada mesmo quando nenhuma
diferença de célula for encontrada.

O sistema deverá distinguir claramente:

VERSÃO PROCESSADA SEM ALTERAÇÕES

de:

VERSÃO NÃO PROCESSADA

Essa distinção é obrigatória para rastreabilidade.

---

# 20. BANCO DE AUDITORIA

O banco de dados será a fonte oficial da trilha consolidada.

A V1 utilizará SQLite.

A arquitetura deverá permitir futura migração para SQL Server sem
reescrever o motor de comparação.

O banco deverá armazenar, no mínimo:

- planilhas;
- versões processadas;
- alterações;
- checkpoints;
- execuções;
- erros de processamento.

---

# 21. TRILHA CONSOLIDADA

A trilha deverá crescer de forma incremental.

Exemplo:

Primeira auditoria:

0.84 → 0.99

Segunda auditoria:

0.99 → 1.20

Terceira auditoria:

1.20 → 1.50

O resultado consolidado deverá representar continuamente:

0.84 → ... → 0.99 → ... → 1.20 → ... → 1.50

Nenhum registro histórico válido deverá ser sobrescrito durante uma
execução incremental normal.

---

# 22. RELATÓRIO POR PLANILHA

O sistema deverá permitir gerar relatório individual.

Exemplo:

CQL028_Trilha_Auditoria.xlsx

O relatório será uma representação dos dados existentes no banco.

O arquivo Excel NÃO será a fonte oficial da evidência.

Se o relatório for perdido ou alterado, deverá ser possível gerá-lo
novamente a partir do banco.

---

# 23. ESTRUTURA DO RELATÓRIO

O relatório deverá possuir inicialmente três abas.

## 23.1 RESUMO

Informações mínimas:

- nome da planilha;
- DriveItem ID;
- primeira versão auditada;
- última versão auditada;
- data da última auditoria;
- quantidade de versões processadas;
- quantidade de alterações;
- quantidade de ADD;
- quantidade de MOD;
- quantidade de DEL.

---

## 23.2 VERSOES

Informações mínimas:

- versão anterior;
- versão atual;
- data/hora;
- autor da versão;
- comentário;
- status;
- quantidade de alterações.

---

## 23.3 TRILHA

Informações mínimas:

- ID da ação;
- versão anterior;
- versão atual;
- data/hora da versão;
- autor da versão;
- comentário;
- aba;
- endereço da célula;
- tipo da alteração;
- valor anterior;
- valor novo.

O relatório deverá utilizar filtros e formatação simples para facilitar
consulta e análise.

---

# 24. HISTÓRICO DE EXECUÇÕES

Cada execução deverá possuir identificação própria.

Exemplo:

AUD-20260915-000001

Deverão ser registrados, no mínimo:

- ID da execução;
- data/hora inicial;
- data/hora final;
- planilha;
- DriveItem ID;
- checkpoint inicial;
- versão final alcançada;
- quantidade de versões processadas;
- quantidade de alterações encontradas;
- status;
- mensagem de erro, quando aplicável.

---

# 25. ERROS

Uma falha não poderá resultar em avanço incorreto do checkpoint.

Exemplo:

0.99 → 1.00 = sucesso
1.00 → 1.01 = erro
1.01 → 1.02 = não processado

O checkpoint deverá permanecer em:

1.00

A próxima execução deverá poder retomar:

1.00 → 1.01

Erros deverão ser registrados para diagnóstico.

---

# 26. INTEGRIDADE DA EVIDÊNCIA

A aplicação deverá preservar a integridade lógica dos registros.

Deverão ser avaliados mecanismos como SHA-256 para verificar a
integridade das evidências processadas.

A utilização de hash deverá possuir finalidade técnica claramente
documentada.

A implementação não deverá criar falsa garantia de segurança ou
imutabilidade.

---

# 27. SELEÇÃO DA PLANILHA

A interface deverá permitir visualizar planilhas elegíveis disponíveis
no escopo configurado do SharePoint.

O usuário poderá selecionar uma planilha.

Exemplo:

CQL028.xlsx

Após a seleção, deverão ser apresentadas informações como:

- nome;
- status;
- última versão auditada;
- última versão disponível;
- quantidade de versões pendentes.

---

# 28. MODOS DE AUDITORIA

A aplicação deverá oferecer pelo menos dois comportamentos.

## Primeira auditoria

Quando não existir checkpoint:

AUDITAR HISTÓRICO DISPONÍVEL

## Auditoria existente

Quando existir checkpoint:

CONTINUAR AUDITORIA

O sistema deverá determinar automaticamente a posição correta de
continuidade.

---

# 29. INTERFACE

A V1 deverá possuir interface simples e funcional.

A aparência visual é secundária.

O fluxo esperado é:

1. iniciar aplicação;
2. acessar fonte configurada;
3. listar planilhas;
4. selecionar planilha;
5. consultar situação;
6. iniciar auditoria;
7. acompanhar resultado;
8. consultar trilha;
9. gerar relatório.

A aplicação não deverá depender de Node.js.

---

# 30. TECNOLOGIAS

Tecnologias principais:

- Python 3.x;
- openpyxl;
- SQLite;
- Microsoft Graph API;
- pytest.

É proibida dependência obrigatória de:

- Node.js;
- npm;
- React;
- Angular;
- Vue;
- Vite.

A interface, caso utilize biblioteca adicional, deverá permanecer
compatível com execução Python simples.

---

# 31. TESTE LOCAL

Antes da integração com SharePoint, o sistema deverá funcionar com uma
fonte local simulada.

Exemplo:

dados_teste/
  CQL028/
    0.84.xlsx
    0.85.xlsx
    0.86.xlsx
    0.87.xlsx

Isso deverá permitir validar o motor independentemente da rede,
credenciais ou disponibilidade do SharePoint.

---

# 32. TESTE SHAREPOINT

Após validação local, utilizar uma única planilha real controlada.

Deverão ser comprovados:

- identificação por DriveItem ID;
- listagem das versões;
- recuperação das versões;
- recuperação dos metadados disponíveis;
- acesso a versões principais/secundárias necessárias;
- ausência total de escrita no SharePoint.

Somente depois disso o processamento em escala poderá ser habilitado.

---

# 33. ESCALABILIDADE

O projeto deverá considerar utilização futura com:

- centenas de planilhas;
- centenas ou milhares de versões;
- grande quantidade de alterações.

Entretanto, otimizações complexas não deverão ser implementadas
prematuramente.

Primeiro deverá ser comprovada a correção funcional.

---

# 34. CRITÉRIOS DE ACEITE DA V1

A V1 será considerada funcional quando for possível:

1. selecionar uma planilha real do SharePoint;

2. identificar seu DriveItem ID;

3. consultar as versões históricas necessárias;

4. recuperar versões sem alterar o SharePoint;

5. comparar versões consecutivas;

6. identificar corretamente ADD, DEL e MOD;

7. armazenar a trilha no banco;

8. registrar versões sem diferenças;

9. estabelecer checkpoint;

10. executar novamente sem duplicidade;

11. encontrar versões novas;

12. continuar do checkpoint;

13. gerar relatório Excel individual;

14. reconstruir o relatório a partir do banco;

15. registrar erros e execuções;

16. comprovar que nenhuma operação de escrita foi realizada no
    SharePoint.

---

# 35. LIMITAÇÕES CONHECIDAS

São reconhecidas inicialmente as seguintes limitações:

- não é possível determinar necessariamente o instante exato de edição
  individual de uma célula;

- não é possível determinar necessariamente o autor individual de cada
  célula;

- autor/data disponíveis são associados à versão salva;

- alterações não persistidas em uma versão não constituem evidência
  histórica recuperável pelo auditor;

- disponibilidade de versões históricas depende das capacidades e
  políticas do SharePoint/Microsoft Graph;

- fórmulas serão comparadas, mas a aplicação não pretende reproduzir
  integralmente o mecanismo de cálculo do Excel.

Essas limitações deverão ser preservadas na documentação e não deverão
ser mascaradas por inferências.

---

# 36. PRINCÍPIOS DO PRODUTO

A implementação deverá seguir, nesta ordem:

1. integridade;
2. segurança;
3. rastreabilidade;
4. correção;
5. idempotência;
6. simplicidade;
7. desempenho;
8. aparência.

A aplicação deverá resolver o problema proposto sem adicionar
complexidade que não produza benefício direto.

---

# 37. CONTROLE DE ESCOPO

Funcionalidades não descritas neste documento não deverão ser
implementadas automaticamente.

Quando surgir necessidade de alteração relevante de requisito:

1. registrar a necessidade;
2. avaliar impacto;
3. atualizar a documentação correspondente;
4. obter decisão do responsável pelo projeto;
5. somente então implementar.

---

# 38. REFERÊNCIA OFICIAL

Este documento é a referência oficial dos requisitos funcionais da V1
do Auditor de Planilhas Excel.

Em caso de dúvida sobre comportamento funcional, consultar este
documento antes da implementação.

Conflitos ou lacunas deverão ser reportados ao responsável pelo projeto,
e não resolvidos através de suposições silenciosas.

---

FIM DO DOCUMENTO