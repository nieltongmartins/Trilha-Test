# ARQUITETURA DA SOLUÇÃO
## Auditor de Planilhas Excel — SharePoint Online

**Documento:** 02_ARQUITETURA.md  
Versão: 1.1
Status: Oficial
Data: 15/09/2026
Última revisão: 15/09/2026

---

# ENCERRAMENTO ARQUITETURAL F4

A arquitetura final da F4 é `BrowserSharePointSource -> VersionSource -> AuditService ->
SQLite`. Selenium abre Edge visível com diretório temporário de download; autenticação e
MFA são manuais. Consultas JSON usam `fetch` GET same-origin dentro da página e binários
são baixados pelo próprio Edge. Não existe exportação de sessão para cliente HTTP.

A descoberta usa `GetFolderByServerRelativeUrl(...)/Files` e `/Folders` de forma recursiva
nos escopos autorizados. A chave persistida é `(site_url, "sharepoint-rest", UniqueId)`;
nome e caminho são atualizados como atributos mutáveis. Histórico usa `/Versions` com
`CreatedBy`; atual usa os metadados do arquivo e `/$value`. Ambos implementam o mesmo
`VersionSource` já consumido pelo motor incremental. Downloads incompletos ou Open XML
inválido encerram a execução, preservando a última transação/checkpoint confirmada.

`GraphSharePointSource` permanece isolada e opcional para ambientes autorizados; a
política corporativa de consentimento impediu seu uso como provider obrigatório da V1.
Nenhum endpoint de escrita SharePoint integra a solução. F4 está concluída; F5 não foi
iniciada.

# REVISÃO ARQUITETURAL INTERMEDIÁRIA F4 — REGISTRO HISTÓRICO

Fluxo oficial da V1: `BrowserSharePointSource` descobre arquivos recursivamente nos
escopos configurados, normaliza identidade/proveniência e enumera/baixa versões;
`AuditService` mantém pares adjacentes, transação e checkpoint; SQLite continua fonte
canônica. O provider Graph permanece isolado em `app/sources/graph.py` e não é
dependência da V1.

O transporte do provider oficial recebe somente URL e resposta através de Selenium.
Ele permite exclusivamente GET same-origin sob `/_api/`, usa `credentials: same-origin`
dentro do Edge e nunca lê cookies, storage ou tokens. A sessão começa em Edge visível
e o usuário conclui a autenticação Microsoft normalmente.

Descoberta: `GetFolderByServerRelativeUrl(...)/Files` e `/Folders`, recursivamente;
identidade: `File.UniqueId` retornado pelo REST, com site e escopo como contexto;
localização: `ServerRelativeUrl`; versões históricas: `/Versions?$expand=CreatedBy`;
conteúdo: `/Versions(ID)/$value`. O ID é preservado separadamente do label e fornece
a ordem técnica observada. Os binários são validados antes da comparação e removidos
com o diretório temporário. Naquele estágio, a versão atual permanecia fora da coleção histórica até validação; a
validação e implementação finais estão registradas no encerramento acima.

---

# 1. OBJETIVO

Este documento define a arquitetura técnica oficial da aplicação
Auditor de Planilhas Excel.

A arquitetura deve atender aos requisitos definidos em:

docs/01_ESPECIFICACAO_FUNCIONAL.md

Os principais objetivos técnicos são:

- simplicidade;
- funcionamento confiável;
- separação de responsabilidades;
- operação somente leitura no SharePoint;
- processamento incremental;
- idempotência;
- rastreabilidade;
- possibilidade de testes sem SharePoint;
- ausência de dependência de Node.js;
- possibilidade futura de migração de SQLite para SQL Server.

A aplicação deverá evitar complexidade arquitetural desnecessária.

---

# 2. PRINCÍPIOS ARQUITETURAIS

A solução deverá seguir os seguintes princípios:

1. Python como plataforma principal.

2. SharePoint tratado exclusivamente como fonte de leitura.

3. Banco de auditoria tratado como fonte oficial da trilha consolidada.

4. Relatórios tratados como representações dos dados do banco.

5. Motor de comparação independente do SharePoint.

6. Fonte de versões intercambiável:
   - fonte local para desenvolvimento e testes;
   - fonte SharePoint para produção;
   - mecanismo concreto de aquisição desacoplado do motor de auditoria.

7. Processamento incremental por planilha.

8. Checkpoint associado à identidade técnica estável da planilha.

9. Operações críticas protegidas por transação.

10. Falha de processamento não poderá avançar checkpoint indevidamente.

11. Reexecução não poderá duplicar evidências.

12. Simplicidade deverá ser preferida a abstrações prematuras.

13. Microsoft Graph não constitui dependência arquitetural obrigatória.

14. A aquisição de versões SharePoint deverá utilizar somente mecanismos
    suportados e autorizados no ambiente corporativo.

15. Restrições de autenticação ou autorização não deverão ser contornadas
    por captura de cookies, tokens, senhas ou sessões de usuário.

16. Mecanismos alternativos de aquisição poderão ser implementados sem
    modificar o motor Excel, as regras de auditoria ou a persistência.

---

# 3. VISÃO GERAL

Arquitetura conceitual:

                    ┌─────────────────────┐
                    │      INTERFACE      │
                    │   Python / Desktop  │
                    └──────────┬──────────┘
                               │
                               ▼
                    ┌─────────────────────┐
                    │  SERVIÇO AUDITORIA  │
                    │     Orquestração    │
                    └──────────┬──────────┘
                               │
             ┌─────────────────┼──────────────────┐
             │                 │                  │
             ▼                 ▼                  ▼
    ┌────────────────┐ ┌───────────────┐ ┌────────────────┐
    │ FONTE DE DADOS │ │ MOTOR EXCEL   │ │ BANCO AUDITORIA│
    │                │ │               │ │                │
    │ │ VersionSource  │
     │ Local/SharePoint│ │ │ Reader + Diff │ │ SQLite / SQL   │
    └────────────────┘ └───────────────┘ └────────────────┘
                                                   │
                                                   ▼
                                        ┌──────────────────┐
                                        │ GERADOR RELATÓRIO│
                                        │      Excel       │
                                        └──────────────────┘

---

# 4. COMPONENTES PRINCIPAIS

A aplicação será dividida nos seguintes componentes:

1. Configuração
2. Interface
3. Serviço de auditoria
4. Fonte de dados
5. Mecanismo de aquisição SharePoint
6. Leitor Excel
7. Comparador
8. Persistência
9. Relatórios
10. Logging

Cada componente deverá possuir responsabilidade clara.

---

# 5. ESTRUTURA INICIAL DO PROJETO

Estrutura recomendada:

auditoria_excel/
│
├── main.py
├── requirements.txt
├── .env.example
├── .gitignore
├── README.md
│
├── app/
│   ├── __init__.py
│   │
│   ├── config.py
│   │
│   ├── database.py
│   │
│   ├── models.py
│   │
│   ├── exceptions.py
│   │
│   ├── logging_config.py
│   │
│   ├── audit_service.py
│   │
│   ├── report_service.py
│   │
│   ├── integrity.py
│   │
│   ├── excel/
│   │   ├── __init__.py
│   │   ├── reader.py
│   │   └── comparator.py
│   │
│   └── sources/
│       ├── __init__.py
│       ├── base.py
│       ├── local.py
│       └── sharepoint.py
│
├── tests/
│   ├── conftest.py
│   ├── test_reader.py
│   ├── test_comparator.py
│   ├── test_database.py
│   └── test_audit_service.py
│
├── data/
│   ├── database/
│   ├── temp/
│   └── reports/
│
├── logs/
│
└── docs/
    ├── 01_ESPECIFICACAO_FUNCIONAL.md
    ├── 02_ARQUITETURA.md
    ├── 03_PLANO_DE_DESENVOLVIMENTO.md
    └── 04_HISTORICO_IMPLEMENTACOES.md

A estrutura poderá sofrer pequenos ajustes durante a implementação,
desde que não altere os princípios definidos neste documento.

---

# 6. MAIN.PY

`main.py` será o ponto de entrada da aplicação.

Responsabilidades:

- inicializar configurações;
- inicializar logging;
- inicializar banco;
- iniciar interface ou modo de execução correspondente.

Não deverá conter:

- lógica de comparação;
- SQL complexo;
- chamadas detalhadas ao mecanismo de aquisição SharePoint;
- lógica de negócio da auditoria.

Essas responsabilidades pertencem aos componentes específicos.

---

# 7. CONFIGURAÇÃO

Arquivo:

app/config.py

Responsável por:

- caminhos;
- configurações do banco;
- configurações SharePoint;
- parâmetros de execução;
- diretório temporário;
- diretório de relatórios;
- nível de logging.

Credenciais NÃO poderão ser armazenadas diretamente no código.

Configurações sensíveis deverão utilizar variáveis de ambiente ou
mecanismo seguro equivalente.

Arquivo:

.env.example

poderá documentar os nomes das variáveis necessárias, mas nunca conter
credenciais reais.

---

# 8. FONTE DE DADOS

A aplicação deverá trabalhar com uma abstração simples de fonte.

O serviço de auditoria não deverá depender diretamente do Microsoft
Graph.

Deverão existir inicialmente duas implementações:

## 8.1 Fonte Local

Arquivo:

app/sources/local.py

Objetivo:

simular o comportamento necessário para desenvolvimento e testes.

Nos testes, as versões locais `.xlsx` deverão ser geradas programaticamente
com `openpyxl` em diretório fornecido por `pytest/tmp_path` (ou mecanismo
temporário equivalente), e descartadas ao final. Fixtures binárias não deverão
ser versionadas.

A fonte local deverá fornecer ao serviço de auditoria informações
equivalentes às necessárias na fonte SharePoint.

---

## 8.2 Fonte SharePoint

Arquivo:

## 8.2 Fonte SharePoint

Arquivo:

app/sources/sharepoint.py

Objetivo:

fornecer ao serviço de auditoria acesso normalizado às planilhas,
versões e metadados disponíveis no SharePoint, independentemente do
mecanismo concreto de aquisição autorizado no ambiente.

Responsabilidades:

- localizar planilhas elegíveis;
- estabelecer sua identidade técnica;
- listar ou descobrir versões;
- recuperar metadados disponíveis;
- recuperar conteúdo histórico para leitura;
- recuperar a versão atual quando necessário;
- normalizar os dados para o contrato interno da aplicação.

Todas as operações deverão ser exclusivamente de leitura.

A fonte SharePoint não deverá obrigar o restante da aplicação a conhecer
detalhes de Microsoft Graph, URLs históricas, autenticação ou qualquer
outro mecanismo concreto de aquisição.

Microsoft Graph poderá ser utilizado como implementação quando estiver
disponível e autorizado.

Caso Graph não esteja disponível, outro mecanismo somente poderá ser
adotado após validação técnica e de segurança.

Responsável por:

- autenticação;
- localizar site/biblioteca configurados;
- listar planilhas elegíveis;
- obter DriveItem ID;
- listar versões;
- recuperar metadados das versões;
- baixar conteúdo histórico para leitura;
- recuperar versão atual quando necessário.

Todas as operações deverão ser de leitura.

---

# 9. CONTRATO DA FONTE

A interface interna da fonte deverá ser pequena.

Conceitualmente:

listar_planilhas()

listar_versoes(id_planilha)

baixar_versao(id_planilha, id_versao)

obter_metadados(...)

O serviço de auditoria deverá conseguir operar com qualquer implementação
compatível com esse contrato sem alteração de sua lógica de negócio.

A identidade utilizada por `id_planilha` deverá representar a identidade
técnica estável definida pela fonte, não necessariamente um DriveItem ID.

Detalhes específicos do mecanismo de aquisição não deverão vazar para
AuditService, ExcelReader ou Comparator.

---

# 10. MODELO INTERNO DE VERSÃO

Cada versão deverá ser normalizada para uma estrutura interna.

Exemplo conceitual:

VersionInfo
- id
- numero
- data_hora
- autor
- comentario
- tamanho
- origem

IMPORTANTE:

`numero` é uma representação de exibição da versão.

Sempre que a API fornecer um identificador próprio da versão,
esse identificador deverá ser preservado.

Não depender exclusivamente de conversão numérica de strings como:

"0.99"
"1.00"
"1.10"

para identificar versões.

---

# 11. ORDENAÇÃO DE VERSÕES

A aplicação não deverá utilizar ordenação lexicográfica simples.

Exemplo incorreto:

"1.10" < "1.2"

dependendo da comparação textual.

A ordem deverá utilizar informação confiável fornecida pela fonte.

Preferências:

1. identificador/ordem oficial da API, quando apropriado;
2. data/hora + identificador;
3. parser controlado de versão, se necessário.

A estratégia definitiva deverá ser validada contra o comportamento real da fonte SharePoint e do mecanismo de aquisição adotado.

Quando a fonte fornecer uma ordem oficial ou identificador confiável, essa informação deverá prevalecer sobre inferências baseadas apenas no número exibido da versão.

---

# 12. DOWNLOAD TEMPORÁRIO

Versões históricas poderão precisar ser baixadas temporariamente.

Diretório:

data/temp/

Cada instância da fonte utiliza um subdiretório exclusivo, identificado por marcador
da aplicação. Nomes de XLSX são derivados por SHA-256 da identidade contextual da
planilha e do ID oficial da versão, evitando path traversal e colisões entre planilhas,
versões e execuções.

Os arquivos temporários não são evidência oficial. O motor os remove imediatamente
depois de construir o snapshot em memória, tanto em sucesso quanto em falha posterior;
a baseline permanece no snapshot somente durante a comparação incremental. O
encerramento da fonte remove seu subdiretório e a inicialização seguinte remove apenas
subdiretórios órfãos reconhecidos, com marcador válido e processo proprietário ausente.
Artefatos externos, diretórios sem marcador e workspaces de processos ativos não são
removidos.

Assim, a aplicação evita acúmulo durante auditorias extensas e recupera resíduos de
interrupção abrupta sem limpeza ampla do diretório configurado.

Nunca modificar o arquivo baixado.

---

# 13. LEITOR EXCEL

Arquivo:

app/excel/reader.py

Responsável por transformar um arquivo Excel em uma representação
comparável.

Biblioteca principal:

openpyxl

A leitura deverá preservar fórmulas:

data_only=False

O leitor não deverá salvar alterações no arquivo.

---

# 14. SNAPSHOT

Uma versão Excel deverá ser convertida em snapshot lógico.

Modelo conceitual:

Snapshot
└── Aba
    └── Célula
        ├── endereço
        └── valor

Exemplo:

{
    "Resultados": {
        "A1": "Data",
        "B1": "Resultado",
        "A2": "14/09/2026",
        "B2": 15.5
    }
}

O modelo definitivo poderá utilizar classes/dataclasses se isso melhorar
clareza sem criar complexidade desnecessária.

---

# 15. CÉLULAS VAZIAS

A implementação deverá definir de forma determinística o tratamento de:

- célula inexistente;
- célula existente com None;
- string vazia;
- fórmula;
- número zero;
- booleano False.

Não utilizar simplesmente:

if not valor

pois:

0

False

e:

""

possuem significados distintos.

As regras deverão ser cobertas por testes.

---

# 16. MOTOR DE COMPARAÇÃO

Arquivo:

app/excel/comparator.py

Entrada:

snapshot anterior
snapshot atual

Saída:

coleção de alterações.

Cada alteração deverá possuir:

- aba;
- endereço;
- tipo;
- valor anterior;
- valor novo.

Tipos:

ADD
DEL
MOD

O comparador não deverá:

- acessar banco;
- acessar SharePoint;
- gerar relatório;
- atualizar checkpoint.

Ele apenas compara estados.

---

# 17. DETERMINISMO

Para os mesmos dois snapshots, o comparador deverá sempre produzir o
mesmo resultado.

A ordenação das alterações deverá ser estável.

Preferencialmente:

1. aba;
2. linha;
3. coluna.

Isso facilita:

- testes;
- reprodução;
- geração de hash;
- análise humana.

---

# 18. SERVIÇO DE AUDITORIA

Arquivo:

app/audit_service.py

Será o núcleo de orquestração.

Responsabilidades:

- receber planilha selecionada;
- consultar checkpoint;
- consultar versões disponíveis;
- determinar versões pendentes;
- preservar a versão-base;
- baixar versões necessárias;
- criar snapshots;
- comparar versões;
- persistir resultados;
- atualizar checkpoint;
- registrar execução;
- registrar erros.

---

# 19. FLUXO DA PRIMEIRA AUDITORIA

Exemplo:

CQL028.xlsx

Versões:

0.84
0.85
0.86
0.87

Fluxo:

1. localizar planilha;
2. registrar/atualizar cadastro;
3. verificar ausência de checkpoint;
4. obter versões disponíveis;
5. ordenar versões;
6. selecionar histórico;
7. recuperar 0.84;
8. recuperar 0.85;
9. comparar 0.84 → 0.85;
10. persistir;
11. checkpoint = 0.85;
12. comparar 0.85 → 0.86;
13. persistir;
14. checkpoint = 0.86;
15. continuar;
16. checkpoint final = 0.87.

---

# 20. FLUXO INCREMENTAL

Situação:

Checkpoint:

0.99

Novas versões:

1.00
1.01
1.02

Fluxo correto:

recuperar versão-base 0.99

0.99 → 1.00
1.00 → 1.01
1.01 → 1.02

Checkpoint final:

1.02

A versão 0.99 não deverá ser eliminada antes da construção da primeira
comparação.

---

# 21. OTIMIZAÇÃO DE MEMÓRIA

O sistema não deverá carregar centenas de versões simultaneamente.

Preferência:

manter apenas os snapshots necessários à comparação atual.

Exemplo:

snapshot N
snapshot N+1

Após processar:

descartar snapshot N

reutilizar snapshot N+1 como anterior quando possível.

Fluxo:

SNAP 0.99
     +
SNAP 1.00
     ↓
COMPARA
     ↓
descarta 0.99

SNAP 1.00
     +
SNAP 1.01
     ↓
COMPARA

Isso reduz uso de memória.

---

# 22. BANCO DE DADOS

Banco inicial:

SQLite

Local sugerido:

data/database/auditoria.db

O banco será a fonte oficial da trilha consolidada.

---

# 23. MODELO DE DADOS

Modelo conceitual:

PLANILHA
    │
    ├──────── CHECKPOINT
    │
    ├──────── VERSAO_PROCESSADA
    │               │
    │               └──── ALTERACAO
    │
    └──────── EXECUCAO_AUDITORIA
                    │
                    └──── ERRO_PROCESSAMENTO

---

# 24. TABELA PLANILHA

Nome sugerido:

planilha

Campos mínimos:

id
drive_item_id
nome_atual
site_id
drive_id
caminho_sharepoint
data_cadastro
data_atualizacao

Regras:

- `drive_item_id` deverá ser único dentro do contexto técnico necessário;
- nome não será chave;
- alterações de nome deverão atualizar apenas informação descritiva.

IMPORTANTE:

A integração deverá confirmar se a identidade técnica precisa ser
composta por mais de um identificador, como drive/site + item ID.

A modelagem deverá preservar os identificadores necessários para evitar
colisões ou ambiguidades entre bibliotecas.

Os campos relacionados ao Microsoft Graph poderão permanecer na
modelagem existente por compatibilidade e uso futuro.

Entretanto, a arquitetura não deverá exigir que `drive_item_id`,
`site_id` ou `drive_id` estejam disponíveis em todos os mecanismos de
aquisição.

Caso a fonte adotada não forneça esses identificadores, deverá ser
definida uma identidade técnica alternativa estável antes da utilização
em produção.

Qualquer alteração de schema necessária para suportar essa identidade
deverá ser avaliada explicitamente, preservando os dados já existentes
e a compatibilidade com o motor de auditoria.

---

# 25. TABELA CHECKPOINT

Nome sugerido:

checkpoint

Campos mínimos:

id
planilha_id
versao_id
versao_numero
data_hora_versao
data_atualizacao

Regra:

uma planilha possui no máximo um checkpoint ativo.

O checkpoint representa a última versão cuja comparação necessária foi
concluída e persistida com sucesso.

---

# 26. TABELA VERSAO_PROCESSADA

Nome sugerido:

versao_processada

Campos mínimos:

id
planilha_id
versao_anterior_id
versao_anterior_numero
versao_atual_id
versao_atual_numero
data_hora_versao
autor
comentario
tamanho
quantidade_alteracoes
status
hash_origem
data_processamento
execucao_id

Status possíveis inicialmente:

PROCESSADA
SEM_ALTERACOES
ERRO

A implementação poderá ajustar os nomes, mantendo o significado.

---

# 27. TABELA ALTERACAO

Nome sugerido:

alteracao

Campos mínimos:

id
versao_processada_id
planilha_id
tipo
aba
endereco
valor_anterior
valor_novo

Tipos permitidos:

ADD
DEL
MOD

Deverá existir proteção contra duplicidade.

Uma estratégia possível é uma constraint lógica envolvendo:

planilha
versão anterior
versão atual
aba
endereço

A modelagem definitiva deverá evitar duplicação sem impedir registros
legítimos.

---

# 28. TABELA EXECUCAO_AUDITORIA

Nome sugerido:

execucao_auditoria

Campos mínimos:

id
codigo_execucao
planilha_id
inicio
fim
checkpoint_inicial
versao_final
versoes_processadas
alteracoes_encontradas
status
mensagem

Status sugeridos:

EM_EXECUCAO
CONCLUIDA
CONCLUIDA_SEM_NOVIDADES
FALHA

---

# 29. TABELA ERRO_PROCESSAMENTO

Nome sugerido:

erro_processamento

Campos mínimos:

id
execucao_id
planilha_id
versao_anterior
versao_atual
tipo_erro
mensagem
data_hora

Não armazenar:

- tokens;
- senhas;
- segredos;
- conteúdo sensível desnecessário.

---

# 30. RELACIONAMENTOS

Relação conceitual:

planilha
   1
   │
   ├──── 1 checkpoint
   │
   ├──── N versao_processada
   │          │
   │          └──── N alteracao
   │
   └──── N execucao_auditoria
               │
               └──── N erro_processamento

---

# 31. TRANSAÇÕES

A persistência de uma comparação deverá ser atômica sempre que possível.

Para:

0.99 → 1.00

a operação lógica deverá ser:

BEGIN

1. registrar versão processada;
2. registrar alterações;
3. atualizar checkpoint.

COMMIT

Se ocorrer falha:

ROLLBACK

Assim não poderá ocorrer situação onde:

checkpoint = 1.00

mas as alterações de:

0.99 → 1.00

não foram persistidas.

---

# 32. PROTEÇÃO CONTRA DUPLICIDADE

O sistema deverá possuir duas camadas:

1. lógica da aplicação;
2. constraints/índices do banco.

A aplicação deverá verificar se determinada comparação já foi
processada.

O banco deverá impedir inserção duplicada em caso de erro lógico ou
reexecução inesperada.

---

# 33. HASH E INTEGRIDADE

Arquivo sugerido:

app/integrity.py

Poderá calcular SHA-256 do conteúdo binário das versões baixadas.

Objetivo:

registrar uma impressão digital da versão efetivamente utilizada na
comparação.

Exemplo:

SHA256(arquivo_0.99.xlsx)

O hash NÃO substitui:

- controle de acesso;
- backup;
- assinatura digital;
- proteção do banco.

Sua função é auxiliar na verificação de integridade e
reprodutibilidade.

A implementação poderá ser introduzida somente quando seu uso estiver
claramente integrado ao fluxo.

---

# 34. RELATÓRIOS

Arquivo:

app/report_service.py

Biblioteca:

openpyxl

Não utilizar Node.js.

O relatório deverá ser produzido a partir do banco.

Nunca utilizar o relatório anterior como fonte primária para determinar
o checkpoint.

Fluxo:

BANCO
  │
  ▼
REPORT SERVICE
  │
  ▼
CQL028_Trilha_Auditoria.xlsx

---

# 35. DIRETÓRIO DE RELATÓRIOS

Local padrão:

data/reports/

Exemplo:

data/reports/
├── CQL028_Trilha_Auditoria.xlsx
├── CQL029_Trilha_Auditoria.xlsx
└── CQL030_Trilha_Auditoria.xlsx

O nome é apenas para conveniência humana.

A associação técnica continuará baseada no banco e nos identificadores
do SharePoint.

---

# 36. ATUALIZAÇÃO DO RELATÓRIO

Quando novas versões forem auditadas, a aplicação poderá regenerar o
relatório completo a partir do banco.

Exemplo:

Banco contém:

0.84 → ... → 0.99

Relatório:

CQL028_Trilha_Auditoria.xlsx

Nova auditoria:

0.99 → ... → 1.20

Após sucesso:

o banco passa a conter o histórico consolidado até 1.20.

O relatório poderá então ser regenerado contendo todo o histórico até
1.20.

Não é necessário utilizar o próprio Excel anterior como banco
incremental.

---

# 37. INTERFACE DA V1

A interface deverá ser simples.

A tecnologia definitiva deverá ser escolhida priorizando:

- execução em Python;
- ausência de Node.js;
- facilidade de distribuição;
- estabilidade;
- baixa complexidade.

Uma opção aceitável para desktop é Tkinter/ttk, por estar integrada ao
ecossistema Python.

Outra alternativa Python somente poderá ser utilizada se trouxer
benefício concreto.

Não introduzir servidor web apenas para produzir uma interface se isso
não for necessário.

---

# 38. FLUXO DA INTERFACE

Tela principal conceitual:

┌─────────────────────────────────────────────────┐
│        AUDITOR DE PLANILHAS SHAREPOINT          │
├─────────────────────────────────────────────────┤
│                                                 │
│ Planilha: [ CQL028.xlsx                 ▼ ]     │
│                                                 │
│ Status:                Já auditada              │
│ Última auditada:       0.99                     │
│ Última disponível:     1.20                     │
│ Versões pendentes:     21                       │
│                                                 │
│ [ CONTINUAR AUDITORIA ]                         │
│                                                 │
│ [ VISUALIZAR TRILHA ] [ GERAR RELATÓRIO ]       │
│                                                 │
└─────────────────────────────────────────────────┘

Na primeira auditoria:

[ AUDITAR HISTÓRICO ]

Em auditorias posteriores:

[ CONTINUAR AUDITORIA ]

---

# 39. PROCESSAMENTO EM SEGUNDO PLANO DA INTERFACE

Uma auditoria potencialmente longa não deverá congelar a interface.

Quando necessário, processamento poderá ocorrer em thread de trabalho
controlada.

Atualizações visuais deverão respeitar as regras da biblioteca de
interface utilizada.

Não introduzir arquitetura assíncrona complexa sem necessidade.

---

# 40. LOGGING

Arquivo:

app/logging_config.py

Diretório:

logs/

Registrar eventos técnicos relevantes:

- início da aplicação;
- início de auditoria;
- planilha selecionada;
- quantidade de versões;
- início/fim de comparação;
- checkpoint;
- conclusão;
- falhas.

Não registrar:

- senha;
- token;
- segredo;
- credenciais.

Logs não substituem a tabela oficial de execuções.

---

# 41. AUTENTICAÇÃO MICROSOFT

A estratégia de autenticação e acesso deverá ser compatível com os
mecanismos oficialmente suportados e autorizados no ambiente
corporativo.

Critérios:

- utilizar somente mecanismos suportados;
- respeitar as políticas corporativas de autenticação e autorização;
- não armazenar senha em código;
- não armazenar credenciais de usuário em texto;
- aplicar menor privilégio;
- operar exclusivamente em leitura;
- não extrair cookies ou tokens de sessões existentes;
- não utilizar sessões autenticadas de navegador ou Office como forma
  de contornar restrições programáticas;
- permitir diagnóstico claro quando determinado mecanismo estiver
  indisponível.

A existência de acesso do usuário através do navegador ou Microsoft
Excel não implica automaticamente autorização ou capacidade de acesso
programático.

Nenhum método obsoleto, inseguro ou destinado a contornar controles
corporativos deverá ser implementado.

---

# 42. MICROSOFT GRAPH

A integração deverá validar experimentalmente, para cada mecanismo
candidato:

1. identificação da planilha;
2. identidade técnica estável;
3. descoberta/listagem de versões;
4. identificadores das versões;
5. metadados;
6. autor da versão;
7. data/hora;
8. comentário, quando disponível;
9. recuperação do conteúdo histórico;
10. versões principais;
11. versões secundárias;
12. ordenação confiável;
13. comportamento da autenticação;
14. garantia de operação exclusivamente em leitura.

Microsoft Graph permanece candidato preferencial quando disponível e
autorizado, mas não é requisito arquitetural exclusivo.

Nenhum comportamento deverá ser presumido sem validação.

---

# 43. LIMITAÇÃO CRÍTICA A VALIDAR

# 43. RESTRIÇÃO CORPORATIVA E INVESTIGAÇÃO F4

Durante a F4 foi constatado que o ambiente corporativo não disponibiliza
ao projeto os dados e autorizações necessários para utilização da
integração Microsoft Graph originalmente prevista.

Também foi validado que uma versão histórica real pode ser acessada pelo
usuário autenticado através de endereço SharePoint contendo
`_vti_history`.

No teste controlado realizado com CQL028.xlsx, o endereço histórico
correspondente à versão 0.97 abriu corretamente essa versão através da
sessão autenticada do usuário.

Entretanto, uma requisição Python HTTP não autenticada ao mesmo recurso
retornou HTTP 403.

Portanto:

- a URL histórica não deverá ser considerada acesso programático
  automaticamente disponível;
- não deverão ser extraídos cookies ou tokens do navegador/Office;
- não deverão ser utilizados mecanismos de contorno da autenticação;
- Microsoft Graph deverá permanecer disponível arquiteturalmente para
  utilização futura caso seja autorizado;
- a F4 deverá investigar alternativas suportadas e autorizadas de
  aquisição automatizada;
- importação manual deverá ser considerada contingência, e não solução
  automática oficial, enquanto a investigação técnica não estiver
  concluída.

Nenhuma mudança estrutural adicional deverá ser implementada antes da
validação da alternativa escolhida.

---

# 44. PERFORMANCE

O sistema deverá ser projetado para crescer, mas sem otimização
prematura.

O requisito de capacidade de produção considera aproximadamente 2.000 planilhas,
algumas com mais de 3.000 versões, e potencialmente milhões de alterações. Um único
banco SQLite canônico deverá atender às várias planilhas, mantendo por identidade os
respectivos históricos, checkpoints, execuções e erros. A adoção de outro banco exige
evidência obtida por medição; não é uma decisão preventiva.

A validação de capacidade deverá medir índices e consultas, crescimento do banco,
memória, relatórios, processamento inicial de históricos extensos, processamento
incremental e lotes de muitas planilhas. Os gargalos observados, e não apenas o volume
estimado, orientarão qualquer otimização.

Medidas iniciais:

- processamento incremental;
- apenas duas versões em memória quando possível;
- transações;
- índices no banco;
- arquivos temporários controlados;
- não recalcular histórico consolidado.

Somente após medição real deverão ser introduzidas otimizações
adicionais.

---

# 45. SQL SERVER FUTURO

A V1 utilizará SQLite.

Entretanto, regras de negócio não deverão depender de características
exclusivas do SQLite quando isso puder ser evitado.

O acesso ao banco deverá ficar concentrado na camada de persistência.

Migração futura:

SQLite
   ↓
SQL Server

não deverá exigir reescrita de:

- reader;
- comparator;
- SharePoint source;
- regras de auditoria.

---

# 46. BACKUP

O banco:

data/database/auditoria.db

é um ativo crítico.

Em produção deverá existir estratégia de:

- backup;
- controle de acesso;
- retenção;
- restauração.

A definição da infraestrutura de backup poderá ocorrer na fase de
preparação para produção.

---

# 47. TESTABILIDADE

O desenho deverá permitir:

LocalSource
     │
     ▼
AuditService
     │
     ├── ExcelReader
     ├── Comparator
     └── Database

sem qualquer conexão com SharePoint.

Isso permitirá testar a maior parte da aplicação localmente.

---

# 48. TESTES AUTOMATIZADOS

Utilizar:

pytest

Cobertura prioritária:

- comparação;
- ADD;
- DEL;
- MOD;
- fórmulas;
- vazio versus zero;
- múltiplas abas;
- checkpoint;
- idempotência;
- transações;
- retomada após falha.

Não perseguir porcentagem de cobertura apenas como métrica.

Priorizar cenários críticos.

---

# 49. DEPENDÊNCIAS

Manter `requirements.txt` pequeno.

Dependências somente deverão ser adicionadas quando utilizadas.

Base esperada:

openpyxl
pytest

Dependências relacionadas à aquisição SharePoint deverão ser adicionadas
somente após definição e validação do mecanismo autorizado.

`requests`, bibliotecas Microsoft ou MSAL poderão ser utilizados quando
forem efetivamente necessários ao mecanismo escolhido.

Nenhuma biblioteca deverá ser adicionada apenas com base na arquitetura
original do Microsoft Graph.

---

# 50. SEGURANÇA

Princípios mínimos:

- menor privilégio;
- SharePoint read-only;
- nenhum segredo no Git;
- `.env` ignorado;
- logs sem segredos;
- validação de caminhos;
- arquivos temporários controlados;
- banco protegido no ambiente de produção.

---

# 51. TRATAMENTO DE ERROS

Exceções técnicas deverão ser convertidas em mensagens compreensíveis
quando apresentadas ao usuário.

Exemplo:

Em vez de apenas:

HTTP 403

apresentar:

"Não foi possível acessar programaticamente o histórico de versões da
planilha através do mecanismo configurado. Verifique a configuração,
autorização e compatibilidade do método de acesso ao SharePoint."

O detalhe técnico poderá permanecer no log.

---

# 52. COMPORTAMENTO EM FALHA

Falha em determinada comparação não deverá destruir o histórico
consolidado anteriormente.

Exemplo:

0.99 → 1.00 OK
1.00 → 1.01 FALHA

Estado final:

histórico até 1.00 preservado
checkpoint = 1.00

A aplicação deverá permitir nova tentativa posteriormente.

---

# 53. EXCLUSÃO DE DADOS

A V1 não deverá possuir botão genérico de:

"Excluir histórico"

ou:

"Limpar auditoria"

na interface normal.

Qualquer mecanismo futuro de reprocessamento ou manutenção deverá ser
explicitamente projetado para evitar destruição acidental de evidência.

---

# 54. ESCOPO DA V1

Incluído:

- fonte local;
- SharePoint;
- seleção de planilha;
- histórico de versões;
- comparação;
- checkpoint;
- persistência;
- trilha consolidada;
- relatório Excel;
- interface simples;
- logs;
- testes;
- tratamento de erros.

Não incluído inicialmente:

- Power BI;
- dashboards complexos;
- aplicação web pública;
- múltiplos servidores;
- microsserviços;
- Kubernetes;
- Node.js;
- IA;
- análise semântica das alterações;
- interpretação automática do motivo de uma alteração.

---

# 55. FLUXO COMPLETO

Fluxo esperado:

USUÁRIO
   │
   ▼
SELECIONA CQL028.xlsx
   │
   ▼
IDENTIFICA PLANILHA E IDENTIDADE TÉCNICA
   │
   ▼
CONSULTA BANCO
   │
   ├── sem checkpoint ──► auditoria histórica
   │
   └── com checkpoint ──► auditoria incremental
                              │
                              ▼
                    CONSULTA FONTE SHAREPOINT
                              │
                              ▼
                      LISTA VERSÕES
                              │
                              ▼
                     DETERMINA PENDÊNCIAS
                              │
                              ▼
                    ADQUIRE N E N+1
                              │
                              ▼
                      CRIA SNAPSHOTS
                              │
                              ▼
                         COMPARA
                              │
                              ▼
                      ADD / DEL / MOD
                              │
                              ▼
                       TRANSAÇÃO DB
                         /       \
                        /         \
                ALTERAÇÕES     CHECKPOINT
                              │
                              ▼
                     PRÓXIMA VERSÃO
                              │
                              ▼
                        CONCLUSÃO
                              │
                              ▼
                     GERAR RELATÓRIO

---

# 56. REGRA DE ALTERAÇÃO DA ARQUITETURA

O Codex poderá propor alterações nesta arquitetura quando identificar:

- limitação técnica comprovada;
- problema de segurança;
- melhoria significativa de simplicidade;
- requisito incompatível.

Entretanto, alterações estruturais relevantes não deverão ser
implementadas silenciosamente.

Deverá informar:

1. problema encontrado;
2. impacto;
3. solução proposta;
4. arquivos afetados.

E aguardar decisão quando houver mudança relevante de escopo ou
arquitetura.

---

# 57. PRIORIDADE EM CASO DE CONFLITO

Em caso de conflito entre conveniência técnica e integridade da
auditoria, prevalece:

INTEGRIDADE DA AUDITORIA.

Em seguida:

SEGURANÇA

CORREÇÃO

RASTREABILIDADE

SIMPLICIDADE

PERFORMANCE

APARÊNCIA

---

# 58. REFERÊNCIAS INTERNAS

Este documento deve ser utilizado em conjunto com:

01_ESPECIFICACAO_FUNCIONAL.md
03_PLANO_DE_DESENVOLVIMENTO.md
04_HISTORICO_IMPLEMENTACOES.md

A especificação funcional define:

O QUE o sistema deve fazer.

Este documento define:

COMO a solução será estruturada.

O plano de desenvolvimento define:

EM QUE ORDEM será implementada.

O histórico registra:

O QUE realmente foi implementado.

---

# 59. CRITÉRIO ARQUITETURAL FINAL

A arquitetura será considerada adequada quando o AuditService puder
operar com diferentes implementações compatíveis de VersionSource sem
reescrever:

- motor de comparação;
- regras de checkpoint;
- persistência;
- trilha consolidada;
- geração de relatório.

A substituição entre:

LocalSource

e:

SharePointSource

não deverá exigir alteração das regras de negócio.

Da mesma forma, a substituição do mecanismo interno de aquisição da
SharePointSource — por exemplo Microsoft Graph ou outro mecanismo
suportado e autorizado — não deverá exigir reescrita do motor de
auditoria.

---

FIM DO DOCUMENTO
