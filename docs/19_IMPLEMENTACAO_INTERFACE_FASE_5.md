# Fase 5 — consolidação da interface e do tempo compartilhado

## Interface

Os cartões variavam porque o `LabelFrame` aceitava a geometria requisitada por
labels cujo conteúdo incluía etapa, média e ETA. Agora cada cartão tem 82 pixels,
propagação de geometria desativada, duas colunas com o mesmo grupo `uniform` e
quatro regiões invariáveis: título, barra, etapa resumida e tempo decorrido. A
barra é criada uma vez e apenas sua `DoubleVar` é atualizada.

O slot não apresenta média ou previsão. Mensagens técnicas permanecem no log;
textos visuais são normalizados e limitados. O tempo da tarefa sempre usa
`HH:MM:SS` e volta a zero quando o evento da nova tarefa chega.

## Definições do bloco global

* **Média recente:** média robusta da janela de tarefas completas
  (`TOTAL_TASK`) observadas por todos os slots no `SharedExecutionTimingModel`.
* **Taxa recente:** checkpoints oficialmente promovidos por minuto. Não é a soma
  da velocidade dos workers e o intervalo de pausa é descontado.
* **Estimativa restante:** após três commits, versões restantes divididas pela
  taxa oficial; durante o bootstrap, tarefas restantes multiplicadas pela média
  compartilhada e divididas pelos slots ativos.
* **Slots ativos:** slots com trabalho no instante da telemetria sobre a
  quantidade configurada.

Antes de existirem dados, os três valores mostram `calculando...`. Depois da
primeira estimativa válida, a tela mantém o último valor durante transições sem
nova amostra. Na conclusão, o ETA é `00:00:00` e os valores finais não são
apagados.

## Ciclo de vida

Há uma instância lógica do modelo por identidade técnica da planilha durante a
sessão da interface. Trocas de tarefa e atualizações visuais não a recriam.
Pause congela o relógio sem apagar amostras; Resume retoma o mesmo relógio.
Stop e encerramento congelam o relógio, e Continue da mesma planilha preserva
amostras e commits removendo o intervalo parado. Outra planilha recebe um modelo
isolado e não herda estatísticas.

O cálculo temporal não foi movido para a camada visual: cada evento de slot
continua obtendo seu percentual do `estimate_task`, com pesos aprendidos por
etapa e cauda assintótica. Ocultar média e ETA no cartão não altera a progressão.

## Verificação visual automatizada

O teste de geometria parametriza 1, 2, 5 e 8 slots numa janela 1366×768 e percorre
AGUARDANDO, BAIXANDO, VALIDANDO, LENDO XLSX, COMPARANDO, STAGED, CHECKPOINT
CONFIRMADO, PAUSADO e ERRO, inclusive com mensagem longa. Ele compara largura e
altura após cada atualização e exige altura de 82 pixels. Em ambiente sem
servidor gráfico, o teste é explicitamente ignorado; deve ser executado num host
Tk/Windows ou X11 para validar a geometria real e capturar a tela.
