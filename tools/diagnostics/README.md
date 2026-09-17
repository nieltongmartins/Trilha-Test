# Diagnóstico isolado de SIGINT no Windows

Estes programas **não importam módulos da aplicação** e não devem ser usados em
produção. Cada execução grava horário, PID do Python, thread, sinal, estágio e,
quando criado, PID do EdgeDriver em `tools/diagnostics/teste-*-<pid>.log`. No
Windows também é registrada a lista de processos ligados ao console atual via
`GetConsoleProcessList`; essa lista mostra compartilhamento do console, mas não
identifica sozinha quem originou um evento.

O observador de SIGINT existe somente nestes probes. Ele registra o evento e
encaminha ao handler que já estava instalado, portanto Ctrl+C continua gerando
`KeyboardInterrupt`; SIGINT não é ignorado. O bloco `finally` chama `quit()`
apenas no WebDriver criado pelo próprio probe e então a interrupção é
repropagada.

## Ordem e comandos no Windows real

Execute em **Prompt de Comando ou PowerShell**, a partir da raiz, com o mesmo
Python 3.14 e a mesma sessão de console usados pela aplicação. Feche a janela do
Tk ou aguarde o período indicado. Não execute `main.py` nesta etapa.

```powershell
py -3.14 tools/diagnostics/test_a_tk_only.py
py -3.14 tools/diagnostics/test_b_tk_thread.py
py -3.14 tools/diagnostics/test_c_edge_only.py --hold-seconds 30
py -3.14 tools/diagnostics/test_d_tk_edge.py --hold-seconds 30
```

O teste D espera três segundos entre `webdriver.Edge()` e
`driver.get("about:blank")`, e registra os dois estágios separadamente. Execute o
teste E somente se A–D não localizarem o evento:

```powershell
py -3.14 tools/diagnostics/test_e_tk_edge_sharepoint.py "https://EMPRESA.sharepoint.com/sites/SITE" --hold-seconds 30
```

Opcionalmente, repita C/D com o executável homologado para distinguir Selenium
Manager da inicialização do EdgeDriver:

```powershell
py -3.14 tools/diagnostics/test_c_edge_only.py --driver-path "C:\caminho\msedgedriver.exe"
py -3.14 tools/diagnostics/test_d_tk_edge.py --driver-path "C:\caminho\msedgedriver.exe"
```

Envie os logs produzidos pelos testes, começando pelo C. Eles não registram a
URL completa do SharePoint no teste E (somente esquema e host).

## Limites e interpretação

* A: controla Python + Tk sem thread de aplicação.
* B: adiciona uma thread comum, ainda sem Selenium.
* C: responde primeiro se `webdriver.Edge()`/`about:blank` provoca o evento sem
  Tk.
* D: verifica a combinação Tk + Selenium e separa criação e navegação.
* E: acrescenta a URL corporativa somente depois dos controles.
* `GetConsoleProcessList` e a árvore de PIDs são evidência de console/processos,
  não prova de autoria do SIGINT. O handler Python recebe o número do sinal, não
  o PID emissor.
* Não foi adicionada `CREATE_NEW_PROCESS_GROUP`: fazê-lo antes de comprovar
  propagação de evento seria uma mudança especulativa no modo suportado pelo
  Selenium de iniciar seu `Service`.

## Teste mínimo do fluxo do projeto

Depois dos controles acima, este teste usa a implementação real de
`BrowserSharePointSource`, mas deliberadamente não cria banco, `AuditService`,
relatório, descoberta recursiva nem auditoria. Ele mantém apenas um `Tk`, uma
worker daemon, a abertura do Edge, a navegação para `site_url` e a validação da
autenticação:

```powershell
py -3.14 -m tools.diagnostics.test_sharepoint_source_tk `
  "https://EMPRESA.sharepoint.com/sites/SITE" `
  --scope "/sites/SITE/Documentos"
```

Após concluir o login/MFA, a mensagem `SharePoint autenticado` deve aparecer e
a janela Tk deve continuar aberta. Este teste não instala handlers de sinais e
não fecha a janela automaticamente.
