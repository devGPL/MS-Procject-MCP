# MS Project MCP Server

Controla o Microsoft Project via automação COM através do Model Context Protocol (MCP).

99 ferramentas para ler, editar e analisar cronogramas direto do seu assistente de IA — sem sair do chat.

<!-- mcp-name: io.github.devGPL/msproject-mcp -->

---

## Requisitos

| Item | Detalhe |
|------|---------|
| SO | **Windows** (a automação COM não existe em macOS/Linux) |
| Microsoft Project | Instalado e **em execução** (testado no MS Project 16.0) |
| Python | 3.10 ou superior |
| Arquitetura | x64 e ARM64, sem diferença de instalação — nenhuma dependência precisa de compilador |

```bash
pip install mcp pywin32 python-dateutil
```

> `python-dateutil` só é necessário para `add_recurring_task`.

---

## Instalação

```bash
git clone https://github.com/devGPL/MS-Procject-MCP
cd MS-Procject-MCP
pip install -e .
```

Isso instala as dependências declaradas (`mcp`, `pywin32` no Windows, `python-dateutil`) e cria o comando `msproject-mcp`. Teste que ele sobe:

```bash
msproject-mcp
```

Ele fica aguardando no stdin — é o transporte do MCP. `Ctrl+C` encerra. Se aparecerem as duas linhas de diagnóstico em stderr, está funcionando.

`python server.py` continua funcionando de forma idêntica, caso prefira não instalar.

Para ligar o caminho rápido de leitura (leia [Dois caminhos de leitura](#dois-caminhos-de-leitura) antes — ele não funciona em Windows ARM64):

```bash
pip install -e ".[fast]"
```

> **Antes de usar qualquer ferramenta**: o MS Project precisa estar aberto com um arquivo carregado. O servidor se conecta a uma instância já em execução; ele não abre o Project sozinho.

> **Por que `pip install -e .` e não `pip install mcp pywin32`**: o erro mais comum de configuração é instalar as dependências num interpretador e apontar o cliente MCP para outro. Com o pacote instalado, o `msproject-mcp` que o cliente executa é o do ambiente que tem as bibliotecas — e some o caminho absoluto do JSON de configuração.

---

## Configuração — Claude

### Claude Desktop

Edite `claude_desktop_config.json`:

- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "msproject": {
      "command": "msproject-mcp",
      "args": []
    }
  }
}
```

Reinicie o Claude Desktop. O servidor aparece no ícone de ferramentas (🔨) do chat.

### Claude Code

Via CLI (recomendado — grava a config pra você):

```bash
claude mcp add msproject --scope user -- msproject-mcp
```

Escopos disponíveis:

| Escopo | Onde grava | Quando usar |
|--------|-----------|-------------|
| `user` | Config global do usuário | Disponível em todos os projetos |
| `project` | `.mcp.json` no repositório | Compartilhado com o time via Git |
| `local` | Só a sessão/pasta atual | Testes pontuais |

Conferir o registro:

```bash
claude mcp list
```

Alternativa manual — crie `.mcp.json` na raiz do projeto:

```json
{
  "mcpServers": {
    "msproject": {
      "command": "msproject-mcp",
      "args": []
    }
  }
}
```

---

## Configuração — Codex

### Via CLI

```bash
codex mcp add msproject -- msproject-mcp
```

Conferir:

```bash
codex mcp list
```

### Via arquivo

Edite `~/.codex/config.toml` (no Windows: `%USERPROFILE%\.codex\config.toml`):

```toml
[mcp_servers.msproject]
command = "msproject-mcp"
args = []
```

> **`msproject-mcp` não encontrado**: o comando fica no diretório `Scripts` (Windows) ou `bin` do ambiente onde você rodou `pip install -e .`. Se esse diretório não está no `PATH` do cliente, use o caminho absoluto do executável em `command`, ou instale dentro de um venv e aponte para o `msproject-mcp` dele.

Sem instalar o pacote, aponte para o `server.py` diretamente — aqui as barras invertidas precisam de escape `\\`, ou use string literal com aspas simples:

```toml
[mcp_servers.msproject]
command = "C:\\Python312\\python.exe"
args = ["C:\\caminho\\para\\MS-Procject-MCP\\server.py"]
```

---

## MS Project numa VM (Parallels, VMware, Hyper-V)

Se o Claude roda no macOS e o MS Project numa VM Windows, **não** dá para usar o padrão de SSH que funciona com outros MCPs.

O motivo é específico deste servidor. `GetActiveObject` lê a *Running Object Table* do Windows, que é **por sessão de logon**. Um login SSH cai na sessão 0; o MS Project aberto na área de trabalho está na sessão 1. Da sessão 0 ele é invisível, e a tentativa falha com `MK_E_UNAVAILABLE (0x800401E3)` mesmo com o programa aberto.

Servidores que conversam por porta TCP com o aplicativo alvo atravessam sessão sem problema — este precisa anexar via COM, e COM não atravessa.

A solução é inverter o que cruza a fronteira: em vez do COM atravessar a sessão, o **HTTP atravessa a máquina**.

### Na VM Windows

Instale e rode **a partir da área de trabalho** — um terminal aberto na VM, um atalho, ou a pasta Inicializar. Nunca por SSH, e **nunca como administrador**:

```bash
msproject-mcp --transport streamable-http --host 0.0.0.0 --port 8765
```

Libere a porta no firewall, uma vez só (PowerShell como administrador):

```powershell
New-NetFirewallRule -DisplayName "MS Project MCP" -Direction Inbound -LocalPort 8765 -Protocol TCP -Action Allow
```

### Subir junto com o Windows

Para não depender de abrir um terminal a cada login, crie `C:\msproject-mcp\start-server.cmd`:

```bat
@echo off
title MS Project MCP Server
set LOG=%LOCALAPPDATA%\msproject-mcp.log
echo ===== iniciado em %DATE% %TIME% ===== >> "%LOG%"
"%LOCALAPPDATA%\Programs\Python\Python312-arm64\Scripts\msproject-mcp.exe" --transport streamable-http --host 0.0.0.0 --port 8765 >> "%LOG%" 2>&1
echo ===== encerrado em %DATE% %TIME% (codigo %ERRORLEVEL%) ===== >> "%LOG%"
```

E um `.vbs` na pasta Inicializar (`Win+R` → `shell:startup`) para lançá-lo minimizado:

```vbs
CreateObject("WScript.Shell").Run """C:\msproject-mcp\start-server.cmd""", 7, False
```

A pasta Inicializar roda na sessão interativa e sem elevação, que é exatamente o que o COM exige. Um Agendador de Tarefas configurado com "executar com privilégios mais altos" **quebraria** por causa do nível de integridade.

O log em `%LOCALAPPDATA%\msproject-mcp.log` é onde procurar quando o servidor não sobe — a janela minimizada some sem deixar rastro se o processo morrer.

### No Mac

```bash
claude mcp add --transport http msproject http://10.211.55.5:8765/mcp
```

Troque o IP pelo da sua VM (`ipconfig` no Windows). No Claude Desktop, use `"url": "http://10.211.55.5:8765/mcp"` em vez de `command`/`args`.

> **Não eleve o terminal.** A Running Object Table também é separada por nível de integridade: um servidor iniciado como administrador roda em *High* e não enxerga o MS Project aberto normalmente, em *Medium*. Mesma máquina, mesma sessão, mesmo usuário — e o erro é idêntico ao de "não está rodando". A regra de firewall precisa de admin uma única vez; o servidor, nunca.

> **Segurança**: `--host 0.0.0.0` aceita conexões de qualquer interface e o servidor não tem autenticação própria. Use apenas em rede host-only do Parallels/VMware. Numa rede compartilhada com terceiros, qualquer um que alcance a porta controla o seu MS Project.

---

## Verificando a conexão

Peça ao assistente para rodar `health_check`. A resposta traz a versão do MS Project e o status do arquivo aberto. Se falhar:

| Erro | Causa | Solução |
|------|-------|---------|
| `Could not attach to MS Project` com o MS Project claramente aberto | Servidor iniciado como administrador. A Running Object Table é filtrada por nível de integridade: um processo *High* não enxerga um *Medium* | Suba o servidor num PowerShell **normal**. Rodar como admin piora, não ajuda |
| `MS Project is not running` | Nenhuma instância ativa | Abra o MS Project |
| `No project file is open` | Project aberto sem arquivo | Abra ou crie um `.mpp` |
| Servidor não aparece no cliente | Caminho errado ou JSON/TOML inválido | Confira o caminho absoluto e reinicie o cliente |
| `ModuleNotFoundError: win32com` | `pywin32` ausente no Python usado | Instale no mesmo interpretador que está no `command` |
| `Failed building wheel for cryptography` / `linker link.exe not found` | Versão do `mcp` fora do teto declarado, puxando `cryptography`, que não tem wheel para Windows ARM | `pip install -e .` a partir deste repositório — o teto `mcp<1.20` evita a dependência |

---

## Ferramentas

Organizadas por módulo. Cada arquivo carrega no seu cabeçalho a regra de fronteira que decide o que entra nele.

### Projeto, multiprojeto e intercâmbio (17)
`msp_projects.py` — ciclo de vida do arquivo .mpp, navegação entre projetos abertos, import/export e a sonda de conectividade

`open_project` · `new_project` · `get_project_info` · `set_project_properties` · `save_project` · `save_project_as` · `close_project` · `import_xml` · `export_xml` · `list_projects` · `switch_project` · `cross_project_link` · `undo_last` · `insert_subproject` · `snapshot_to_json` · `health_check` · `snapshot_diff`

### Leitura de tarefas (17)
`msp_tasks_read.py` — consultas, filtros, agregações e exportações — nada aqui escreve no projeto

`get_tasks` · `get_task` · `get_tasks_by_rag` · `get_overdue_tasks` · `get_tasks_by_resource` · `search_tasks` · `get_progress_summary` · `get_wbs_structure` · `filter_tasks` · `group_tasks_by` · `get_milestone_report` · `get_progress_by_wbs` · `export_csv` · `apply_filter` · `get_constraints` · `get_actual_work` · `get_timephased_data`

### Escrita de tarefas (19)
`msp_tasks_write.py` — criação, atualização, exclusão, modo de agendamento e edições em lote

`update_task` · `bulk_update_rag` · `bulk_update_tasks` · `add_task` · `bulk_add_tasks` · `delete_task` · `set_task_mode` · `bulk_set_task_mode` · `set_constraint` · `clear_estimated_flags` · `indent_task` · `set_deadline` · `set_task_active` · `dry_run_bulk_update` · `move_task` · `copy_task_structure` · `bulk_set_deadlines` · `set_task_hyperlink` · `add_recurring_task`

### Dependências (5)
`msp_dependencies.py` — a rede de precedência entre tarefas

`add_predecessor` · `bulk_add_predecessors` · `remove_predecessor` · `get_task_dependencies` · `get_dependency_chain`

### Recursos (11)
`msp_resources.py` — pool, alocações, disponibilidade e tabelas de custo

`get_resources` · `add_resource` · `assign_resource` · `get_resource_workload` · `bulk_assign_resources` · `remove_resource_assignment` · `update_resource` · `delete_resource` · `get_resource_availability` · `get_resource_rate_tables` · `set_resource_rate_table`

### Calendários (10)
`msp_calendars.py` — calendários base, exceções e horas de trabalho

`get_calendars` · `set_calendar_exception` · `set_project_calendar` · `set_task_calendar` · `create_calendar` · `list_calendar_exceptions` · `delete_calendar` · `delete_calendar_exception` · `set_resource_calendar` · `set_working_hours`

### Cronograma (11)
`msp_schedule.py` — o que é calculado sobre a rede de tarefas: caminho crítico, folga, nivelamento, avanço de status

`get_critical_path` · `get_schedule_analysis` · `validate_schedule` · `level_resources` · `find_available_slack` · `calculate_project` · `update_project` · `reschedule_incomplete_work` · `get_critical_path_sequence` · `get_critical_tasks_for_period` · `what_if_delay`

### Linha de base e custo (6)
`msp_baselines_costs.py` — planejado versus realizado: baselines, valor agregado, custo e variação

`save_baseline` · `clear_baseline` · `get_earned_value` · `compare_baselines` · `get_cost_summary` · `get_variance_report`

### Campos personalizados (3)
`msp_customfields.py` — os slots Text/Number/Date/Flag/Duration

`rename_custom_fields` · `update_custom_fields` · `get_custom_field_values`

**99 ferramentas** em 9 módulos.

---

## Dados retornados por tarefa

Toda consulta (`get_task`, `get_tasks`, etc.) devolve 35+ campos por tarefa, entre eles:

| Campo | Descrição |
|-------|-----------|
| `actual_start` / `actual_finish` | Início/término reais |
| `remaining_duration_days` | Trabalho restante |
| `total_slack_days` / `free_slack_days` | Folga total e livre |
| `deadline` | Prazo indicativo |
| `priority` | Prioridade de nivelamento (0–1000) |
| `constraint_type` / `constraint_date` | Restrição de agendamento |
| `manual` | Agendamento manual vs automático |
| `type` | FixedUnits / FixedDuration / FixedWork |
| `hyperlink` / `hyperlink_text` | Hyperlink da tarefa |

---

## Dois caminhos de leitura

Quatorze ferramentas de leitura respondem a partir do **arquivo `.mpp` salvo** quando conseguem, e caem no COM quando não conseguem. Toda resposta diz qual dos dois respondeu, no bloco `source`.

O motivo é medido. Numa agenda real de 8.429 tarefas:

| | |
|---|---|
| Obter um objeto tarefa pelo COM | 2,635 ms |
| Ler uma propriedade dele | 0,240 ms |
| Enumeração como fatia de uma varredura | 83% |
| Piso para tocar todas as tarefas pelo COM | ~20 s |
| Ler as mesmas 8.429 do arquivo (`mpxj`) | ~1,6 s |

Enumerar é o caro; otimizar a leitura de propriedade ataca 17% do problema.

**Quem usa:** `get_tasks` · `get_task` · `search_tasks` · `get_tasks_by_rag` · `get_overdue_tasks` · `get_tasks_by_resource` · `get_progress_summary` · `get_wbs_structure` · `filter_tasks` · `group_tasks_by` · `get_progress_by_wbs` · `get_constraints` · `export_csv` · `get_critical_path` · `get_schedule_analysis` · `find_available_slack`

As demais continuam só no COM porque precisam de campos que o arquivo não entrega por essa via: linha de base (`get_milestone_report`, `get_variance_report`), trabalho (`get_actual_work`), séries timephased, objetos de dependência (`get_critical_path_sequence`, `what_if_delay`) ou a própria janela (`apply_filter`).

**O que custa.** O caminho rápido lê o arquivo salvo, não o aplicativo em execução — e essa é uma questão de correção, não de velocidade:

- Edições feitas no MS Project e ainda não salvas são invisíveis para ele.
- Uma escrita deste servidor que não persistisse ficaria invisível para a leitura seguinte. É por isso que toda ferramenta que muta salva.

A decisão é tomada **a cada chamada**, nunca uma vez na inicialização. Quando o MS Project reporta alterações não salvas, a resposta traz `warning` dizendo de quando é o arquivo que respondeu. Se o caminho rápido levantar exceção, a resposta vem pelo COM com `action_required` — falha silenciosa aqui é como uma mudança que não entrega nada continua parecendo que funciona.

**Onde não funciona:** `jpype` não publica wheel para Windows em ARM, e o `mpxj` roda na JVM (Java 9+). Nessas máquinas o módulo se declara indisponível, tudo cai no COM e o `source` diz o motivo.

**Ordem de inicialização.** `jpype.startJVM()` precisa rodar **antes** de qualquer inicialização de COM na thread: com a STA já de pé, a chamada bloqueia por ~60 segundos — exatamente o timeout de tool call de clientes MCP comuns, que matam e recriam o servidor, e assim o cache do parse nunca sobrevive à segunda chamada. Por isso o servidor sobe a JVM ansiosamente no `main()`, antes da primeira ferramenta tocar o COM (~1 s nessa ordem). `MSPROJECT_MCP_LAZY_JVM=1` desliga o start ansioso e volta ao comportamento antigo, se precisar.

**Pegadinhas do Java em clientes MCP:**

- Clientes MCP lançam o servidor com **ambiente reduzido** — o `JAVA_HOME` do seu usuário pode não chegar ao processo. Se o caminho rápido se declarar indisponível por "could not start a JVM", defina `JAVA_HOME` no bloco `env` da configuração do cliente.
- O runtime precisa do módulo **`jdk.charsets`** (o `mpxj` lê campos em MacRoman). Runtimes jlink enxutos — o pacote PyPI `jdk4py`, por exemplo — não o incluem e falham com `UnsupportedCharsetException: MacRoman`; nesse caso o servidor cai no COM e o `source` traz o motivo. Confira com `java --list-modules`.
- O Java 8 (download padrão de java.com) não atende — o `mpxj` exige 9+; um JDK LTS recente (17/21) é o caminho seguro.

---

## Tamanho das respostas

`get_tasks` num cronograma real devolvia **8,07 MB**. Não chega como lentidão — chega como conversa que acaba cedo.

Medido ponta a ponta com 8.243 tarefas, servidor completo:

| resposta | antes | depois |
|---|---|---|
| `get_tasks` | 7,46 MB | 81,5 KB |
| `filter_tasks` | 7,66 MB | 81,7 KB |
| `get_schedule_analysis` | 1,93 MB | 37,4 KB |
| `get_wbs_structure` | 3,51 MB | 43,6 KB |
| definições das 99 tools | 50,7 KB | 37,5 KB |

Três decisões, nesta ordem de efeito:

**Limite de 200 itens por listagem.** É o único corte que muda o resultado: compactar 64% de algo cem vezes maior que a janela ainda deixa algo inconsumível. A resposta traz `count`/`total` completos e um bloco `page` com `returned`, `offset` e como pedir o resto. Duas portas continuam sem teto: `filter_tasks` com `limit=-1`, e `export_csv`, que escreve arquivo em vez de resposta.

**Campos vazios omitidos.** 57% dos valores de um cronograma real são vazios, zero ou `False`. **Chave ausente significa vazio, zero ou falso** — quem lia `task["notes"]` passa a precisar de `task.get("notes", "")`. Duas exceções: os campos que identificam a linha (`unique_id`, `id`, `name`) e os números onde zero é medição (`percent_complete`, `total_slack_days`, `free_slack_days`, `duration_days`, `outline_level`).

**JSON compacto acima de 4 KB.** Abaixo disso continua indentado — resposta pequena é lida por gente.

Nas definições das ferramentas, o `"title"` que o pydantic gera para cada parâmetro (`"title": "Unique Id"` ao lado de `unique_id`) foi removido: 8 KB, ~2,3k tokens, que todo cliente paga em toda sessão antes da primeira chamada. As descrições das 40 maiores foram reescritas para manter enums, formatos, unidades e sentinelas, e largar a prosa.

`get_wbs_structure` passou a limitar em `max_level=3` por padrão, e diz quantas tarefas ficaram abaixo do corte. Passe `0` para a árvore inteira.

---

## Limitações conhecidas

- **Referências COM obsoletas** — com vários projetos abertos, trocar de projeto invalida as referências existentes. Chame `switch_project` antes de operar em outro arquivo.
- **Desfazer** — `undo_last` suporta até 10 operações consecutivas. O undo via COM é menos confiável que o da interface.
- **Bloqueio de arquivo** — só um processo mantém a conexão COM. Não deixe caixas de diálogo do MS Project abertas com o servidor ativo.
- **Fusos horários** — o COM pode devolver datas com timezone; o servidor normaliza via `_to_naive()`.
- **Tarefas recorrentes** — `RecurringTaskInsert` do MS Project só funciona por diálogo. `add_recurring_task` simula a recorrência criando ocorrências individuais sob uma tarefa-resumo. Requer `python-dateutil`.
- **Dados timephased** — `get_timephased_data` fica lento em intervalos longos. Prefira semanas/meses a anos.

---

## Testes

```bash
python tests/test_new_tools.py  # 11 verificações (CRUD básico)
python tests/test_phase2.py     # 10 verificações (recursos, baselines, WBS)
python tests/test_phase3.py     # 15 verificações (campos personalizados, calendários, cronograma)
python tests/test_phase4.py     # 25 verificações (operações avançadas, multiprojeto, filtros)
python tests/test_phase5.py     # 11 verificações (correções, custo/trabalho)
python tests/test_phase6.py     # 63 verificações (timephased, calendários, variação)
python tests/test_phase7.py     # 45 verificações (caminho crítico, what-if)
```

**180 verificações** em 7 suítes. Não são funções `pytest` — cada suíte é um script sequencial com asserts inline. Todas exigem o MS Project em execução: elas criam e fecham projetos temporários.

Duas suítes tratam dos dois caminhos de leitura:

```bash
python tests/test_vistas_paridade.py  # 96 verificações — não precisa de Windows
python tests/test_backend_parity.py   # exige Windows, MS Project e projeto SALVO
```

`test_vistas_paridade.py` cobre também o corte de payload: que o limite vale, que campos vazios não são publicados, que zero permanece, que respostas grandes saem compactas e que nenhum schema publicado carrega `title`. Ela roda cada ferramenta convertida duas vezes sobre um projeto falso — uma forçada pelo COM, outra recebendo a lista já lida — e exige resposta idêntica. Verifica também que o caminho rápido não toca em nenhuma propriedade COM e que uma ferramenta que publica cinco campos não lê trinta e cinco. Ela prova que os dois **corpos** concordam; que o `mpxj` e o Microsoft Project leem o mesmo arquivo do mesmo jeito, só `test_backend_parity.py` prova — e essa nunca rodou, porque a VM de desenvolvimento é ARM64.

Antes de qualquer commit, rode os portões — eles não precisam de Windows nem do MS Project:

```bash
./tools/gates.sh
```

| Portão | O que verifica |
|---|---|
| `check_names.py` | nome carregado sem nada que o defina — pega helper esquecido num import |
| `snap_tools.py` | as 99 ferramentas ainda registradas, comparando por nome contra um baseline |
| `check_unused.py` | import morto e banner de seção sem código embaixo |

O registro de ferramentas é comparado contra `tools/baseline_tools.json`. Se você adicionar ou remover uma ferramenta de propósito, regere-o:

```bash
python3 tools/snap_tools.py . tools/baseline_tools.json
```

---

## Arquitetura

Onze módulos sobre o framework FastMCP. `server.py` não registra ferramenta alguma: importa os nove que registram e roda o servidor.

| Módulo | Ferramentas | Linhas |
|---|---|---|
| `server.py` | — | 112 |
| `msp_core.py` | — | 538 |
| `msp_fast.py` | — | 422 |
| `msp_projects.py` | 17 | 635 |
| `msp_tasks_read.py` | 17 | 869 |
| `msp_tasks_write.py` | 19 | 1007 |
| `msp_dependencies.py` | 5 | 297 |
| `msp_resources.py` | 11 | 605 |
| `msp_calendars.py` | 10 | 485 |
| `msp_schedule.py` | 11 | 996 |
| `msp_baselines_costs.py` | 6 | 388 |
| `msp_customfields.py` | 3 | 168 |

Total: 6522 linhas.

`msp_fast.py` é o outro leitor: analisa o `.mpp` salvo com `mpxj` em vez de percorrer o COM, e `msp_core.VistaCOM` apresenta uma tarefa COM sob os mesmos nomes de campo — é o que deixa o corpo de uma ferramenta ser escrito uma vez e servido pelos dois. Ver [Dois caminhos de leitura](#dois-caminhos-de-leitura).

`msp_core.py` é a fronteira COM — nenhum módulo de ferramenta fala com o Microsoft Project sem passar por ele. Todo `import win32com` está **dentro** de corpo de função, nunca no topo. É por isso que os módulos importam limpo em macOS e Linux sem pywin32, e é o que permite aos portões em `tools/` verificarem o registro de ferramentas sem Windows.

Cada módulo carrega no cabeçalho a regra que decide o que pertence a ele. Alguns casos não são óbvios e estão documentados lá:

- `set_task_calendar` e `set_resource_calendar` moram em `msp_calendars`, não com tarefas ou recursos — o sujeito é o calendário, a tarefa é só o alvo.
- `export_csv` mora em `msp_tasks_read`, não ao lado de `export_xml`, porque chama `filter_tasks` e essa aresta precisa ficar dentro de um módulo.
- As tabelas de custo ficam em `msp_resources`, não com relatórios de custo: elas escrevem em `r.CostRateTables`, ou seja, editam o recurso.

### Requer Windows de verdade

Nove ferramentas dirigem a janela do MS Project em vez do modelo de objetos, usando `SelectRow` com `EditCut`/`EditCopy`/`EditPaste` ou `OutlineIndent`/`OutlineOutdent`. Elas dependem de estado de seleção dentro do aplicativo, que nenhum fake reproduz — um teste offline passaria sem fazer nada:

`add_recurring_task` · `copy_task_structure` · `delete_task` · `indent_task` · `move_task` · `undo_last` · `insert_subproject` · `apply_filter` · `set_project_calendar`

---

## Contribuindo

Veja [CONTRIBUTING.md](CONTRIBUTING.md) para setup, convenções e como enviar mudanças.

## Licença

MIT — veja [LICENSE](LICENSE).
