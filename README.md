# MS Project MCP Server

Controla o Microsoft Project via automação COM através do Model Context Protocol (MCP).

99 ferramentas para ler, editar e analisar cronogramas direto do seu assistente de IA — sem sair do chat.

---

## Requisitos

| Item | Detalhe |
|------|---------|
| SO | **Windows** (a automação COM não existe em macOS/Linux) |
| Microsoft Project | Instalado e **em execução** (testado no MS Project 16.0) |
| Python | 3.10 ou superior |

```bash
pip install mcp pywin32 python-dateutil
```

> `python-dateutil` só é necessário para `add_recurring_task`.

---

## Instalação

1. Clone o repositório e anote o caminho absoluto de `server.py`.

```bash
git clone <url-do-repo>
cd MS-Procject-MCP
pip install mcp pywin32 python-dateutil
```

2. Teste que o servidor sobe:

```bash
python server.py
```

3. Configure o cliente (Claude ou Codex) — instruções abaixo.

> **Antes de usar qualquer ferramenta**: o MS Project precisa estar aberto com um arquivo carregado. O servidor se conecta a uma instância já em execução; ele não abre o Project sozinho.

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
      "command": "python",
      "args": ["C:\\caminho\\para\\MS-Procject-MCP\\server.py"]
    }
  }
}
```

Reinicie o Claude Desktop. O servidor aparece no ícone de ferramentas (🔨) do chat.

### Claude Code

Via CLI (recomendado — grava a config pra você):

```bash
claude mcp add msproject --scope user -- python C:\caminho\para\MS-Procject-MCP\server.py
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
      "command": "python",
      "args": ["C:\\caminho\\para\\MS-Procject-MCP\\server.py"]
    }
  }
}
```

---

## Configuração — Codex

### Via CLI

```bash
codex mcp add msproject -- python C:\caminho\para\MS-Procject-MCP\server.py
```

Conferir:

```bash
codex mcp list
```

### Via arquivo

Edite `~/.codex/config.toml` (no Windows: `%USERPROFILE%\.codex\config.toml`):

```toml
[mcp_servers.msproject]
command = "python"
args = ["C:\\caminho\\para\\MS-Procject-MCP\\server.py"]
```

> **Barras invertidas no TOML**: escape com `\\` (como acima) ou use string literal com aspas simples: `args = ['C:\caminho\para\server.py']`.

Se o Python não estiver no `PATH`, aponte o executável completo:

```toml
[mcp_servers.msproject]
command = "C:\\Python312\\python.exe"
args = ["C:\\caminho\\para\\MS-Procject-MCP\\server.py"]
```

---

## Verificando a conexão

Peça ao assistente para rodar `health_check`. A resposta traz a versão do MS Project e o status do arquivo aberto. Se falhar:

| Erro | Causa | Solução |
|------|-------|---------|
| `MS Project is not running` | Nenhuma instância ativa | Abra o MS Project |
| `No project file is open` | Project aberto sem arquivo | Abra ou crie um `.mpp` |
| Servidor não aparece no cliente | Caminho errado ou JSON/TOML inválido | Confira o caminho absoluto e reinicie o cliente |
| `ModuleNotFoundError: win32com` | `pywin32` ausente no Python usado | Instale no mesmo interpretador que está no `command` |

---

## Ferramentas (99)

### Projeto (7)
`open_project` · `new_project` · `get_project_info` · `set_project_properties` · `save_project` · `save_project_as` · `close_project`

### Consulta de tarefas (9)
`get_tasks` · `get_task` · `get_critical_path` · `get_tasks_by_rag` · `get_overdue_tasks` · `get_tasks_by_resource` · `search_tasks` · `get_progress_summary` · `get_wbs_structure`

### Edição de tarefas (12)
`update_task` · `bulk_update_rag` · `bulk_update_tasks` · `add_task` · `bulk_add_tasks` · `add_recurring_task` · `delete_task` · `set_task_mode` · `bulk_set_task_mode` · `set_constraint` · `clear_estimated_flags` · `indent_task`

### Dependências (4)
`add_predecessor` · `bulk_add_predecessors` · `remove_predecessor` · `get_task_dependencies`

### Recursos (7)
`get_resources` · `add_resource` · `assign_resource` · `update_resource` · `delete_resource` · `set_resource_calendar` · `get_resource_availability`

### Alocação de recursos (3)
`bulk_assign_resources` · `remove_resource_assignment` · `get_resource_workload`

### Tabelas de custo (2)
`get_resource_rate_tables` · `set_resource_rate_table`

### Campos personalizados (3)
`rename_custom_fields` · `update_custom_fields` · `get_custom_field_values`

### Importação / exportação (6)
`import_xml` · `export_xml` · `export_csv` · `snapshot_to_json` · `snapshot_diff` · `insert_subproject`

### Calendários (8)
`get_calendars` · `create_calendar` · `delete_calendar` · `set_calendar_exception` · `delete_calendar_exception` · `list_calendar_exceptions` · `set_project_calendar` · `set_working_hours`

### Cronograma e análise (12)
`get_schedule_analysis` · `validate_schedule` · `calculate_project` · `get_milestone_report` · `level_resources` · `find_available_slack` · `get_constraints` · `set_task_calendar` · `set_task_hyperlink` · `get_critical_path_sequence` · `get_critical_tasks_for_period` · `what_if_delay`

### Atualização de progresso (2)
`update_project` — marca tudo concluído até uma data (ritual semanal do PMO) · `reschedule_incomplete_work`

### Dados timephased (1)
`get_timephased_data` — dados período a período de trabalho/custo (curva S, fluxo de caixa)

### Linhas de base e valor agregado (4)
`save_baseline` · `clear_baseline` · `compare_baselines` · `get_earned_value` (BCWS, BCWP, ACWP, SPI, CPI)

### Variação e relatórios (1)
`get_variance_report`

### Custo e trabalho (2)
`get_cost_summary` · `get_actual_work`

### Acompanhamento (2)
`get_progress_by_wbs` · `get_dependency_chain`

### Operações avançadas (8)
`set_deadline` · `bulk_set_deadlines` · `set_task_active` · `dry_run_bulk_update` · `move_task` · `copy_task_structure` · `cross_project_link` · `undo_last`

### Multiprojeto (3)
`list_projects` · `switch_project` · `apply_filter`

### Filtro e agrupamento (2)
`filter_tasks` · `group_tasks_by`

### Conectividade (1)
`health_check`

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
python tests/test_new_tools.py  # 11 testes (CRUD básico)
python tests/test_phase2.py     # 10 testes (recursos, baselines, WBS)
python tests/test_phase3.py     # 15 testes (campos personalizados, calendários, cronograma)
python tests/test_phase4.py     # 23 testes (operações avançadas, multiprojeto, filtros)
python tests/test_phase5.py     # 11 testes (correções, custo/trabalho)
python tests/test_phase6.py     # 75 testes (timephased, calendários, variação)
python tests/test_phase7.py     # 57 testes (caminho crítico, what-if)
```

**202 testes** em 7 suítes. Todas exigem o MS Project em execução — elas criam e fecham projetos temporários.

---

## Arquitetura

Servidor em arquivo único (`server.py`, ~5.200 linhas) sobre o framework FastMCP. Todas as chamadas COM passam pelos helpers `get_app()` / `get_proj()`. Datas são normalizadas com `_to_naive()` e formatadas com `_fmt_date()`.

### Evolução por fase

| Fase | Ferramentas | Foco |
|------|-------------|------|
| 1–2 | 25 | CRUD básico, dependências, recursos |
| 3 | 44 | Campos personalizados, calendários, análise de cronograma |
| 4 | 65 | Operações avançadas, multiprojeto, filtros |
| 5 | 79 | Correções, controle de custo/trabalho |
| 6 | 96 | Timephased, gestão de calendários, disponibilidade, variação |
| 7 | 99 | Inteligência de caminho crítico: sequência ordenada, filtro por período, what-if |

---

## Contribuindo

Veja [CONTRIBUTING.md](CONTRIBUTING.md) para setup, convenções e como enviar mudanças.

## Licença

MIT — veja [LICENSE](LICENSE).
