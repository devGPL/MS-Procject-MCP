# CLAUDE.md

Guidance for AI coding agents (Claude Code, OpenAI Codex, etc.) working in this repository.

## What this is

**MS Project MCP Server** — a Python MCP server (11 modules, ~6,500 lines, 99 tools) that drives a **live Microsoft Project desktop application via Windows COM automation** (pywin32). Every write, and every read that needs the application's own state, goes through MS Project's COM API (`win32com.client`), so `.mpp` save/load works natively.

Reads have a second path. Sixteen scanning tools parse the SAVED `.mpp` with `mpxj` (JVM, via `jpype`) when it is installed and the file is reachable, because obtaining a task object over COM costs 2.635 ms against 0.240 ms to read a property — 83% of a scan spent before a field is read. Every such response carries a `source` block saying which path answered. See `msp_fast.py`, whose header explains the tradeoff, and the README section "Dois caminhos de leitura".

**Hard platform constraint: Windows only, with Microsoft Project installed (tested on 16.0).** The server cannot run on macOS/Linux. Development/editing of the code can happen anywhere; execution and tests require Windows + MS Project.

## Repo layout

```
server.py             # Entry point only: imports the tool modules, runs the server
msp_core.py           # COM boundary: get_app/get_proj, task_to_dict, VistaCOM, helpers
msp_fast.py           # The other reader: parses the saved .mpp with mpxj
msp_projects.py       # 17 tools — file lifecycle, multi-project, import/export
msp_tasks_read.py     # 17 tools — listings, filters, roll-ups, CSV
msp_tasks_write.py    # 19 tools — task CRUD, modes, constraints, deadlines
msp_dependencies.py   #  5 tools — links between tasks
msp_resources.py      # 11 tools — pool, assignments, availability, rate tables
msp_calendars.py      # 10 tools — base calendars, exceptions, working hours
msp_schedule.py       # 11 tools — critical path, slack, validation, leveling
msp_baselines_costs.py#  6 tools — baselines, variance, earned value, cost
msp_customfields.py   #  3 tools — Text/Number/Date/Flag/Duration fields
tests/                # 7 live-integration suites + 2 backend-parity suites
tools/                # Offline gates: ./tools/gates.sh, no Windows needed
README.md             # Authoritative docs: tool inventory, install, test commands
CONTRIBUTING.md       # Code conventions + PR flow (its Testing section is STALE — trust README)
```

Packaged with `pyproject.toml`; no env vars, no config files. Full install:

```bash
pip install -e .            # mcp, pywin32 (Windows only), python-dateutil
pip install -e ".[fast]"    # adds mpxj + jpype1 for the fast read path
```

`jpype1` publishes no wheel for Windows on ARM, and needs Java 9+. Without it every read falls back to COM and says so — correct, and slow.

## Running / registering the server

- Standalone: `msproject-mcp`, or `python server.py` (stdio transport, FastMCP default). Prints a banner; MS Project should be running with a file open (or use `open_project`/`new_project`, which auto-launch it via `Dispatch`).
- **The server must start from the Windows desktop session, unelevated.** Three separate conditions produce the same attach failure, and running as administrator -- the instinct on an access error -- is one of them. `msp_core.get_app()` documents all three.
- Claude Desktop (`claude_desktop_config.json`):
  ```json
  { "mcpServers": { "msproject": { "command": "python", "args": ["C:/path/to/server.py"] } } }
  ```
- Claude Code: `claude mcp add msproject -- python C:/path/to/server.py`
- OpenAI Codex (`~/.codex/config.toml`):
  ```toml
  [mcp_servers.msproject]
  command = "python"
  args = ["C:/path/to/server.py"]
  ```

## Architecture

- **No Python-side state.** State lives entirely in the running MS Project instance (and, for the fast path, in the saved file). `get_app(require_project=True)` attaches via `GetActiveObject("MSProject.Application")` and raises if MS Project isn't running / no project open. Every COM tool operates on `app.ActiveProject`.
- **`msp_core.py` is the COM boundary.** No tool module talks to Microsoft Project without it, and every `import win32com` sits INSIDE a function body — which is why the modules import cleanly on macOS/Linux and the gates in `tools/` can check tool registration without Windows.
- **Tool = `@mcp.tool()` function returning a JSON string** (`json.dumps(..., indent=2)`). Errors are returned as `{"error": "..."}` JSON, never raised. Docstrings become MCP tool descriptions — keep them accurate.
- **Each module's header states the rule that decides what belongs in it.** Read it before moving a tool; several placements look wrong and are deliberate (`set_task_calendar` lives with calendars, `export_csv` with task reads).
- **Core helpers:** `get_app`/`get_proj`, `_com_retry` (retries the busy-application errors), `calculo_suspenso` (context manager for bulk writes), `_parse_date` (strict `YYYY-MM-DD`), `_fmt_date`, `_dia`/`_dt_de_vista` (view date fields), `_to_naive`, `_find_task`, `_uid_map`, `_get_mpd` (MinutesPerDay, fallback 480), `_custom_field_id`, `task_to_dict` (35-field task dict).
- **One table of field readers.** `_LEITORES_TAREFA` in `msp_core` defines the 35 published task fields once. `task_to_dict` materialises all of them; `VistaCOM` evaluates one at a time on demand and caches. A scanning tool is written against key names and either backend serves it: `msp_fast.varredura(proj)` returns parsed dicts or `None`, plus the `source` block, and the body reads `v["critical"]` either way. Do NOT "simplify" that by calling `task_to_dict` on the COM side — it costs 35 property reads where a tool needs three, which on the ARM VM (the only path available there) turns a ~32s scan into ~92s.

## API conventions (essential for both editing code and calling tools)

- **Tasks are identified by `unique_id` (COM UniqueID)** in every tool signature. The volatile row `id` appears only inside predecessor strings (e.g. `"5FS+2d"`).
- **Dates:** always `YYYY-MM-DD` strings. Durations/slack/lag are stored by COM in minutes and converted to days via MinutesPerDay; work values convert minutes → hours.
- **RAG status convention:** the `Text1` custom field holds `Red`/`Amber`/`Green`. `rag` params/fields everywhere read/write Text1. `rename_custom_fields` can label it "RAG Status" for the UI.
- **Multi-field / bulk params are JSON-encoded strings**, not objects: single objects for property maps (`set_project_properties`, `update_custom_fields`, `filter_tasks`), lists of objects for bulk ops (`bulk_add_tasks`, `bulk_update_tasks`, `bulk_add_predecessors`, `bulk_set_deadlines`, ...). `bulk_set_task_mode` accepts either a list or a scope object.
- **Sentinel defaults mean "unchanged":** `-1` for numbers, `""` for strings, `None` for bools (e.g. `update_task`, `update_resource`). Consequence: you cannot set a field to an explicitly falsy value through those tools.
- **Most mutating tools call `app.FileSave()` immediately** — the file on disk changes on every call. Known exceptions that do NOT save: `rename_custom_fields`, `bulk_assign_resources`, `remove_resource_assignment`, `update_resource`, `bulk_set_deadlines`, `set_task_calendar`. Call `save_project` after those if persistence matters.
- **Bulk write tools suspend auto-calculation** (`app.Calculation = 0`), then `CalculateProject()` + restore in a `finally` block. Follow this pattern for new bulk tools.
- **Multi-project:** `list_projects` / `switch_project` use `get_app(require_project=False)`. Always `switch_project` before touching a different open file (see pitfalls).

## Tool map (99 tools — full inventory with signatures in README)

- **Project lifecycle:** open/new/save/save_as/close_project, set_project_properties, get_project_info, health_check
- **Task queries:** get_tasks, get_task, search_tasks, filter_tasks (rich AND-filter + sort + pagination), group_tasks_by, get_wbs_structure, get_tasks_by_rag, get_overdue_tasks, get_tasks_by_resource
- **Task writes:** add_task, bulk_add_tasks, update_task, bulk_update_tasks, bulk_update_rag, delete_task, indent_task, move_task, copy_task_structure, set_task_mode/bulk_set_task_mode, set_constraint, set_deadline/bulk_set_deadlines, set_task_active, set_task_calendar, set_task_hyperlink, clear_estimated_flags, add_recurring_task, dry_run_bulk_update (preview, no mutation)
- **Dependencies:** add_predecessor, bulk_add_predecessors, remove_predecessor, get_task_dependencies, get_dependency_chain, cross_project_link
- **Resources:** get_resources, add_resource, update_resource, delete_resource, assign_resource, bulk_assign_resources, remove_resource_assignment, get_resource_workload, get_resource_availability, set_resource_calendar, get/set_resource_rate_tables, level_resources
- **Custom fields:** update_custom_fields (Text1-30/Number1-20/Date1-10/Flag1-20/Duration1-10), get_custom_field_values, rename_custom_fields
- **Calendars:** get_calendars, create_calendar, delete_calendar, set_project_calendar, set_calendar_exception, list/delete_calendar_exceptions, set_working_hours
- **Analysis & tracking:** get_critical_path, get_critical_path_sequence (ordered chain), get_critical_tasks_for_period, what_if_delay (read-only simulation), get_schedule_analysis, validate_schedule, find_available_slack, get_progress_summary, get_progress_by_wbs, get_milestone_report, save/clear_baseline, compare_baselines, get_variance_report, get_earned_value, get_cost_summary, get_actual_work, get_timephased_data, update_project, reschedule_incomplete_work, calculate_project
- **Import/export & snapshots:** import_xml, export_xml, export_csv, snapshot_to_json, snapshot_diff (pure file diff, no COM), insert_subproject, apply_filter, undo_last
- **Multi-project:** list_projects, switch_project

## Known pitfalls (from README "Known Limitations" + code reading)

- **COM proxy staleness:** with multiple projects open, switching invalidates held COM references — `switch_project` first, and keep COM-heavy tests last in a suite.
- **Dialog deadlock:** one process owns the COM connection; an open MS Project GUI dialog blocks/fails COM calls.
- **UI-selection-dependent tools:** `delete_task`, `indent_task`, `move_task`, `copy_task_structure` use `SelectRow` + Edit menu operations — they depend on the active view.
- **`undo_last`** supports at most 10 steps and is less reliable than UI undo.
- **`get_timephased_data`** is slow on long ranges — query weeks/months, not years.
- **`add_recurring_task`** simulates recurrence (rrule-generated subtasks under a summary) because the real COM method is dialog-only; needs python-dateutil.
- **Accepted-but-ignored params (latent quirks, do not rely on them):** `add_task.after_unique_id`, `assign_resource.units` and `bulk_assign_resources` `units`, `set_calendar_exception.working`. `remove_predecessor` matches predecessor IDs by string prefix (removing ID 1 can hit `12FS`).
- **`what_if_delay`** is an estimate (pure slack arithmetic, ~1.4× working→calendar day conversion), not a real reschedule.

## Tests

Live integration tests — **require Windows + MS Project running; no mocks, no fixtures committed**. Each suite creates its own temp project via `new_project` and cleans up with `close_project(save=False)`.

```bash
python tests/test_new_tools.py   # 11 — core CRUD
python tests/test_phase2.py      # 10 — resources, WBS, baselines, EV
python tests/test_phase3.py      # 15 — multi-project, filtering, custom fields
python tests/test_phase4.py      # 23 — advanced ops, cross-project, CSV
python tests/test_phase5.py      # 11 — bug-fix validations, cost/work
python tests/test_phase6.py      # 75 — calendars, timephased, variance, rate tables
python tests/test_phase7.py      # 57 — critical path intelligence
```

Two suites are about the two read paths, and one of them needs nothing:

```bash
python tests/test_vistas_paridade.py   # 80 — runs on macOS, no MS Project
python tests/test_backend_parity.py    # needs Windows + a SAVED project
```

`test_vistas_paridade.py` runs each converted tool twice over a fake project — once forced down COM, once handed the parsed list — and requires identical output. It proves the two BODIES agree; that mpxj and Microsoft Project read the same file the same way is what `test_backend_parity.py` proves, and that one has never run (the dev VM is ARM64).

Before any commit, run the offline gates — no Windows, no MS Project, 0.43s:

```bash
./tools/gates.sh
```

They catch a name loaded with nothing defining it, a tool missing from the registry, and a dead import or orphan banner. Adding or removing a tool on purpose means regenerating the baseline: `python3 tools/snap_tools.py . tools/baseline_tools.json`.

- No pytest — plain asyncio scripts calling tools **in-process** via `await mcp.call_tool(name, kwargs)` on the imported FastMCP instance.
- Phases 4–7 exit non-zero on failure; earlier suites always exit 0 (read the printed summary).
- **New tests: follow the phase 6–7 style** — `from server import mcp` (sys.path insert), kwargs-style `call()` wrapped in try/except so COM errors register as FAIL, `ok()`/`skip()` helpers, boxed summary, `sys.exit(0 if FAIL == 0 else 1)`.
- Use generic placeholder names ("Alice", "Task A", "Programme Alpha") and far-future dates (2026–2027) unless overdue behavior is under test. Verify mutations by re-reading state, not just the tool's status reply.
- CONTRIBUTING.md's Testing section is outdated (pre-`tests/` paths, missing suites) — README's Tests section is authoritative.

## Making changes

- A new tool goes in the module whose header rule covers it = `@mcp.tool()` + docstring + `get_app()`/`get_proj()` + JSON-string return + `{"error": ...}` on failure + `FileSave()` if mutating + a test in the matching `tests/test_phaseN.py` + README tool-inventory update + `./tools/gates.sh` + a regenerated `tools/baseline_tools.json`.
- A new tool that SCANS every task should read through `msp_fast.varredura` and a view rather than walking `proj.Tasks` directly, and must publish the `source` block. Add it to `tests/test_vistas_paridade.py`, which will fail if the two backends disagree.
- Branch naming: `feature/your-feature-name`. Clear commit messages; no strict commit convention.
- When adding date logic, go through `_parse_date`/`_fmt_date`/`_to_naive` — never compare aware COM datetimes to naive `datetime.now()` directly.
