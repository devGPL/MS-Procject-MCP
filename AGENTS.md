# CLAUDE.md

Guidance for AI coding agents (Claude Code, OpenAI Codex, etc.) working in this repository.

## What this is

**MS Project MCP Server** — a single-file Python MCP server (`server.py`, ~5,200 lines, 99 tools) that drives a **live Microsoft Project desktop application via Windows COM automation** (pywin32). It is NOT a file parser: there is no mpxj, no Java, no JVM anywhere. All reads and writes go through MS Project's own COM API (`win32com.client`), so `.mpp` save/load works natively.

**Hard platform constraint: Windows only, with Microsoft Project installed (tested on 16.0).** The server cannot run on macOS/Linux. Development/editing of the code can happen anywhere; execution and tests require Windows + MS Project.

## Repo layout

```
server.py          # Everything: FastMCP server "MS Project", all 99 tools, helpers
tests/             # 7 live-integration suites (202 tests), one per feature phase
README.md          # Authoritative docs: tool inventory, install, test commands
CONTRIBUTING.md    # Code conventions + PR flow (its Testing section is STALE — trust README)
```

There is no `requirements.txt`, no packaging, no env vars, no config files. Full install:

```bash
pip install mcp pywin32 python-dateutil
```

(`python-dateutil` is optional — only `add_recurring_task` imports it, lazily.)

## Running / registering the server

- Standalone: `python server.py` (stdio transport, FastMCP default). Prints a banner; MS Project should be running with a file open (or use `open_project`/`new_project`, which auto-launch it via `Dispatch`).
- Claude Desktop (`claude_desktop_config.json`) — also shown in a comment at the bottom of server.py:
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

## Architecture (all in server.py)

- **No Python-side state.** State lives entirely in the running MS Project instance. `get_app(require_project=True)` (line ~19) attaches via `GetActiveObject("MSProject.Application")` and raises if MS Project isn't running / no project open. Every tool operates on `app.ActiveProject`.
- **Tool = `@mcp.tool()` function returning a JSON string** (`json.dumps(..., indent=2)`). Errors are returned as `{"error": "..."}` JSON, never raised. Docstrings become MCP tool descriptions — keep them accurate.
- **Helpers (lines 19–181):** `get_app`/`get_proj`, `_parse_date` (strict `YYYY-MM-DD`), `_fmt_date`, `_to_naive` (strips tzinfo from COM datetimes before comparing to `datetime.now()`), `_find_task` (linear scan by UniqueID), `_get_mpd` (MinutesPerDay, fallback 480), `_custom_field_id` (maps "Text5"/"Number1"/"Flag3"/"Date1"/"Duration2" → pjCustomTask COM field IDs), `task_to_dict` (~35-field task dict).
- **File sections** are marked with `# ---` banner comments naming the phase (Phase 3 … Phase 7). README's phase table (25 → 44 → 65 → 79 → 96 → 99 tools) maps the file's growth.

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

- No pytest — plain asyncio scripts calling tools **in-process** via `await mcp.call_tool(name, kwargs)` on the imported FastMCP instance.
- Phases 4–7 exit non-zero on failure; earlier suites always exit 0 (read the printed summary).
- **New tests: follow the phase 6–7 style** — `from server import mcp` (sys.path insert), kwargs-style `call()` wrapped in try/except so COM errors register as FAIL, `ok()`/`skip()` helpers, boxed summary, `sys.exit(0 if FAIL == 0 else 1)`.
- Use generic placeholder names ("Alice", "Task A", "Programme Alpha") and far-future dates (2026–2027) unless overdue behavior is under test. Verify mutations by re-reading state, not just the tool's status reply.
- CONTRIBUTING.md's Testing section is outdated (pre-`tests/` paths, missing suites) — README's Tests section is authoritative.

## Making changes

- All code goes in `server.py`; keep the phase-banner section structure. New tool = `@mcp.tool()` + docstring + `get_app()`/`get_proj()` + JSON-string return + `{"error": ...}` on failure + `FileSave()` if mutating + a test in the matching `tests/test_phaseN.py` + README tool-inventory update.
- Branch naming: `feature/your-feature-name`. Clear commit messages; no strict commit convention.
- When adding date logic, go through `_parse_date`/`_fmt_date`/`_to_naive` — never compare aware COM datetimes to naive `datetime.now()` directly.
