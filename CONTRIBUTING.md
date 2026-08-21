# Contributing to MS Project MCP Server

Thanks for your interest in contributing! Here's how to get started.

## Prerequisites

- Windows with Microsoft Project installed, to run the tools
- Python 3.10+
- `pip install -e .` — installs mcp, pywin32 (Windows only) and python-dateutil

The modules import cleanly on macOS and Linux without pywin32, because every
COM import sits inside a function body. The gates below run there too; only
the test suites need Windows.

## Getting Started

1. Fork the repository
2. Clone your fork:
   ```bash
   git clone git@github.com:<your-username>/MS-Procject-MCP.git
   cd MS-Procject-MCP
   ```
3. Create a branch:
   ```bash
   git checkout -b feature/your-feature-name
   ```

## Development

All server code lives in a single file: `server.py`. New tools follow this pattern:

```python
@mcp.tool()
def your_tool(param: str) -> str:
    """Short description of what the tool does."""
    app  = get_app()
    proj = get_proj(app)

    # ... your logic ...

    return json.dumps({"status": "ok", ...}, indent=2)
```

### Conventions

- Put the tool in the module that owns its subject. Each `msp_*.py` states its
  boundary rule in the module docstring — read that before adding a tool, and
  extend it if your tool is a genuine edge case.
- Keep every `import win32com` inside the function body. Module-scope COM
  imports break importing on non-Windows machines, and with it every gate.
- Use `get_app()` / `get_proj()` for COM access
- Use `_to_naive()` when comparing COM dates with `datetime.now()`
- Use `_fmt_date()` to format dates for JSON output
- Use `_parse_date()` to convert `YYYY-MM-DD` strings for COM input
- Use `_find_task()`, `_find_resource()` and `_find_calendar()` rather than
  writing another `for x in proj.<Collection>` scan. The case-insensitivity
  bug fixed in this release existed because seven copies of one lookup had
  drifted apart.
- Use `_uid_map()` in bulk tools — calling `_find_task()` per item rescans the
  whole task collection every time
- Return JSON strings from all tools
- Include docstrings — they become the tool description in MCP

## Testing

### Gates — run these before every commit

No Windows, no MS Project, under half a second:

```bash
./tools/gates.sh
```

| Gate | Catches |
|---|---|
| `check_names.py` | a name loaded with nothing binding it — a helper missed in an import list |
| `snap_tools.py` | a tool gone from the registry, compared by name against a baseline |
| `check_unused.py` | dead imports, section banners with no code under them |

`snap_tools.py` matters more than it looks. Tools register as a side effect of
the `@mcp.tool` decorator running at import, so dropping a module from
`server.py`'s import list removes its tools with no ImportError to point at the
cause — the client just reports an unknown tool.

If you add or remove a tool deliberately, regenerate the baseline:

```bash
python3 tools/snap_tools.py . tools/baseline_tools.json
```

### Integration suites — Windows only

Each file creates a temporary project, runs its checks, and cleans up. These
are sequential scripts with inline asserts, not pytest test functions.

```bash
python tests/test_new_tools.py  # 11 checks
python tests/test_phase2.py     # 10
python tests/test_phase3.py     # 15
python tests/test_phase4.py     # 25
python tests/test_phase5.py     # 11
python tests/test_phase6.py     # 63
python tests/test_phase7.py     # 45
```

180 checks in total. **MS Project must be running before you execute them.**
Run one file at a time and close the temporary project between runs: six of
the seven have no try/finally, so a crash leaves the wrong ActiveProject
behind and the next suite fails for an unrelated reason.

15 of the 99 tools are never exercised by any suite, including `add_task`,
`delete_task`, `open_project` and `save_project`.

### Adding Tests

If you add a new tool, add corresponding tests. Follow the existing pattern:

```python
async def call(tool_name, **kwargs):
    result = await mcp.call_tool(tool_name, kwargs)
    contents = result[0] if isinstance(result, tuple) else result
    text = contents[0].text if contents else ""
    return json.loads(text) if text else {}
```

## Submitting Changes

1. Run all test phases and confirm they pass
2. Commit with a clear message describing what and why
3. Push to your fork and open a Pull Request
4. Describe the change, link any related issues, and note which tests cover it

## Reporting Issues

Open an issue with:
- What you expected vs what happened
- MS Project version (Help > About)
- Python version (`python --version`)
- Relevant error output

## Code of Conduct

Be respectful and constructive. We're all here to make MS Project automation better.
