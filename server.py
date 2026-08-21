"""
MS Project MCP Server — entry point.

Controls local Microsoft Project via COM automation. This file registers no
tools of its own: it imports the nine modules that do, and runs the server.

    msp_core              COM boundary, shared helpers, the `mcp` instance
    msp_projects          project lifecycle, multi-project, interchange
    msp_tasks_read        queries, filters, roll-ups, exports
    msp_tasks_write       creation, update, deletion, bulk edits
    msp_dependencies      the precedence network
    msp_resources         pool, assignments, availability, rates
    msp_calendars         base calendars, exceptions, working hours
    msp_schedule          critical path, slack, leveling, status rolls
    msp_baselines_costs   planned versus actual

Install: pip install mcp pywin32 python-dateutil
Run:     python server.py
Register in claude_desktop_config.json (see the README).

The tools are only reachable with Microsoft Project already running and a file
open; the server attaches to a live instance rather than launching one.
"""

import sys

from msp_core import mcp

# Tool modules register themselves on `mcp` as a side effect of import. These
# are NOT unused imports: dropping one silently removes its tools from the
# registry, with no ImportError to point at the cause. tools/snap_tools.py
# exists to catch exactly that.
import msp_baselines_costs  # noqa: F401
import msp_calendars  # noqa: F401
import msp_customfields  # noqa: F401
import msp_dependencies  # noqa: F401
import msp_projects  # noqa: F401
import msp_resources  # noqa: F401
import msp_schedule  # noqa: F401
import msp_tasks_read  # noqa: F401
import msp_tasks_write  # noqa: F401


if __name__ == "__main__":
    # stdout is the MCP stdio transport -- anything written here corrupts the
    # protocol. Diagnostics go to stderr.
    print("Starting MS Project MCP Server...", file=sys.stderr)
    print("MS Project must be running with a file open before using tools.", file=sys.stderr)
    mcp.run()
