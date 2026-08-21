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

Install: pip install -e .
Run:     msproject-mcp  (or: python server.py)
Register in claude_desktop_config.json (see the README).

The tools are only reachable with Microsoft Project already running and a file
open; the server attaches to a live instance rather than launching one.
"""

import argparse
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


def main():
    """Console entry point. Also reachable as `python server.py`.

    Defaults to stdio, which is what an MCP client on the same desktop session
    expects. The HTTP transports exist for one specific case, described under
    --transport below.
    """
    parser = argparse.ArgumentParser(
        prog="msproject-mcp",
        description="MCP server controlling Microsoft Project via COM.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
Running from another machine
----------------------------
COM cannot be reached across a Windows logon session. GetActiveObject reads
the Running Object Table, which is per-session, so a process in session 0 --
where an SSH login lands -- cannot see Microsoft Project running on the
interactive desktop in session 1. It fails with MK_E_UNAVAILABLE even though
the application is open.

So this server cannot be launched over SSH the way a stdio server normally is.
It has to run inside the same session as Microsoft Project, and expose itself
over the network instead:

    msproject-mcp --transport streamable-http --host 0.0.0.0 --port 8765

Start that from the Windows desktop -- a terminal, a shortcut, or a Startup
entry -- never from an SSH shell. The client then connects to
http://<windows-host>:8765/mcp.

--host 0.0.0.0 accepts connections from any interface and the server has no
authentication of its own. Bind it only to a network you trust, such as a
Parallels or VMware host-only network, and leave it on 127.0.0.1 otherwise.
""",
    )
    parser.add_argument(
        "--transport", choices=["stdio", "sse", "streamable-http"], default="stdio",
        help="stdio (default) for a client on this desktop; the HTTP transports for a remote client",
    )
    parser.add_argument(
        "--host", default="127.0.0.1",
        help="interface to bind for the HTTP transports (default: 127.0.0.1)",
    )
    parser.add_argument(
        "--port", type=int, default=8765,
        help="port for the HTTP transports (default: 8765)",
    )
    args = parser.parse_args()

    # stdout is the MCP stdio transport -- anything written here corrupts the
    # protocol. Diagnostics go to stderr.
    print("Starting MS Project MCP Server...", file=sys.stderr)
    print("MS Project must be running with a file open before using tools.", file=sys.stderr)

    if args.transport != "stdio":
        mcp.settings.host = args.host
        mcp.settings.port = args.port
        path = mcp.settings.sse_path if args.transport == "sse" else mcp.settings.streamable_http_path
        print(f"Listening on http://{args.host}:{args.port}{path}", file=sys.stderr)
        if args.host not in ("127.0.0.1", "localhost"):
            print("This port has no authentication -- keep it on a trusted network.",
                  file=sys.stderr)

    mcp.run(transport=args.transport)


if __name__ == "__main__":
    main()
