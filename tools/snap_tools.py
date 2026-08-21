#!/usr/bin/env python3
"""
Gate 2 — MCP tool-registration snapshot.

Catches the other failure mode of splitting server.py: the 99 tools register
as a side effect of the @mcp.tool decorator running at import time. Drop a
module from server.py's import list and its tools vanish from the registry —
no ImportError, no traceback. The client just reports an unknown tool.

Snapshots every registered tool's name and input schema, canonically ordered,
plus a sha256 over the whole thing. Comparing the JSON (not just the hash or
the count) is what makes a swap visible: one tool disappearing while another
appears keeps the total at 99.

Usage:
    python3 tools/snap_tools.py REPO_DIR OUT.json
    python3 tools/snap_tools.py REPO_DIR OUT.json --baseline tools/baseline_tools.json

With --baseline, diffs against it and exits 1 on any difference.

Runs on macOS/Linux without pywin32: every win32com import in server.py is
inside a function body, so importing the module touches no COM.
"""

import asyncio
import hashlib
import json
import os
import sys


def snapshot(repo_dir):
    sys.path.insert(0, os.path.abspath(repo_dir))
    import server  # noqa: E402  — path must be set first

    tools = asyncio.run(server.mcp.list_tools())
    entries = [
        {"name": t.name, "inputSchema": t.inputSchema}
        for t in sorted(tools, key=lambda t: t.name)
    ]
    canonical = json.dumps(entries, sort_keys=True, ensure_ascii=False, indent=1)
    digest = hashlib.sha256(canonical.encode("utf-8")).hexdigest()
    return entries, canonical, digest


def diff_against(entries, baseline_path):
    with open(baseline_path, "r", encoding="utf-8") as fp:
        base = json.load(fp)

    now_by_name = {e["name"]: e["inputSchema"] for e in entries}
    base_by_name = {e["name"]: e["inputSchema"] for e in base}

    added = sorted(set(now_by_name) - set(base_by_name))
    removed = sorted(set(base_by_name) - set(now_by_name))
    changed = sorted(
        n for n in set(now_by_name) & set(base_by_name)
        if now_by_name[n] != base_by_name[n]
    )

    for name in removed:
        print(f"  REMOVED  {name}")
    for name in added:
        print(f"  ADDED    {name}")
    for name in changed:
        print(f"  SCHEMA   {name}")

    return len(added) + len(removed) + len(changed)


def main(argv):
    if len(argv) < 3:
        print("usage: snap_tools.py REPO_DIR OUT.json [--baseline PATH]", file=sys.stderr)
        return 2

    repo_dir, out_path = argv[1], argv[2]
    baseline_path = None
    if "--baseline" in argv:
        baseline_path = argv[argv.index("--baseline") + 1]

    entries, canonical, digest = snapshot(repo_dir)

    with open(out_path, "w", encoding="utf-8") as fp:
        fp.write(canonical + "\n")

    print(f"TOOLS: {len(entries)}")
    print(f"SHA256: {digest}")

    if baseline_path:
        drift = diff_against(entries, baseline_path)
        if drift:
            print(f"DRIFT: {drift} tool(s) differ from baseline")
            return 1
        print("DRIFT: none")

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
