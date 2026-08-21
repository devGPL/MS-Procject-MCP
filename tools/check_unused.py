#!/usr/bin/env python3
"""
Gate 3 — dead imports and orphan section banners.

Gates 1 and 2 check that every name resolves and every tool still registers.
Neither notices the residue a module split leaves behind: an import whose only
callers moved to another file, and a section banner whose tools all left. Both
appeared in three consecutive extraction steps and both were found by hand.

Two checks, module level only:

  dead import     a top-level import never referenced in the file. Function-
                  local imports are ignored -- every COM import in this
                  project is deliberately deferred inside a function body.

  orphan banner   a `# ---- / # TITLE / # ----` block with no @mcp.tool
                  between it and the next banner.

Imports marked `# noqa: F401` are exempt, which is not a formality here: the
tool modules are imported by server.py purely so their decorators run, and
that import genuinely has no reference in the file. Dropping one silently
removes its tools -- the case Gate 2 exists to catch. The marker is how a
deliberate side-effect import is told apart from a leftover.

Usage:
    python3 tools/check_unused.py FILE [FILE ...]

Exit code 0 when nothing is flagged, 1 otherwise.
"""

import ast
import io
import re
import sys

BAR = "# " + "-" * 75
BANNER = re.compile(rf"{re.escape(BAR)}\n# ([^\n]*)\n{re.escape(BAR)}\n")


def dead_imports(path, src, tree):
    """Top-level imports with no reference anywhere in the file."""
    lines = src.split("\n")
    bound = []
    for node in tree.body:
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            line = lines[node.lineno - 1]
            for alias in node.names:
                name = alias.asname or alias.name.split(".")[0]
                bound.append((name, node.lineno, "noqa: F401" in line))

    used = {n.id for n in ast.walk(tree) if isinstance(n, ast.Name)}
    used |= {n.value.id for n in ast.walk(tree)
             if isinstance(n, ast.Attribute) and isinstance(n.value, ast.Name)}

    return [(name, lineno) for name, lineno, exempt in bound
            if not exempt and name not in used]


CODE = re.compile(r"^(def |class |@|if __name__)", re.M)


def orphan_banners(src):
    """Section banners with no code under them.

    The test is any top-level definition, not @mcp.tool specifically: msp_core
    heads its helper section with a banner and registers no tools, and reading
    that as orphaned would be wrong.
    """
    found = []
    for match in BANNER.finditer(src):
        after = src[match.end():]
        nxt = after.find(BAR)
        tail = after if nxt == -1 else after[:nxt]
        if not CODE.search(tail):
            line = src[:match.start()].count("\n") + 2
            found.append((match.group(1), line))
    return found


def check_file(path):
    src = io.open(path, encoding="utf-8").read()
    tree = ast.parse(src, filename=path)

    problems = []
    for name, lineno in dead_imports(path, src, tree):
        problems.append((lineno, f"{path}:{lineno}: dead import '{name}'"))
    for title, lineno in orphan_banners(src):
        problems.append((lineno, f"{path}:{lineno}: orphan banner '{title}'"))
    return problems


def main(argv):
    paths = argv[1:]
    if not paths:
        print("usage: check_unused.py FILE [FILE ...]", file=sys.stderr)
        return 2

    total = 0
    for path in paths:
        try:
            problems = check_file(path)
        except SyntaxError as exc:
            print(f"{path}: SYNTAX ERROR line {exc.lineno}: {exc.msg}")
            total += 1
            continue
        for _, message in sorted(problems):
            print(message)
        total += len(problems)

    print(f"UNUSED TOTAL: {total}")
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
