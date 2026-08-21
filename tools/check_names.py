#!/usr/bin/env python3
"""
Gate 1 — unresolved-name checker (pure stdlib, AST only).

Catches the failure mode that splitting server.py into modules will produce:
a helper that stays behind in another module and is never imported. Python
raises NameError for that only when the line runs, and server.py wraps most
calls in bare `except Exception:` — so the symptom surfaces as a silent
{"error": ...} on Windows, weeks later, instead of an ImportError at startup.

Reports every bare Name loaded in a scope where nothing binds it. Attribute
access (obj.attr) is not checked: only the root name is.

Usage:
    python3 tools/check_names.py FILE [FILE ...]

Exit code 0 when nothing is unresolved, 1 otherwise.

Deliberately not pyflakes/ruff: this repo has no linter installed and no
pyproject, and the whole point of this gate is that it runs anywhere with
nothing to install.
"""

import ast
import builtins
import sys

BUILTINS = set(dir(builtins)) | {"__file__", "__name__", "__doc__", "__spec__"}

# Nodes that open a new scope for the names bound inside their body.
SCOPED = (ast.FunctionDef, ast.AsyncFunctionDef, ast.Lambda, ast.ClassDef)


class Scope:
    def __init__(self, parent=None, kind="module"):
        self.names = set()
        self.parent = parent
        self.kind = kind

    def bind(self, name):
        if name:
            self.names.add(name)

    def resolves(self, name):
        scope = self
        while scope is not None:
            if name in scope.names:
                return True
            # A class body's names are not visible to functions nested in it.
            scope = scope.parent
        return name in BUILTINS


def _bind_target(target, scope):
    """Bind every name a assignment/for/with/comprehension target introduces."""
    for node in ast.walk(target):
        if isinstance(node, ast.Name):
            scope.bind(node.id)


def _bind_args(args, scope):
    for group in (args.posonlyargs, args.args, args.kwonlyargs):
        for arg in group:
            scope.bind(arg.arg)
    for extra in (args.vararg, args.kwarg):
        if extra is not None:
            scope.bind(extra.arg)


def collect_bindings(body, scope):
    """Bind names declared anywhere in this scope, without descending into
    nested scopes (their bodies get their own Scope). A name assigned anywhere
    in a function is local to the entire function, so this must run before
    any load is checked."""
    stack = list(body)
    while stack:
        node = stack.pop()

        if isinstance(node, SCOPED):
            # The nested definition's *name* belongs to this scope; its body
            # does not. Lambdas bind no name.
            if not isinstance(node, ast.Lambda):
                scope.bind(node.name)
            continue

        if isinstance(node, (ast.Assign, ast.AugAssign, ast.AnnAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            for target in targets:
                _bind_target(target, scope)
        elif isinstance(node, (ast.For, ast.AsyncFor)):
            _bind_target(node.target, scope)
        elif isinstance(node, (ast.With, ast.AsyncWith)):
            for item in node.items:
                if item.optional_vars is not None:
                    _bind_target(item.optional_vars, scope)
        elif isinstance(node, ast.ExceptHandler):
            scope.bind(node.name)
        elif isinstance(node, (ast.Import, ast.ImportFrom)):
            for alias in node.names:
                scope.bind(alias.asname or alias.name.split(".")[0])
        elif isinstance(node, (ast.Global, ast.Nonlocal)):
            for name in node.names:
                scope.bind(name)
        elif isinstance(node, ast.NamedExpr):
            _bind_target(node.target, scope)
        elif isinstance(node, (ast.ListComp, ast.SetComp, ast.DictComp, ast.GeneratorExp)):
            # Comprehensions have their own scope in Python 3; binding their
            # targets here instead is an over-approximation that can only
            # suppress findings, never invent one.
            for gen in node.generators:
                _bind_target(gen.target, scope)

        for child in ast.iter_child_nodes(node):
            stack.append(child)


def check_scope(body, scope, path, findings, args_node=None):
    """Walk this scope's body, reporting unresolved loads and recursing into
    nested scopes."""
    if args_node is not None:
        _bind_args(args_node, scope)
    collect_bindings(body, scope)

    stack = list(body)
    while stack:
        node = stack.pop()

        if isinstance(node, SCOPED):
            # Decorators, defaults and annotations evaluate in THIS scope.
            for deco in getattr(node, "decorator_list", []):
                stack.append(deco)
            if isinstance(node, ast.ClassDef):
                stack.extend(node.bases)
                child = Scope(scope, "class")
                check_scope(node.body, child, path, findings)
            else:
                args = node.args
                for default in list(args.defaults) + [d for d in args.kw_defaults if d]:
                    stack.append(default)
                child = Scope(scope, "function")
                sub_body = node.body if not isinstance(node, ast.Lambda) else [ast.Expr(node.body)]
                check_scope(sub_body, child, path, findings, args_node=args)
            continue

        if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load):
            if not scope.resolves(node.id):
                findings.append((path, node.lineno, node.col_offset, node.id))
            continue

        for child in ast.iter_child_nodes(node):
            stack.append(child)


def check_file(path):
    with open(path, "r", encoding="utf-8") as fp:
        tree = ast.parse(fp.read(), filename=path)
    findings = []
    check_scope(tree.body, Scope(None, "module"), path, findings)
    return findings


def main(argv):
    paths = argv[1:]
    if not paths:
        print("usage: check_names.py FILE [FILE ...]", file=sys.stderr)
        return 2

    total = 0
    for path in paths:
        try:
            findings = check_file(path)
        except SyntaxError as exc:
            print(f"{path}: SYNTAX ERROR line {exc.lineno}: {exc.msg}")
            total += 1
            continue
        for fpath, line, col, name in sorted(findings, key=lambda f: (f[1], f[2])):
            print(f"{fpath}:{line}:{col}: unresolved name '{name}'")
        total += len(findings)

    print(f"UNRESOLVED TOTAL: {total}")
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
