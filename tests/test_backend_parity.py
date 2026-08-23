"""Parity between the two read paths: COM and the parsed .mpp.

A fast path that disagrees with the slow one is worse than no fast path,
because the disagreement arrives as a plausible number. This compares the two
field by field on the same project and fails on any difference that is not
explicitly allowed.

Requires Windows, Microsoft Project open with a SAVED project, and the fast
path available (mpxj + jpype + Java 9+). Skips cleanly when any of that is
missing rather than reporting a false pass.

    python tests/test_backend_parity.py [n_tarefas]
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import msp_core
import msp_fast

# Fields where the two readers legitimately differ, with the reason. Anything
# not listed here must match exactly.
TOLERADAS = {
    # COM reports the resource names it has assigned; mpxj returns None when
    # the field was never materialised in the file.
    "resource_names": "empty vs None on unassigned tasks",
    # COM's predecessor string carries link type and lag ('5FS+2d'); the file
    # exposes the relations as objects and this reader joins ids only.
    "predecessors": "COM includes link type and lag, the parser does not",
}


def comparar(a, b):
    """Return the list of differing keys, ignoring the tolerated ones."""
    difs = []
    for k in sorted(set(a) | set(b)):
        if k in TOLERADAS:
            continue
        va, vb = a.get(k), b.get(k)
        if va == vb:
            continue
        # Empty string and None mean the same thing across the two readers.
        if (va in ("", None)) and (vb in ("", None)):
            continue
        # Durations: tolerate float noise at the second decimal.
        if isinstance(va, float) or isinstance(vb, float):
            try:
                if abs(float(va or 0) - float(vb or 0)) < 0.02:
                    continue
            except Exception:
                pass
        difs.append((k, va, vb))
    return difs


def main():
    limite = int(sys.argv[1]) if len(sys.argv) > 1 else 300

    if not msp_fast.disponivel():
        print("SKIP: fast path unavailable -- %s" % msp_fast.motivo_indisponivel())
        return 0

    try:
        app = msp_core.get_app()
        proj = msp_core.get_proj(app)
    except Exception as exc:
        print("SKIP: no MS Project to compare against (%s)" % str(exc)[:120])
        return 0

    try:
        if not bool(proj.Saved):
            print("SKIP: the project has unsaved changes, so the file and the")
            print("      application are legitimately different. Save and rerun.")
            return 0
    except Exception:
        pass

    caminho = str(proj.FullName)
    print("projeto: %s" % proj.Name)
    print("arquivo: %s" % caminho)
    print("comparando ate %d tarefas...\n" % limite)

    do_arquivo = {d["unique_id"]: d
                  for d in msp_fast.tarefas(msp_fast.ler(caminho))}

    mpd = msp_core._get_mpd(proj)
    vistas = 0
    divergentes = 0
    ausentes = 0
    contagem = {}

    for t in proj.Tasks:
        if t is None:
            continue
        try:
            if t.Summary:
                continue
        except Exception:
            continue
        vistas += 1
        if vistas > limite:
            break

        try:
            do_com = msp_core.task_to_dict(t, mpd)
        except Exception as exc:
            print("  COM falhou na tarefa %d: %s" % (vistas, str(exc)[:80]))
            divergentes += 1
            continue

        par = do_arquivo.get(do_com["unique_id"])
        if par is None:
            ausentes += 1
            continue

        difs = comparar(do_com, par)
        if difs:
            divergentes += 1
            if divergentes <= 5:
                print("  uid=%s '%s'" % (do_com["unique_id"], str(do_com["name"])[:34]))
                for k, va, vb in difs[:6]:
                    print("     %-22s COM=%-22r arquivo=%r" % (k, str(va)[:20], str(vb)[:20]))
            for k, _, _ in difs:
                contagem[k] = contagem.get(k, 0) + 1

    print("")
    print("  comparadas          %d" % min(vistas, limite))
    print("  ausentes no arquivo %d" % ausentes)
    print("  divergentes         %d" % divergentes)
    if contagem:
        print("\n  campos que divergem:")
        for k, n in sorted(contagem.items(), key=lambda x: -x[1]):
            print("    %-24s %d" % (k, n))
    if TOLERADAS:
        print("\n  ignorados de proposito:")
        for k, motivo in TOLERADAS.items():
            print("    %-24s %s" % (k, motivo))

    ok = divergentes == 0 and ausentes == 0
    print("\n  %s" % ("PARIDADE OK" if ok else "DIVERGENCIA -- o caminho rapido nao pode ser confiado"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
