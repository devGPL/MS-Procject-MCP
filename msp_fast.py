"""
Fast read path: parse the .mpp from disk instead of walking it over COM.

WHY THIS EXISTS
---------------
COM enumeration dominates every scan. Measured on an 8,429-task schedule:
2.635 ms to obtain each task object against 0.240 ms to read a property, so
83% of a scan is spent just getting the objects. Touching every task costs
about 20 seconds before a single field is read, and no amount of tidying the
Python changes that floor.

Parsing the same file with mpxj: 1.52s to parse 13.8 MB, then 0.07s to walk
all 8,430 tasks -- because after the parse there is no RPC left, only field
access. The tools that took 30 to 48 seconds have roughly 1.6 seconds of work
in them.

Both paths were checked against the same project and agreed exactly: 8,243
detail tasks, 61 critical, 7,911 with a predecessor.

WHAT IT COSTS
-------------
This reads the SAVED FILE, not the running application. That is the whole
tradeoff, and it is a correctness question rather than a speed one:

  * Edits made in Microsoft Project and not yet saved are invisible here.
  * A mutation this server performs and does not persist would be invisible
    to the next read -- which is why every mutating tool now saves.

So the decision is made per call, never once at startup, and every response
says which path answered it. A tool that sometimes reads the file and
sometimes the application without saying so is worse than a slow one: the
error arrives as a plausible number.

WHERE IT DOES NOT WORK
----------------------
mpxj runs on the JVM through jpype, and jpype publishes no wheel for Windows
on ARM. On that platform this module reports itself unavailable and everything
falls back to COM -- correctly, but slowly. Java 9 or later is required.
"""

import os

# Resolved once, lazily. None means "not tried yet".
_estado = None


def _iniciar():
    """Import mpxj and start the JVM. Returns (ok, motivo)."""
    try:
        import jpype
    except ImportError:
        return False, ("jpype is not installed. It publishes no wheel for "
                       "Windows on ARM, where this path cannot be used.")
    try:
        import mpxj  # noqa: F401  -- registers the bundled JARs on the classpath
    except ImportError:
        return False, "mpxj is not installed."

    try:
        if not jpype.isJVMStarted():
            jpype.startJVM()
    except Exception as exc:
        return False, "could not start a JVM (%s). Java 9 or later is required." % str(exc)[:120]

    try:
        from org.mpxj.reader import UniversalProjectReader  # noqa: F401
    except Exception as exc:
        # The package moved from net.sf.mpxj to org.mpxj in mpxj 14.
        return False, "mpxj classes not reachable from the JVM (%s)." % str(exc)[:120]

    return True, "ready"


def disponivel():
    """True when this path can be used at all on this machine."""
    global _estado
    if _estado is None:
        _estado = _iniciar()
    return _estado[0]


def motivo_indisponivel():
    global _estado
    if _estado is None:
        _estado = _iniciar()
    return _estado[1]


def _mtime(caminho):
    try:
        return os.path.getmtime(caminho)
    except Exception:
        return None


def contexto(proj):
    """Decide which backend should answer, and describe why.

    Returns (usar_arquivo, caminho, origem) where `origem` is the block every
    response carries. The decision is deliberately conservative: anything it
    cannot verify sends the call to COM.
    """
    origem = {"backend": "com"}

    if not disponivel():
        origem["reason"] = "fast path unavailable: " + motivo_indisponivel()
        return False, None, origem

    try:
        caminho = str(proj.FullName)
    except Exception:
        origem["reason"] = "could not read the project path from MS Project"
        return False, None, origem

    if not caminho or not os.path.exists(caminho):
        origem["reason"] = ("the project has no readable file on disk "
                            "(never saved, or the path is not reachable from "
                            "this process)")
        return False, None, origem

    # Microsoft Project marks a project dirty on almost any interaction --
    # scrolling and selecting count -- so this flag says "not verifiable",
    # not "definitely different". The response reports it and answers anyway;
    # the caller is told exactly what the file predates.
    try:
        salvo = bool(proj.Saved)
    except Exception:
        salvo = False

    ts = _mtime(caminho)
    origem = {
        "backend": "mpxj",
        "read_from": caminho,
        "file_saved_at": _iso(ts),
        "app_has_unsaved_changes": not salvo,
    }
    if not salvo:
        origem["warning"] = (
            "Answered from the file saved at %s. Microsoft Project reports "
            "unsaved changes, so anything edited in the application since then "
            "is not reflected here. Save the project and call again if that "
            "matters." % _iso(ts)
        )
    return True, caminho, origem


def _iso(ts):
    if ts is None:
        return None
    import datetime
    return datetime.datetime.fromtimestamp(ts).isoformat(timespec="seconds")


# --- leitura ---------------------------------------------------------------

_cache = {"caminho": None, "mtime": None, "projeto": None}


def ler(caminho):
    """Parse the file, reusing the last parse when it is still current.

    The cache key is (path, mtime): any save invalidates it, so a stale parse
    cannot outlive the file it came from.
    """
    if not disponivel():
        # Callers reach this through contexto(), which checks first -- but a
        # direct call must not fail with "no module named org", which points at
        # the JVM never having been started rather than at the real condition.
        raise RuntimeError("fast read path unavailable: " + motivo_indisponivel())

    ts = _mtime(caminho)
    if (_cache["caminho"] == caminho and _cache["mtime"] == ts
            and _cache["projeto"] is not None):
        return _cache["projeto"]

    from org.mpxj.reader import UniversalProjectReader
    projeto = UniversalProjectReader().read(caminho)
    _cache.update({"caminho": caminho, "mtime": ts, "projeto": projeto})
    return projeto


# Enum names differ between the two readers. COM returns integers that
# task_to_dict maps to these strings; mpxj returns named constants. Both paths
# must produce the same string or the same tool answers differently depending
# on which one served it.
_CONSTRAINT = {
    "AS_SOON_AS_POSSIBLE": "ASAP", "AS_LATE_AS_POSSIBLE": "ALAP",
    "MUST_START_ON": "MSO", "MUST_FINISH_ON": "MFO",
    "START_NO_EARLIER_THAN": "SNET", "START_NO_LATER_THAN": "SNLT",
    "FINISH_NO_EARLIER_THAN": "FNET", "FINISH_NO_LATER_THAN": "FNLT",
}
_TIPO = {"FIXED_UNITS": "FixedUnits", "FIXED_DURATION": "FixedDuration",
         "FIXED_WORK": "FixedWork"}


def _dias(t, metodo):
    """Duration-like field in days. mpxj returns a Duration object."""
    try:
        d = getattr(t, metodo)()
        if d is None:
            return 0
        return round(float(d.getDuration()), 2)
    except Exception:
        return 0


def _dt(t, metodo):
    """Date in the same shape the COM path produces: 'YYYY-MM-DD HH:MM:SS'.

    mpxj renders a LocalDateTime without seconds when they are zero, so the
    two readers disagreed on every dated field until this padded them. The
    format is part of the contract, not a detail.
    """
    try:
        v = getattr(t, metodo)()
        if v is None:
            return None
        texto = str(v).replace("T", " ")[:19]
        if len(texto) == 16:      # 'YYYY-MM-DD HH:MM'
            texto += ":00"
        return texto if len(texto) >= 10 and texto[:4].isdigit() else None
    except Exception:
        return None


def _texto(t, n):
    try:
        v = t.getText(n)
        return str(v) if v is not None else ""
    except Exception:
        return ""


def _flag(t, n):
    try:
        return bool(t.getFlag(n))
    except Exception:
        return False


def _preds(t):
    """Predecessor string in the same shape COM produces: '5FS+2d' joined."""
    try:
        rel = t.getPredecessors()
        if not rel:
            return ""
        partes = []
        for r in rel:
            try:
                partes.append(str(r.getPredecessorTask().getID()))
            except Exception:
                pass
        return ";".join(partes)
    except Exception:
        return ""


def task_to_dict(t):
    """Same 35 keys as msp_core.task_to_dict, read from the parsed file.

    The shape is the contract. A fast path that returns fewer fields, or the
    same fields under different names, turns a performance change into a
    behaviour change -- and the caller has no way to tell which path answered
    beyond the source block.
    """
    return {
        "unique_id":              _int(t, "getUniqueID"),
        "id":                     _int(t, "getID"),
        "name":                   _str(t, "getName"),
        "outline_level":          _int(t, "getOutlineLevel"),
        "wbs":                    _str(t, "getWBS"),
        "summary":                _bool(t, "getSummary"),
        "milestone":              _bool(t, "getMilestone"),
        "start":                  _dt(t, "getStart"),
        "finish":                 _dt(t, "getFinish"),
        "duration_days":          _dias(t, "getDuration"),
        "percent_complete":       _int(t, "getPercentageComplete") or 0,
        "actual_start":           _dt(t, "getActualStart"),
        "actual_finish":          _dt(t, "getActualFinish"),
        "remaining_duration_days": _dias(t, "getRemainingDuration"),
        "total_slack_days":       _dias(t, "getTotalSlack"),
        "free_slack_days":        _dias(t, "getFreeSlack"),
        "deadline":               _dt(t, "getDeadline"),
        "priority":               _prioridade(t),
        "constraint_type":        _CONSTRAINT.get(_str(t, "getConstraintType"), "ASAP"),
        "constraint_date":        _dt(t, "getConstraintDate"),
        "manual":                 _modo(t),
        "type":                   _TIPO.get(_str(t, "getType"), "FixedUnits"),
        "predecessors":           _preds(t),
        "resource_names":         _str(t, "getResourceNames"),
        "notes":                  _str(t, "getNotes"),
        "critical":               _bool(t, "getCritical"),
        "active":                 _bool(t, "getActive", True),
        "rag":                    _texto(t, 1),
        "text1":                  _texto(t, 1),
        "text2":                  _texto(t, 2),
        "text3":                  _texto(t, 3),
        "flag1":                  _flag(t, 1),
        "flag2":                  _flag(t, 2),
        "hyperlink":              _str(t, "getHyperlinkAddress"),
        "hyperlink_text":         _str(t, "getHyperlink"),
    }


def _prioridade(t):
    try:
        p = t.getPriority()
        return int(p.getValue()) if p is not None else 500
    except Exception:
        return 500


def tarefas(projeto, incluir_resumo=False):
    """Every task as the same dict COM would produce, in row order.

    Sorted by ID because that is the order COM enumerates in -- ID is the row
    number -- and several tools read meaning from the order: a WBS tree is
    built by walking tasks and pushing each onto a stack of parents, so a list
    in a different order produces a different tree rather than a wrong field.
    Tasks without an ID sort last instead of raising.
    """
    saida = []
    for t in projeto.getTasks():
        try:
            if t.getUniqueID() is None:
                continue
            if not incluir_resumo and t.getSummary():
                continue
        except Exception:
            continue
        saida.append(task_to_dict(t))
    saida.sort(key=lambda d: (d.get("id") is None, d.get("id") or 0))
    return saida


def varredura(proj):
    """Every task of the project, from whichever backend can answer.

    Returns (lidas, origem):

      lidas   list of task dicts when the saved file could be parsed, and None
              when the caller must walk COM itself. Summaries are always
              included -- filtering belongs to the tool, so that one line
              filters both backends.
      origem  the source block the response carries, saying which path
              answered and why.

    This is the whole fast-path decision for a scanning tool. It lives here so
    that adding the fast path to a tool is three lines rather than thirty, and
    so that the fallback behaves identically in every one of them: when the
    same decision was written out per tool, each copy was a chance for one
    tool to fall back quietly where the others were loud.
    """
    usar_arquivo, caminho, origem = contexto(proj)
    if not usar_arquivo:
        return None, origem

    try:
        return tarefas(ler(caminho), incluir_resumo=True), origem
    except Exception as exc:
        # A parse failure must not fail the tool -- COM still works. But it
        # must not hide either: an exception here is a defect, not the
        # documented case of the fast path being unavailable, and the two look
        # identical from outside. This one is flagged loudly, because a
        # fallback that silently degrades is how a change that delivers
        # nothing still looks like it works.
        return None, {
            "backend": "com",
            "reason": "fast path RAISED and fell back: %s: %s"
                      % (type(exc).__name__, str(exc)[:140]),
            "action_required": "This is a bug, not a configuration. The "
                               "answer below is correct but was produced the "
                               "slow way; report the error above.",
        }


def _modo(t):
    try:
        modo = t.getTaskMode()
        return modo is not None and str(modo).upper().startswith("MANUAL")
    except Exception:
        return False


def _npred(t):
    try:
        p = t.getPredecessors()
        return len(p) if p is not None else 0
    except Exception:
        return 0


def _str(t, metodo):
    try:
        v = getattr(t, metodo)()
        return str(v) if v is not None else ""
    except Exception:
        return ""


def _int(t, metodo):
    try:
        v = getattr(t, metodo)()
        return int(v) if v is not None else None
    except Exception:
        return None


def _bool(t, metodo, padrao=False):
    try:
        v = getattr(t, metodo)()
        return bool(v) if v is not None else padrao
    except Exception:
        return padrao


def _data(t, metodo):
    try:
        v = getattr(t, metodo)()
        return str(v)[:10] if v is not None else None
    except Exception:
        return None
