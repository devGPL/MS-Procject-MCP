"""
COM boundary and shared helpers for the MS Project MCP server.

This module owns the connection to Microsoft Project, the conversion of COM
objects into plain dicts, and the FastMCP instance every tool registers on.
Tool modules import from here; this module imports from none of them.

Two invariants hold this file together, and breaking either one breaks the
whole verification story:

1. `import win32com.client` stays INSIDE get_app(). Every COM import in this
   project is deferred, which is why the module imports cleanly on macOS and
   Linux with no pywin32 present, and why the gates in tools/ can check tool
   registration without Windows.
2. This is the COM boundary, not a place for business logic. What belongs to
   a domain belongs in that domain's module.

Known divergence, unverified on Windows: the getattr() calls in this project
apply their default only for AttributeError. A Python fake raises exactly
that, while real COM may raise pywintypes.com_error -- so an offline test can
pass where production fails. Confirm on Windows before trusting any offline
suite as proof of COM behaviour.
"""

import contextlib
import datetime
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("MS Project")

# Reverse maps for COM integer enums. Module scope rather than rebuilt inside
# task_to_dict, which runs once per task in every listing.
CONSTRAINT_NAMES = {
    0: "ASAP", 1: "ALAP", 2: "MSO", 3: "MFO",
    4: "SNET", 5: "SNLT", 6: "FNET", 7: "FNLT",
}
TASK_TYPE_NAMES = {0: "FixedUnits", 1: "FixedDuration", 2: "FixedWork"}

# pjTimescaleUnit values accepted by the timephased APIs.
TIMESCALE_MAP = {"daily": 3, "weekly": 4, "monthly": 5}


# ---------------------------------------------------------------------------
# COM helpers
# ---------------------------------------------------------------------------


# COM errors that mean "busy, ask again" rather than "broken".
#
#   RPC_E_CALL_REJECTED      0x80010001  the app refused the call
#   RPC_E_SERVERCALL_RETRYLATER 0x8001010A  the app asked us to wait
#
# Microsoft Project raises these while it is busy -- recalculating, showing a
# modal dialog, or still settling after a large operation. They are transient
# and unrelated to whether the application is reachable, but they surface as a
# raw Windows-language exception with no context, on a project that is open and
# working.
_COM_BUSY = (-2147418111, -2147417846)


def _com_retry(chamada, tentativas=4, espera=0.4):
    """Run a COM call, retrying while the application says it is busy.

    Backs off geometrically: 0.4s, 0.8s, 1.6s -- about 2.8s total before giving
    up. Long enough to outlast a recalculation, short enough that a genuinely
    unreachable application still fails quickly.

    Only busy errors are retried. Anything else propagates immediately, because
    repeating a call that failed for a real reason just delays the report.
    """
    import time
    ultima = None
    for tentativa in range(tentativas):
        try:
            return chamada()
        except Exception as exc:
            codigo = getattr(exc, "hresult", None)
            if codigo is None:
                args = getattr(exc, "args", ())
                codigo = args[0] if args and isinstance(args[0], int) else None
            if codigo not in _COM_BUSY:
                raise
            ultima = exc
            if tentativa < tentativas - 1:
                time.sleep(espera * (2 ** tentativa))
    raise ultima



@contextlib.contextmanager
def calculo_suspenso(app):
    """Suspend Microsoft Project's recalculation for the duration of a batch.

    Changing a task normally triggers a recalculation of the whole dependency
    network. Doing that once per item turns a batch of N changes over M tasks
    into N recalculations -- quadratic, and it dominates everything else.

    Measured: converting 8,243 tasks from manual to automatic scheduling ran
    for tens of minutes at 100% of a core, because each conversion rescheduled
    the network the previous ones had just built.

    Restores the setting in a finally, so an exception mid-batch cannot leave
    the application with calculation switched off -- which would silently stop
    it updating for the user afterwards.
    """
    suspendeu = False
    try:
        app.Calculation = 0
        suspendeu = True
    except Exception:
        pass
    try:
        yield
    finally:
        if suspendeu:
            try:
                app.Calculation = -1
            except Exception:
                pass



def get_app(require_project=True):
    """Get the running MS Project instance. Raises if it cannot be reached."""
    import win32com.client
    try:
        app = _com_retry(
            lambda: win32com.client.GetActiveObject("MSProject.Application"))
    except Exception as exc:
        # "Not running" is the most likely cause but not the only one, and
        # reporting it as fact sends people to check something already true.
        #
        # The Running Object Table is partitioned two ways, and both partitions
        # are invisible from the outside:
        #
        #   By logon session. A server started over SSH lands in session 0 and
        #   cannot see an MS Project on the interactive desktop in session 1.
        #
        #   By integrity level. A server launched from an elevated shell runs
        #   High while MS Project runs Medium, and High does not see Medium's
        #   registrations. Same desktop, same user, same session -- separate
        #   tables. Running as administrator makes this WORSE, not better,
        #   which is the opposite of what anyone debugging it will try.
        raise RuntimeError(
            "Could not attach to MS Project (" + type(exc).__name__ + ": "
            + str(exc)[:160] + "). Check, in order: MS Project is running with "
            "a project open; this server was NOT started from an elevated "
            "shell -- 'run as administrator' hides a normally-launched MS "
            "Project rather than helping; and this server runs in the same "
            "Windows logon session as MS Project, which rules out starting it "
            "over SSH."
        )
    # A busy application usually rejects the first real call rather than the
    # attach, so this count doubles as the responsiveness check.
    if require_project and _com_retry(lambda: app.Projects.Count) == 0:
        raise RuntimeError(
            "No project file is open in MS Project. Please open a file first."
        )
    return app


def get_proj(app):
    return app.ActiveProject


def _get_mpd(proj):
    """Get MinutesPerDay safely (falls back to 480 for freshly-imported XML)."""
    try:
        return proj.MinutesPerDay
    except Exception:
        return 480


def _parse_date(s):
    """Parse YYYY-MM-DD string to datetime for COM. Returns None if empty."""
    if not s:
        return None
    return datetime.datetime.strptime(s, "%Y-%m-%d")


def _fmt_com(dt):
    """Format a COM date, or None when there is no date.

    MS Project returns a locale-dependent sentinel for an empty date --
    "ND" on a Portuguese install, "NA" on an English one -- and the old
    code passed it straight through, so a task with no deadline came back
    as {"deadline": "ND"}. A client reading that sees a value where there
    is none, and the string it sees depends on the language of the machine
    the server happens to run on.

    Anything that does not start with a four-digit year is treated as
    "no date", which covers both sentinels without hardcoding either.
    """
    try:
        if dt is None:
            return None
        texto = str(dt)[:19]
        if len(texto) < 10 or not texto[:4].isdigit():
            return None
        return texto
    except Exception:
        return None


def _safe(read, default=None):
    """Read a COM property defensively.

    `read` must be a callable. Passing the property itself evaluates it in
    the caller's frame, before this function runs, so the except clause
    never sees the failure and `default` is never applied.

    A None result is treated like a failed read: COM returns None for
    properties that do not apply to a task type, and callers divide by
    these values.
    """
    try:
        value = read()
    except Exception:
        return default
    return default if value is None else value


# The 35 task fields, one reader each, keyed by the name they are published
# under. This table is the single definition of the task shape on the COM
# side: task_to_dict materialises every entry, VistaCOM evaluates one entry
# at a time on demand, and both therefore agree by construction. When they
# were two separate bodies, a field added to one was a field missing from the
# other -- and the caller could not tell, because a missing key reads as an
# empty value rather than as an error.
#
# Which readers are wrapped in _safe() is deliberate and unchanged: the
# unguarded ones are the properties every task has, and a failure there is a
# broken COM reference rather than a field that does not apply.
_LEITORES_TAREFA = {
    "unique_id":              lambda t, mpd: t.UniqueID,
    "id":                     lambda t, mpd: t.ID,
    "name":                   lambda t, mpd: t.Name,
    "outline_level":          lambda t, mpd: t.OutlineLevel,
    "wbs":                    lambda t, mpd: t.WBS,
    "summary":                lambda t, mpd: bool(t.Summary),
    "milestone":              lambda t, mpd: bool(t.Milestone),
    "start":                  lambda t, mpd: _fmt_com(t.Start),
    "finish":                 lambda t, mpd: _fmt_com(t.Finish),
    "duration_days":          lambda t, mpd: round(t.Duration / mpd, 2) if t.Duration else 0,
    "percent_complete":       lambda t, mpd: t.PercentComplete,
    "actual_start":           lambda t, mpd: _fmt_com(_safe(lambda: t.ActualStart)),
    "actual_finish":          lambda t, mpd: _fmt_com(_safe(lambda: t.ActualFinish)),
    "remaining_duration_days": lambda t, mpd: round(_safe(lambda: t.RemainingDuration, 0) / mpd, 2),
    "total_slack_days":       lambda t, mpd: round(_safe(lambda: t.TotalSlack, 0) / mpd, 2),
    "free_slack_days":        lambda t, mpd: round(_safe(lambda: t.FreeSlack, 0) / mpd, 2),
    "deadline":               lambda t, mpd: _fmt_com(_safe(lambda: t.Deadline)),
    "priority":               lambda t, mpd: _safe(lambda: t.Priority, 500),
    "constraint_type":        lambda t, mpd: CONSTRAINT_NAMES.get(_safe(lambda: t.ConstraintType, 0), "ASAP"),
    "constraint_date":        lambda t, mpd: _fmt_com(_safe(lambda: t.ConstraintDate)),
    "manual":                 lambda t, mpd: bool(_safe(lambda: t.Manual, False)),
    "type":                   lambda t, mpd: TASK_TYPE_NAMES.get(_safe(lambda: t.Type, 0), "FixedUnits"),
    "predecessors":           lambda t, mpd: t.Predecessors,
    "resource_names":         lambda t, mpd: t.ResourceNames,
    "notes":                  lambda t, mpd: t.Notes,
    "critical":               lambda t, mpd: bool(t.Critical),
    "active":                 lambda t, mpd: bool(t.Active),
    "rag":                    lambda t, mpd: t.Text1 or "",
    "text1":                  lambda t, mpd: t.Text1 or "",
    "text2":                  lambda t, mpd: t.Text2 or "",
    "text3":                  lambda t, mpd: t.Text3 or "",
    "flag1":                  lambda t, mpd: bool(t.Flag1),
    "flag2":                  lambda t, mpd: bool(t.Flag2),
    "hyperlink":              lambda t, mpd: _safe(lambda: t.HyperlinkAddress, "") or "",
    "hyperlink_text":         lambda t, mpd: _safe(lambda: t.Hyperlink, "") or "",
}

CAMPOS_TAREFA = tuple(_LEITORES_TAREFA)


def task_to_dict(t, mpd):
    """Convert a COM Task object to a plain dict.

    Takes MinutesPerDay rather than the project: reading it here meant one COM
    round trip per task in every listing.

    Costs one COM read per field -- 35 of them. That is the right trade when
    the caller publishes the whole task and the wrong one when it publishes
    three fields of it, which is what VistaCOM below exists for.
    """
    return {chave: ler(t, mpd) for chave, ler in _LEITORES_TAREFA.items()}


class VistaCOM:
    """One COM task, read through the key names task_to_dict publishes.

    Exists so a tool can be written ONCE against key names and still be served
    by either backend: the fast path hands it dicts parsed from the .mpp, this
    hands it COM tasks, and the body cannot tell them apart.

    The alternative -- calling task_to_dict on the COM side to get dicts
    everywhere -- would have cost 35 property reads per task where the tool
    needs three. Measured: 0.240 ms per property against 2.635 ms to obtain
    the task object, so materialising every field turns a scan of 8,429 tasks
    from about 32 seconds into about 92. The fast path is unavailable on
    Windows on ARM, so that regression would land on the only path that runs
    there.

    Reads on demand and caches, because a filter asking the same field twice
    is one COM round trip, not two.
    """

    __slots__ = ("com", "_mpd", "_lido")

    def __init__(self, tarefa, mpd):
        self.com = tarefa
        self._mpd = mpd
        self._lido = {}

    def __getitem__(self, chave):
        if chave in self._lido:
            return self._lido[chave]
        ler = _LEITORES_TAREFA.get(chave)
        if ler is None:
            # Same failure a dict gives, so a tool asking for a field that does
            # not exist fails identically on both backends instead of only on
            # one of them.
            raise KeyError(chave)
        valor = ler(self.com, self._mpd)
        self._lido[chave] = valor
        return valor

    def get(self, chave, padrao=None):
        try:
            return self[chave]
        except KeyError:
            return padrao

    def completo(self):
        """Every field, as task_to_dict would return it."""
        return task_to_dict(self.com, self._mpd)


def vista_com(proj, mpd):
    """Walk proj.Tasks yielding one VistaCOM each -- the COM half of a scan.

    Yields every task including summaries. Filtering belongs to the tool, so
    that the same line filters both backends; a filter applied here would
    apply to one of them only.
    """
    for t in proj.Tasks:
        if t is None:
            continue
        yield VistaCOM(t, mpd)


def tarefa_completa(vista):
    """The full 35-field dict for either kind of view.

    The fast path already holds one, so this is the identity there and 35 COM
    reads on the other side -- paid only for the tasks a tool actually
    publishes, never for the ones it filtered out.
    """
    return vista if isinstance(vista, dict) else vista.completo()


def _count_resources(proj):
    """Count resources safely — returns 0 if resource pool is empty/inaccessible."""
    try:
        return sum(1 for r in proj.Resources if r is not None)
    except Exception:
        return 0


def _fmt_date(dt):
    """Format a COM date to 'YYYY-MM-DD' string. Returns None on failure."""
    try:
        return str(dt)[:10] if dt else None
    except Exception:
        return None


def _to_naive(dt):
    """Strip timezone from COM datetime for safe comparison with datetime.now()."""
    if dt is None:
        return None
    try:
        if hasattr(dt, 'tzinfo') and dt.tzinfo is not None:
            return dt.replace(tzinfo=None)
    except Exception:
        pass
    return dt



def _dia(valor):
    """The 'YYYY-MM-DD' half of the date a task view publishes.

    Views carry 'YYYY-MM-DD HH:MM:SS' because that is what task_to_dict has
    always published. Tools that summarise -- a WBS roll-up, a slack listing --
    published the short form via _fmt_date, and shortening here keeps their
    output byte-identical to what they returned before they read from a view.
    """
    return valor[:10] if valor else None


def _dt_de_vista(valor):
    """Parse a view's date string back to a datetime, or None.

    Both backends publish dates as text, so a tool that compares against
    datetime.now() has to parse rather than reach for the COM value -- which
    exists on one side only.
    """
    if not valor:
        return None
    try:
        return datetime.datetime.strptime(valor[:19], "%Y-%m-%d %H:%M:%S")
    except Exception:
        pass
    try:
        return datetime.datetime.strptime(valor[:10], "%Y-%m-%d")
    except Exception:
        return None


def _calendar_names(proj):
    """Names of every base calendar, for validation and error messages.

    Returns [] when the collection cannot be read, which is indistinguishable
    from a project that has no calendars. Callers report "not found" either
    way, so a COM failure currently surfaces as a missing calendar. That is
    pre-existing behaviour, preserved here deliberately -- centralising it is
    what makes the diagnosis fixable in one place later.
    """
    names = []
    try:
        for cal in proj.BaseCalendars:
            if cal is not None:
                names.append(str(cal.Name))
    except Exception:
        pass
    return names


def _find_calendar(proj, name):
    """Return the base calendar matching `name`, case-insensitively.

    MS Project treats calendar names case-insensitively. Matching them
    byte-for-byte is what made set_calendar_exception("standard") fail on the
    same project where set_working_hours("standard") worked.
    """
    if not name:
        return None
    target = str(name).lower()
    try:
        for cal in proj.BaseCalendars:
            if cal is not None and str(cal.Name).lower() == target:
                return cal
    except Exception:
        return None
    return None



def _find_resource(proj, name):
    """Return the resource matching `name`, case-insensitively.

    A resource with no name is skipped rather than raising. Four of the eight
    inline lookups this replaces called r.Name.lower() with no guard, so one
    nameless resource anywhere in the pool made them fail with AttributeError
    -- and which four had the guard was arbitrary.
    """
    if not name:
        return None
    target = str(name).lower()
    try:
        for r in proj.Resources:
            if r is not None and r.Name and str(r.Name).lower() == target:
                return r
    except Exception:
        return None
    return None



def _uid_map(proj):
    """Snapshot of {UniqueID: task} for the whole project.

    For bulk operations: _find_task rescans proj.Tasks per lookup, so N items
    against M tasks costs N*M COM traversals. This costs one.

    It is a SNAPSHOT, and that difference is observable. _find_task re-reads
    the collection on every call, so a task deleted earlier in the same bulk
    run disappears from later lookups; entries in this map stay, and the stale
    COM reference fails on use instead of reporting "not found".
    """
    try:
        return {t.UniqueID: t for t in proj.Tasks if t is not None}
    except Exception:
        return {}


def _uid_to_id_map(proj):
    """Snapshot of {UniqueID: ID} -- the row numbers used by link syntax."""
    try:
        return {t.UniqueID: t.ID for t in proj.Tasks if t is not None}
    except Exception:
        return {}


def _find_task(proj, unique_id):
    """Find a task by UniqueID. Returns the COM Task object or None."""
    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            return t
    return None


def _custom_field_id(field_name):
    """
    Map field name like 'Text5', 'Number1', 'Flag3', 'Date1', 'Duration2'
    to the COM pjCustomTask* field ID constant.
    Returns (field_id, field_type) or raises ValueError.
    """
    name = field_name.strip()
    lower = name.lower()

    # COM field ID bases (pjCustomTask* constants)
    bases = {
        "text":     (188743731, 30),   # Text1-30
        "number":   (188743767, 20),   # Number1-20
        "date":     (188743945, 10),   # Date1-10
        "flag":     (188743752, 20),   # Flag1-20
        "duration": (188743783, 10),   # Duration1-10
    }

    for prefix, (base_id, max_n) in bases.items():
        if lower.startswith(prefix):
            num_str = lower[len(prefix):]
            if num_str.isdigit():
                num = int(num_str)
                if 1 <= num <= max_n:
                    return base_id + (num - 1), prefix
    raise ValueError(f"Unknown custom field: '{field_name}'. Use Text1-30, Number1-20, Date1-10, Flag1-20, Duration1-10.")
