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

def get_app(require_project=True):
    """Get running MS Project instance. Raises if not running."""
    import win32com.client
    try:
        app = win32com.client.GetActiveObject("MSProject.Application")
    except Exception:
        raise RuntimeError(
            "MS Project is not running. Open MS Project and load a file first."
        )
    if require_project and app.Projects.Count == 0:
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


def task_to_dict(t, mpd):
    """Convert a COM Task object to a plain dict.

    Takes MinutesPerDay rather than the project: reading it here meant one COM
    round trip per task in every listing.
    """

    def fmt(dt):
        try:
            if dt is None:
                return None
            return str(dt)[:19]
        except Exception:
            return None

    def safe(read, default=None):
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

    return {
        "unique_id":              t.UniqueID,
        "id":                     t.ID,
        "name":                   t.Name,
        "outline_level":          t.OutlineLevel,
        "wbs":                    t.WBS,
        "summary":                bool(t.Summary),
        "milestone":              bool(t.Milestone),
        "start":                  fmt(t.Start),
        "finish":                 fmt(t.Finish),
        "duration_days":          round(t.Duration / mpd, 2) if t.Duration else 0,
        "percent_complete":       t.PercentComplete,
        "actual_start":           fmt(safe(lambda: t.ActualStart)),
        "actual_finish":          fmt(safe(lambda: t.ActualFinish)),
        "remaining_duration_days": round(safe(lambda: t.RemainingDuration, 0) / mpd, 2),
        "total_slack_days":       round(safe(lambda: t.TotalSlack, 0) / mpd, 2),
        "free_slack_days":        round(safe(lambda: t.FreeSlack, 0) / mpd, 2),
        "deadline":               fmt(safe(lambda: t.Deadline)),
        "priority":               safe(lambda: t.Priority, 500),
        "constraint_type":        CONSTRAINT_NAMES.get(safe(lambda: t.ConstraintType, 0), "ASAP"),
        "constraint_date":        fmt(safe(lambda: t.ConstraintDate)),
        "manual":                 bool(safe(lambda: t.Manual, False)),
        "type":                   TASK_TYPE_NAMES.get(safe(lambda: t.Type, 0), "FixedUnits"),
        "predecessors":           t.Predecessors,
        "resource_names":         t.ResourceNames,
        "notes":                  t.Notes,
        "critical":               bool(t.Critical),
        "active":                 bool(t.Active),
        "rag":                    t.Text1 or "",
        "text1":                  t.Text1 or "",
        "text2":                  t.Text2 or "",
        "text3":                  t.Text3 or "",
        "flag1":                  bool(t.Flag1),
        "flag2":                  bool(t.Flag2),
        "hyperlink":              safe(lambda: t.HyperlinkAddress, "") or "",
        "hyperlink_text":         safe(lambda: t.Hyperlink, "") or "",
    }


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
