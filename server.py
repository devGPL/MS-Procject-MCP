"""
MS Project MCP Server
Controls local Microsoft Project via COM automation.
Install: pip install mcp
Run:     python server.py
Register in claude_desktop_config.json (see bottom of file).
"""

import json
import sys
import datetime

from msp_core import (
    mcp,
    get_app,
    get_proj,
    task_to_dict,
    _find_task,
    _fmt_date,
    _get_mpd,
    _parse_date,
    _to_naive,
)

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


# ---------------------------------------------------------------------------
# TOOLS — Reading tasks
# ---------------------------------------------------------------------------

@mcp.tool()
def get_tasks(
    include_summary: bool = False,
    outline_level: int = 0,
    keyword: str = ""
) -> str:
    """
    Get all tasks from the active project.

    Args:
        include_summary: Include summary/parent tasks (default False).
        outline_level:   Filter to a specific outline level (0 = all).
        keyword:         Filter tasks whose name contains this string (case-insensitive).
    """
    app  = get_app()
    proj = get_proj(app)

    results = []
    for t in proj.Tasks:
        if t is None:
            continue
        if not include_summary and t.Summary:
            continue
        if outline_level > 0 and t.OutlineLevel != outline_level:
            continue
        if keyword and keyword.lower() not in t.Name.lower():
            continue
        results.append(task_to_dict(t, proj))

    return json.dumps({"count": len(results), "tasks": results}, indent=2)


@mcp.tool()
def get_task(unique_id: int) -> str:
    """Get full details for a single task by its UniqueID."""
    app  = get_app()
    proj = get_proj(app)

    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            return json.dumps(task_to_dict(t, proj), indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


@mcp.tool()
def get_critical_path() -> str:
    """Return all tasks on the critical path (non-summary)."""
    app  = get_app()
    proj = get_proj(app)

    results = []
    for t in proj.Tasks:
        if t is not None and t.Critical and not t.Summary:
            results.append(task_to_dict(t, proj))

    return json.dumps({"count": len(results), "tasks": results}, indent=2)


@mcp.tool()
def get_tasks_by_rag(rag: str = "Red") -> str:
    """
    Return tasks filtered by RAG status stored in the Text1 custom field.
    rag: 'Red', 'Amber', or 'Green'
    """
    app  = get_app()
    proj = get_proj(app)

    results = []
    for t in proj.Tasks:
        if t is not None and not t.Summary:
            if (t.Text1 or "").strip().lower() == rag.strip().lower():
                results.append(task_to_dict(t, proj))

    return json.dumps({"rag": rag, "count": len(results), "tasks": results}, indent=2)


@mcp.tool()
def get_overdue_tasks() -> str:
    """Return incomplete tasks whose Finish date is in the past."""
    import datetime
    today = datetime.datetime.now()
    app   = get_app()
    proj  = get_proj(app)

    results = []
    for t in proj.Tasks:
        if t is None or t.Summary or t.Milestone:
            continue
        if t.PercentComplete >= 100:
            continue
        try:
            finish = _to_naive(t.Finish)
            if finish and finish < today:
                results.append(task_to_dict(t, proj))
        except Exception:
            continue

    return json.dumps({"count": len(results), "tasks": results}, indent=2)


@mcp.tool()
def get_tasks_by_resource(resource_name: str) -> str:
    """Return all tasks assigned to a named resource (case-insensitive substring match)."""
    app  = get_app()
    proj = get_proj(app)

    results = []
    name_lower = resource_name.lower()
    for t in proj.Tasks:
        if t is not None and not t.Summary:
            if name_lower in (t.ResourceNames or "").lower():
                results.append(task_to_dict(t, proj))

    return json.dumps({
        "resource": resource_name,
        "count":    len(results),
        "tasks":    results,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Modifying tasks
# ---------------------------------------------------------------------------

@mcp.tool()
def update_task(
    unique_id:        int,
    name:             str = "",
    percent_complete: int = -1,
    notes:            str = "",
    start:            str = "",
    finish:           str = "",
    duration_days:    int = -1,
    manual:           bool = None,
    rag:              str = "",
    text2:            str = "",
    text3:            str = "",
    flag1:            bool = None,
    flag2:            bool = None,
    priority:         int = -1,
    task_type:        str = "",
) -> str:
    """
    Update one or more properties of a task identified by UniqueID.
    Only the fields you provide are changed.

    Args:
        unique_id:        Task UniqueID (required).
        name:             New task name.
        percent_complete: 0-100.
        notes:            Free-text notes.
        start:            Start date as YYYY-MM-DD.
        finish:           Finish date as YYYY-MM-DD.
        duration_days:    Duration in working days (0+ to set).
        manual:           True for manually scheduled, False for auto-scheduled.
        rag:              RAG status: 'Red', 'Amber', or 'Green' (stored in Text1).
        text2:            Custom Text2 field.
        text3:            Custom Text3 field.
        flag1:            Custom Flag1 boolean.
        flag2:            Custom Flag2 boolean.
        priority:         Leveling priority 0-1000 (default 500). 0+ to set.
        task_type:        'FixedUnits', 'FixedDuration', or 'FixedWork'.
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    for t in proj.Tasks:
        if t is None or t.UniqueID != unique_id:
            continue

        changed = []
        if name:
            t.Name = name;              changed.append("name")
        if percent_complete >= 0:
            t.PercentComplete = percent_complete; changed.append("percent_complete")
        if notes:
            t.Notes = notes;            changed.append("notes")
        if start:
            t.Start = _parse_date(start); changed.append("start")
        if finish:
            t.Finish = _parse_date(finish); changed.append("finish")
        if duration_days >= 0:
            t.Duration = duration_days * mpd; changed.append("duration_days")
        if manual is not None:
            t.Manual = manual;          changed.append("manual")
        if rag:
            t.Text1 = rag;              changed.append("rag/text1")
        if text2:
            t.Text2 = text2;            changed.append("text2")
        if text3:
            t.Text3 = text3;            changed.append("text3")
        if flag1 is not None:
            t.Flag1 = flag1;            changed.append("flag1")
        if flag2 is not None:
            t.Flag2 = flag2;            changed.append("flag2")
        if priority >= 0:
            t.Priority = priority;      changed.append("priority")
        if task_type:
            TYPE_MAP = {"fixedunits": 0, "fixedduration": 1, "fixedwork": 2}
            tt = TYPE_MAP.get(task_type.lower())
            if tt is not None:
                t.Type = tt;            changed.append("type")

        app.FileSave()
        return json.dumps({
            "status":   "updated",
            "unique_id": unique_id,
            "name":     t.Name,
            "changed":  changed,
        }, indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


@mcp.tool()
def bulk_update_rag(updates: str) -> str:
    """
    Update RAG status for multiple tasks at once.

    Args:
        updates: JSON string — list of {unique_id, rag} objects.
                 Example: '[{"unique_id": 42, "rag": "Red"}, {"unique_id": 55, "rag": "Green"}]'
    """
    items = json.loads(updates)
    app   = get_app()
    proj  = get_proj(app)

    uid_map = {t.UniqueID: t for t in proj.Tasks if t is not None}

    results = []
    for item in items:
        uid = item["unique_id"]
        rag = item["rag"]
        if uid in uid_map:
            uid_map[uid].Text1 = rag
            results.append({"unique_id": uid, "status": "updated", "rag": rag})
        else:
            results.append({"unique_id": uid, "status": "not_found"})

    app.FileSave()
    return json.dumps({"updated": len([r for r in results if r["status"] == "updated"]),
                       "results": results}, indent=2)


@mcp.tool()
def bulk_update_tasks(updates_json: str) -> str:
    """
    Update multiple tasks in one call. Suspends auto-calc for performance.

    Args:
        updates_json: JSON string — list of objects with fields:
            unique_id (required), name, start, finish, duration_days,
            percent_complete, rag, text2, text3, notes, manual (bool).
            Example: '[{"unique_id": 42, "rag": "Red", "percent_complete": 50}]'
    """
    items = json.loads(updates_json)
    app   = get_app()
    proj  = get_proj(app)
    mpd   = _get_mpd(proj)

    app.Calculation = 0
    try:
        uid_map = {t.UniqueID: t for t in proj.Tasks if t is not None}
        updated = 0
        not_found = []

        for item in items:
            uid = item["unique_id"]
            t = uid_map.get(uid)
            if t is None:
                not_found.append(uid)
                continue

            if "name" in item and item["name"]:
                t.Name = item["name"]
            if "start" in item and item["start"]:
                t.Start = _parse_date(item["start"])
            if "finish" in item and item["finish"]:
                t.Finish = _parse_date(item["finish"])
            if "duration_days" in item and item["duration_days"] is not None and item["duration_days"] >= 0:
                t.Duration = item["duration_days"] * mpd
            if "percent_complete" in item and item["percent_complete"] is not None and item["percent_complete"] >= 0:
                t.PercentComplete = item["percent_complete"]
            if "rag" in item and item["rag"]:
                t.Text1 = item["rag"]
            if "text2" in item and item["text2"]:
                t.Text2 = item["text2"]
            if "text3" in item and item["text3"]:
                t.Text3 = item["text3"]
            if "notes" in item and item["notes"]:
                t.Notes = item["notes"]
            if "manual" in item and item["manual"] is not None:
                t.Manual = item["manual"]

            updated += 1
    finally:
        app.CalculateProject()
        app.Calculation = -1

    app.FileSave()
    return json.dumps({
        "updated":   updated,
        "not_found": not_found,
    }, indent=2)


@mcp.tool()
def add_task(
    name:          str,
    outline_level: int  = 1,
    start:         str  = "",
    finish:        str  = "",
    duration_days: int  = 1,
    milestone:     bool = False,
    notes:         str  = "",
    resource:      str  = "",
    rag:           str  = "",
    after_unique_id: int = 0,
) -> str:
    """
    Add a new task to the active project.

    Args:
        name:            Task name (required).
        outline_level:   WBS level (1 = top-level, 2 = sub-task, etc.).
        start:           Start date YYYY-MM-DD (optional).
        finish:          Finish date YYYY-MM-DD (optional).
        duration_days:   Duration in days (default 1).
        milestone:       True to create as a milestone.
        notes:           Free-text notes.
        resource:        Resource name to assign.
        rag:             RAG status stored in Text1.
        after_unique_id: Insert after this task's UniqueID (0 = append at end).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    app.Calculation = 0

    try:
        task = proj.Tasks.Add(name)
        task.OutlineLevel  = outline_level
        task.Milestone     = milestone
        if milestone:
            task.Duration = 0
        else:
            task.Duration = max(duration_days, 1) * mpd

        if start:
            task.Start = _parse_date(start)
        if finish:
            task.Finish = _parse_date(finish)
        if notes:
            task.Notes = notes
        if resource:
            task.ResourceNames = resource
        if rag:
            task.Text1 = rag

    finally:
        app.CalculateProject()
        app.Calculation = -1

    app.FileSave()
    return json.dumps({
        "status":    "created",
        "unique_id": task.UniqueID,
        "id":        task.ID,
        "name":      task.Name,
    }, indent=2)


@mcp.tool()
def bulk_add_tasks(tasks_json: str) -> str:
    """
    Add multiple tasks in one call. Critical for roadmap generation.
    Suspends auto-calc for performance. Tasks are added sequentially;
    outline_level controls WBS hierarchy.

    Args:
        tasks_json: JSON string — list of task objects with fields:
            name (required), outline_level (default 1), start, finish,
            duration_days (default 1), milestone (bool), resource,
            rag, text2, text3, notes, manual (bool).
            Example: '[{"name": "Phase 1", "outline_level": 1},
                       {"name": "Task A", "outline_level": 2, "start": "2026-04-01"}]'
    """
    tasks = json.loads(tasks_json)
    app   = get_app()
    proj  = get_proj(app)
    mpd   = _get_mpd(proj)

    app.Calculation = 0
    created = []

    try:
        for item in tasks:
            t = proj.Tasks.Add(item["name"])
            t.OutlineLevel = item.get("outline_level", 1)
            t.Milestone = item.get("milestone", False)

            dur = item.get("duration_days", 1)
            if t.Milestone:
                t.Duration = 0
            else:
                t.Duration = max(dur, 1) * mpd

            if item.get("start"):
                t.Start = _parse_date(item["start"])
            if item.get("finish"):
                t.Finish = _parse_date(item["finish"])
            if item.get("resource"):
                t.ResourceNames = item["resource"]
            if item.get("rag"):
                t.Text1 = item["rag"]
            if item.get("text2"):
                t.Text2 = item["text2"]
            if item.get("text3"):
                t.Text3 = item["text3"]
            if item.get("notes"):
                t.Notes = item["notes"]
            if item.get("manual") is not None:
                t.Manual = item["manual"]

            created.append({
                "unique_id": t.UniqueID,
                "id":        t.ID,
                "name":      t.Name,
            })
    finally:
        app.CalculateProject()
        app.Calculation = -1

    app.FileSave()
    return json.dumps({
        "created": len(created),
        "tasks":   created,
    }, indent=2)


@mcp.tool()
def delete_task(unique_id: int) -> str:
    """Delete a task by its UniqueID. This cannot be undone after save."""
    app  = get_app()
    proj = get_proj(app)

    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            task_name = t.Name
            task_id   = t.ID
            app.SelectRow(task_id, False)
            app.EditDelete()
            app.FileSave()
            return json.dumps({
                "status":    "deleted",
                "unique_id": unique_id,
                "name":      task_name,
            }, indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


# ---------------------------------------------------------------------------
# TOOLS — Task scheduling control
# ---------------------------------------------------------------------------

@mcp.tool()
def set_task_mode(unique_id: int, manual: bool = True) -> str:
    """
    Set a task to manually or automatically scheduled.

    Args:
        unique_id: Task UniqueID (required).
        manual:    True for manually scheduled (default), False for auto-scheduled.
    """
    app  = get_app()
    proj = get_proj(app)

    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            t.Manual = manual
            app.FileSave()
            return json.dumps({
                "status":    "updated",
                "unique_id": unique_id,
                "name":      t.Name,
                "manual":    manual,
            }, indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


@mcp.tool()
def bulk_set_task_mode(updates_json: str) -> str:
    """
    Set manual/auto schedule mode for multiple tasks, or by scope.

    Args:
        updates_json: JSON string — EITHER:
            A list of {unique_id, manual} objects:
              '[{"unique_id": 42, "manual": true}]'
            OR a scope object:
              '{"mode": "manual", "scope": "all"}'
              '{"mode": "auto", "scope": "summary"}'
              '{"mode": "manual", "scope": "non_summary"}'
    """
    data = json.loads(updates_json)
    app  = get_app()
    proj = get_proj(app)

    updated = 0

    if isinstance(data, dict) and "scope" in data:
        manual = data.get("mode", "manual").lower() == "manual"
        scope  = data.get("scope", "all").lower()

        for t in proj.Tasks:
            if t is None:
                continue
            if scope == "all":
                t.Manual = manual; updated += 1
            elif scope == "summary" and t.Summary:
                t.Manual = manual; updated += 1
            elif scope == "non_summary" and not t.Summary:
                t.Manual = manual; updated += 1

    else:
        items = data if isinstance(data, list) else [data]
        uid_map = {t.UniqueID: t for t in proj.Tasks if t is not None}
        for item in items:
            uid = item["unique_id"]
            t = uid_map.get(uid)
            if t is not None:
                t.Manual = item.get("manual", True)
                updated += 1

    app.FileSave()
    return json.dumps({"updated": updated}, indent=2)


@mcp.tool()
def set_constraint(unique_id: int, constraint_type: str = "SNET", constraint_date: str = "") -> str:
    """
    Set a scheduling constraint on a task.

    Args:
        unique_id:       Task UniqueID (required).
        constraint_type: One of: ASAP, ALAP, MSO, MFO, SNET, SNLT, FNET, FNLT (default SNET).
        constraint_date: Date as YYYY-MM-DD (required for all types except ASAP/ALAP).
    """
    CONSTRAINT_MAP = {
        "ASAP": 0, "ALAP": 1, "MSO": 2, "MFO": 3,
        "SNET": 4, "SNLT": 5, "FNET": 6, "FNLT": 7,
    }

    ct = constraint_type.upper()
    if ct not in CONSTRAINT_MAP:
        return json.dumps({"error": f"Unknown constraint type '{constraint_type}'. Use: {list(CONSTRAINT_MAP.keys())}"})

    app  = get_app()
    proj = get_proj(app)

    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            t.ConstraintType = CONSTRAINT_MAP[ct]
            if constraint_date and ct not in ("ASAP", "ALAP"):
                t.ConstraintDate = _parse_date(constraint_date)
            app.FileSave()
            return json.dumps({
                "status":          "updated",
                "unique_id":       unique_id,
                "name":            t.Name,
                "constraint_type": ct,
                "constraint_date": constraint_date or "N/A",
            }, indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


@mcp.tool()
def clear_estimated_flags() -> str:
    """
    Remove the estimated '?' flag from all task dates in the active project.
    This cleans up the question marks that appear on dates in MS Project.
    """
    app  = get_app()
    proj = get_proj(app)

    count = 0
    for t in proj.Tasks:
        if t is not None:
            try:
                t.Estimated = False
                count += 1
            except Exception:
                pass

    app.FileSave()
    return json.dumps({"status": "cleared", "tasks_updated": count}, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Utilities
# ---------------------------------------------------------------------------

@mcp.tool()
def search_tasks(query: str, include_summary: bool = False) -> str:
    """
    Search for tasks by name (case-insensitive substring match).
    Returns matching tasks with their UniqueIDs for use in other tools.
    """
    return get_tasks(include_summary=include_summary, keyword=query)


@mcp.tool()
def get_progress_summary() -> str:
    """
    Return a high-level progress summary:
    - Count by % complete bucket (0%, 1-99%, 100%)
    - Count by RAG status (Text1)
    - Count of overdue, critical tasks
    """
    import datetime
    today = datetime.datetime.now()
    app   = get_app()
    proj  = get_proj(app)

    not_started = in_progress = complete = 0
    rag_counts  = {"Red": 0, "Amber": 0, "Green": 0, "Other": 0}
    overdue     = critical = 0

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue

        pct = t.PercentComplete
        if pct == 0:
            not_started += 1
        elif pct < 100:
            in_progress += 1
        else:
            complete += 1

        rag = (t.Text1 or "").strip()
        if rag in rag_counts:
            rag_counts[rag] += 1
        elif rag:
            rag_counts["Other"] += 1

        if t.Critical:
            critical += 1

        try:
            fin = _to_naive(t.Finish)
            if fin and fin < today and pct < 100:
                overdue += 1
        except Exception:
            pass

    total = not_started + in_progress + complete

    return json.dumps({
        "project":     proj.Name,
        "total_tasks": total,
        "by_progress": {
            "not_started": not_started,
            "in_progress": in_progress,
            "complete":    complete,
            "pct_complete": round(complete / total * 100, 1) if total else 0,
        },
        "by_rag": rag_counts,
        "overdue":   overdue,
        "critical":  critical,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — WBS & Hierarchy
# ---------------------------------------------------------------------------

@mcp.tool()
def indent_task(unique_id: int, direction: str = "indent") -> str:
    """
    Promote or demote a task in the WBS hierarchy.

    Args:
        unique_id: Task UniqueID (required).
        direction: 'indent' to demote (increase level) or 'outdent' to promote (decrease level).
    """
    app  = get_app()
    proj = get_proj(app)

    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            old_level = t.OutlineLevel
            app.SelectRow(t.ID, False)
            if direction.lower() == "outdent":
                app.OutlineOutdent()
            else:
                app.OutlineIndent()
            app.FileSave()
            return json.dumps({
                "status":    "updated",
                "unique_id": unique_id,
                "name":      t.Name,
                "old_level": old_level,
                "new_level": t.OutlineLevel,
            }, indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


@mcp.tool()
def get_wbs_structure(max_level: int = 0) -> str:
    """
    Export the full WBS hierarchy as a nested JSON tree.
    Useful for dashboards, reporting, and verifying project structure.

    Args:
        max_level: Maximum outline level to include (0 = all levels).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    def fmt(dt):
        try:
            return str(dt)[:10] if dt else None
        except Exception:
            return None

    # Build flat list first
    flat = []
    for t in proj.Tasks:
        if t is None:
            continue
        if max_level > 0 and t.OutlineLevel > max_level:
            continue
        flat.append({
            "unique_id":     t.UniqueID,
            "id":            t.ID,
            "name":          t.Name,
            "level":         t.OutlineLevel,
            "summary":       bool(t.Summary),
            "milestone":     bool(t.Milestone),
            "start":         fmt(t.Start),
            "finish":        fmt(t.Finish),
            "duration_days": round(t.Duration / mpd, 2) if t.Duration else 0,
            "children":      [],
        })

    # Build tree using stack
    root = {"name": proj.Name, "level": 0, "children": []}
    stack = [root]

    for node in flat:
        level = node["level"]
        # Pop stack back to parent level
        while len(stack) > level:
            stack.pop()
        # Append to current parent
        stack[-1]["children"].append(node)
        # Push this node as potential parent
        stack.append(node)

    return json.dumps(root, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Schedule Analysis
# ---------------------------------------------------------------------------

@mcp.tool()
def get_schedule_analysis() -> str:
    """
    Return float/slack metrics and schedule health for all non-summary tasks.
    TotalSlack and FreeSlack are converted from minutes to working days.
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    def fmt(dt):
        try:
            return str(dt)[:10] if dt else None
        except Exception:
            return None

    tasks = []
    zero_float = 0
    negative_float = 0
    total_slack_sum = 0
    count = 0

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        count += 1

        try:
            ts = round(t.TotalSlack / mpd, 2) if t.TotalSlack is not None else 0
        except Exception:
            ts = 0
        try:
            fs = round(t.FreeSlack / mpd, 2) if t.FreeSlack is not None else 0
        except Exception:
            fs = 0

        if ts == 0:
            zero_float += 1
        if ts < 0:
            negative_float += 1
        total_slack_sum += ts

        tasks.append({
            "unique_id":       t.UniqueID,
            "name":            t.Name,
            "total_slack_days": ts,
            "free_slack_days":  fs,
            "critical":        bool(t.Critical),
            "start":           fmt(t.Start),
            "finish":          fmt(t.Finish),
        })

    return json.dumps({
        "summary": {
            "total_tasks":     count,
            "zero_float":      zero_float,
            "negative_float":  negative_float,
            "avg_total_slack":  round(total_slack_sum / count, 2) if count else 0,
        },
        "tasks": tasks,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 3: Advanced Filtering & Queries
# ---------------------------------------------------------------------------

@mcp.tool()
def filter_tasks(filters_json: str) -> str:
    """
    Powerful AND-logic filtering across all task fields with sort and pagination.

    Args:
        filters_json: JSON string with filter keys (all optional):
            rag, resource, start_after, start_before, finish_after, finish_before,
            min_pct, max_pct, outline_level, critical (bool), milestone (bool),
            active (bool), summary (bool), name_contains,
            text1, text2, text3, flag1 (bool), flag2 (bool),
            sort_by (field name), sort_desc (bool), limit (int), offset (int).
            Example: '{"rag": "Red", "critical": true, "limit": 20}'
    """
    f = json.loads(filters_json)
    app  = get_app()
    proj = get_proj(app)

    predicates = []

    if "rag" in f:
        v = f["rag"].lower()
        predicates.append(lambda t, _v=v: (t.Text1 or "").strip().lower() == _v)
    if "resource" in f:
        v = f["resource"].lower()
        predicates.append(lambda t, _v=v: _v in (t.ResourceNames or "").lower())
    if "start_after" in f:
        d = _parse_date(f["start_after"])
        predicates.append(lambda t, _d=d: t.Start is not None and t.Start >= _d)
    if "start_before" in f:
        d = _parse_date(f["start_before"])
        predicates.append(lambda t, _d=d: t.Start is not None and t.Start <= _d)
    if "finish_after" in f:
        d = _parse_date(f["finish_after"])
        predicates.append(lambda t, _d=d: t.Finish is not None and t.Finish >= _d)
    if "finish_before" in f:
        d = _parse_date(f["finish_before"])
        predicates.append(lambda t, _d=d: t.Finish is not None and t.Finish <= _d)
    if "min_pct" in f:
        v = f["min_pct"]
        predicates.append(lambda t, _v=v: t.PercentComplete >= _v)
    if "max_pct" in f:
        v = f["max_pct"]
        predicates.append(lambda t, _v=v: t.PercentComplete <= _v)
    if "outline_level" in f:
        v = f["outline_level"]
        predicates.append(lambda t, _v=v: t.OutlineLevel == _v)
    if "critical" in f:
        v = f["critical"]
        predicates.append(lambda t, _v=v: bool(t.Critical) == _v)
    if "milestone" in f:
        v = f["milestone"]
        predicates.append(lambda t, _v=v: bool(t.Milestone) == _v)
    if "active" in f:
        v = f["active"]
        predicates.append(lambda t, _v=v: bool(t.Active) == _v)
    if "summary" in f:
        v = f["summary"]
        predicates.append(lambda t, _v=v: bool(t.Summary) == _v)
    if "name_contains" in f:
        v = f["name_contains"].lower()
        predicates.append(lambda t, _v=v: _v in t.Name.lower())
    if "text1" in f:
        v = f["text1"].lower()
        predicates.append(lambda t, _v=v: (t.Text1 or "").strip().lower() == _v)
    if "text2" in f:
        v = f["text2"].lower()
        predicates.append(lambda t, _v=v: (t.Text2 or "").strip().lower() == _v)
    if "text3" in f:
        v = f["text3"].lower()
        predicates.append(lambda t, _v=v: (t.Text3 or "").strip().lower() == _v)
    if "flag1" in f:
        v = f["flag1"]
        predicates.append(lambda t, _v=v: bool(t.Flag1) == _v)
    if "flag2" in f:
        v = f["flag2"]
        predicates.append(lambda t, _v=v: bool(t.Flag2) == _v)

    # Collect matching tasks
    matched = []
    for t in proj.Tasks:
        if t is None:
            continue
        if all(p(t) for p in predicates):
            matched.append(task_to_dict(t, proj))

    # Sort
    sort_by = f.get("sort_by", "")
    if sort_by and matched:
        desc = f.get("sort_desc", False)
        try:
            matched.sort(key=lambda x: x.get(sort_by) or "", reverse=desc)
        except Exception:
            pass

    total = len(matched)
    offset = f.get("offset", 0)
    limit  = f.get("limit", 0)
    if offset > 0:
        matched = matched[offset:]
    if limit > 0:
        matched = matched[:limit]

    return json.dumps({
        "total_matching": total,
        "returned":       len(matched),
        "offset":         offset,
        "limit":          limit or total,
        "tasks":          matched,
    }, indent=2)


@mcp.tool()
def group_tasks_by(field: str, include_tasks: bool = False) -> str:
    """
    Group non-summary tasks by a field and return counts per group.

    Args:
        field:         Field to group by: 'rag', 'resource', 'outline_level', 'critical',
                       'milestone', 'percent_complete', 'text1', 'text2', 'text3',
                       'flag1', 'flag2'.
        include_tasks: If true, include task list per group (default false).
    """
    app  = get_app()
    proj = get_proj(app)

    groups = {}
    total  = 0

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        total += 1

        td = task_to_dict(t, proj) if include_tasks else None

        if field == "resource":
            # Split comma-separated resource names
            names = [n.strip() for n in (t.ResourceNames or "").split(",") if n.strip()]
            if not names:
                names = ["(unassigned)"]
            keys = names
        elif field == "percent_complete":
            pct = t.PercentComplete
            if pct == 0:
                keys = ["0%"]
            elif pct <= 25:
                keys = ["1-25%"]
            elif pct <= 50:
                keys = ["26-50%"]
            elif pct <= 75:
                keys = ["51-75%"]
            elif pct < 100:
                keys = ["76-99%"]
            else:
                keys = ["100%"]
        elif field in ("rag", "text1"):
            keys = [(t.Text1 or "").strip() or "(blank)"]
        elif field == "text2":
            keys = [(t.Text2 or "").strip() or "(blank)"]
        elif field == "text3":
            keys = [(t.Text3 or "").strip() or "(blank)"]
        elif field == "outline_level":
            keys = [str(t.OutlineLevel)]
        elif field == "critical":
            keys = [str(bool(t.Critical))]
        elif field == "milestone":
            keys = [str(bool(t.Milestone))]
        elif field == "flag1":
            keys = [str(bool(t.Flag1))]
        elif field == "flag2":
            keys = [str(bool(t.Flag2))]
        else:
            keys = [str(getattr(t, field, "(unknown)"))]

        for k in keys:
            if k not in groups:
                groups[k] = {"value": k, "count": 0}
                if include_tasks:
                    groups[k]["tasks"] = []
            groups[k]["count"] += 1
            if include_tasks and td:
                groups[k]["tasks"].append(td)

    result = sorted(groups.values(), key=lambda g: g["count"], reverse=True)

    return json.dumps({
        "field":       field,
        "groups":      result,
        "total_tasks": total,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 3: Schedule & Dependency Intelligence
# ---------------------------------------------------------------------------

@mcp.tool()
def validate_schedule() -> str:
    """
    Comprehensive schedule health check — the PMO's best friend.
    Checks for orphan tasks, missing resources, past-due zero-progress,
    empty summaries, missing dates, and negative slack.
    Returns a health score (0-100) and categorized issues.
    """
    app   = get_app()
    proj  = get_proj(app)
    mpd   = _get_mpd(proj)
    today = datetime.datetime.now()

    issues = {
        "orphan_tasks":       {"count": 0, "tasks": []},
        "no_resources":       {"count": 0, "tasks": []},
        "past_due_no_progress": {"count": 0, "tasks": []},
        "empty_summaries":    {"count": 0, "tasks": []},
        "missing_dates":      {"count": 0, "tasks": []},
        "negative_slack":     {"count": 0, "tasks": []},
    }

    total_tasks = 0
    tasks_list  = []  # (task, is_summary, outline_level)

    for t in proj.Tasks:
        if t is None:
            continue
        total_tasks += 1
        tasks_list.append(t)

    for i, t in enumerate(tasks_list):
        tid = {"unique_id": t.UniqueID, "name": t.Name}

        # Empty summaries: summary with no children
        if t.Summary:
            has_child = False
            if i + 1 < len(tasks_list):
                next_t = tasks_list[i + 1]
                if next_t.OutlineLevel > t.OutlineLevel:
                    has_child = True
            if not has_child:
                issues["empty_summaries"]["count"] += 1
                issues["empty_summaries"]["tasks"].append(tid)
            continue  # Skip non-leaf checks for summaries

        # Orphan tasks: no predecessors AND no successors
        preds = (t.Predecessors or "").strip()
        # Check if this task is a predecessor for any other task
        has_successor = False
        task_id_str = str(t.ID)
        for other in tasks_list:
            if other is None or other.UniqueID == t.UniqueID:
                continue
            other_preds = (other.Predecessors or "").strip()
            if other_preds:
                # Check if our task ID appears in other's predecessors
                for part in other_preds.split(","):
                    part = part.strip()
                    # Extract the numeric ID from predecessor string like "5FS" or "5"
                    num = ""
                    for ch in part:
                        if ch.isdigit():
                            num += ch
                        else:
                            break
                    if num == task_id_str:
                        has_successor = True
                        break
            if has_successor:
                break

        if not preds and not has_successor:
            issues["orphan_tasks"]["count"] += 1
            issues["orphan_tasks"]["tasks"].append(tid)

        # No resources (non-milestone)
        if not t.Milestone and not (t.ResourceNames or "").strip():
            issues["no_resources"]["count"] += 1
            issues["no_resources"]["tasks"].append(tid)

        # Past due, zero progress
        try:
            fin = _to_naive(t.Finish)
            if fin and fin < today and t.PercentComplete == 0:
                issues["past_due_no_progress"]["count"] += 1
                issues["past_due_no_progress"]["tasks"].append(tid)
        except Exception:
            pass

        # Missing dates
        try:
            if not t.Start or not t.Finish:
                issues["missing_dates"]["count"] += 1
                issues["missing_dates"]["tasks"].append(tid)
        except Exception:
            pass

        # Negative slack
        try:
            if t.TotalSlack is not None and t.TotalSlack < 0:
                issues["negative_slack"]["count"] += 1
                issues["negative_slack"]["tasks"].append(tid)
        except Exception:
            pass

    total_issues = sum(cat["count"] for cat in issues.values())
    health_score = max(0, round(100 - (total_issues / total_tasks * 100))) if total_tasks else 0

    return json.dumps({
        "project":      proj.Name,
        "health_score": health_score,
        "issues":       issues,
        "summary": {
            "total_tasks":  total_tasks,
            "total_issues": total_issues,
        },
    }, indent=2)


@mcp.tool()
def get_milestone_report(days_ahead: int = 30, upcoming_count: int = 10) -> str:
    """
    Milestone-focused status report for executive dashboards.
    Categorizes milestones as complete, overdue, at_risk, or on_track.
    Includes baseline variance if a baseline is saved.

    Args:
        days_ahead:     Number of days ahead to consider 'at risk' (default 30).
        upcoming_count: Max upcoming milestones to return (default 10).
    """
    app   = get_app()
    proj  = get_proj(app)
    today = datetime.datetime.now()
    horizon = today + datetime.timedelta(days=days_ahead)

    by_status = {"complete": 0, "overdue": 0, "at_risk": 0, "on_track": 0}
    upcoming  = []
    overdue   = []
    total     = 0

    for t in proj.Tasks:
        if t is None or not t.Milestone:
            continue
        total += 1

        finish = None
        try:
            finish = _to_naive(t.Finish)
        except Exception:
            pass

        pct = t.PercentComplete

        # Baseline variance
        variance_days = None
        try:
            bf = _to_naive(t.BaselineFinish)
            if bf and finish:
                delta = finish - bf
                variance_days = delta.days if hasattr(delta, "days") else None
        except Exception:
            pass

        entry = {
            "unique_id":      t.UniqueID,
            "name":           t.Name,
            "finish":         _fmt_date(finish),
            "percent":        pct,
            "variance_days":  variance_days,
        }

        if pct >= 100:
            by_status["complete"] += 1
        elif finish and finish < today:
            by_status["overdue"] += 1
            overdue.append(entry)
        elif finish and finish <= horizon and pct < 100:
            by_status["at_risk"] += 1
            upcoming.append(entry)
        else:
            by_status["on_track"] += 1
            if finish:
                upcoming.append(entry)

    # Sort upcoming by finish asc, overdue by finish asc
    upcoming.sort(key=lambda x: x["finish"] or "")
    overdue.sort(key=lambda x: x["finish"] or "")

    return json.dumps({
        "total_milestones": total,
        "by_status":        by_status,
        "upcoming":         upcoming[:upcoming_count],
        "overdue":          overdue,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 3: Resource Intelligence
# ---------------------------------------------------------------------------


@mcp.tool()
def level_resources() -> str:
    """
    Run MS Project's built-in resource leveling algorithm.
    WARNING: This may shift task dates. Save a baseline first if tracking variance.
    """
    app  = get_app()
    proj = get_proj(app)

    app.LevelNow()
    app.FileSave()

    return json.dumps({
        "status":  "leveled",
        "project": proj.Name,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 3: Deadline & Active Task Management
# ---------------------------------------------------------------------------

@mcp.tool()
def set_deadline(unique_id: int, deadline_date: str) -> str:
    """
    Set a soft deadline on a task. Shows a visual indicator if finish > deadline.
    Unlike hard constraints, deadlines don't affect scheduling.

    Args:
        unique_id:     Task UniqueID (required).
        deadline_date: Deadline as YYYY-MM-DD, or 'clear' to remove.
    """
    app  = get_app()
    proj = get_proj(app)

    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    if deadline_date.lower() == "clear":
        t.Deadline = "NA"
        app.FileSave()
        return json.dumps({
            "status":    "cleared",
            "unique_id": unique_id,
            "name":      t.Name,
            "deadline":  None,
        }, indent=2)

    dl = _parse_date(deadline_date)
    t.Deadline = dl
    app.FileSave()

    deadline_missed = False
    try:
        if t.Finish and t.Finish > dl:
            deadline_missed = True
    except Exception:
        pass

    return json.dumps({
        "status":          "set",
        "unique_id":       unique_id,
        "name":            t.Name,
        "deadline":        deadline_date,
        "finish":          _fmt_date(t.Finish),
        "deadline_missed": deadline_missed,
    }, indent=2)


@mcp.tool()
def set_task_active(unique_id: int, active: bool = True) -> str:
    """
    Activate or deactivate a task. Deactivated tasks are excluded from scheduling
    but remain visible (soft-delete).

    Args:
        unique_id: Task UniqueID (required).
        active:    True to activate (default), False to deactivate.
    """
    app  = get_app()
    proj = get_proj(app)

    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    t.Active = active
    app.FileSave()

    return json.dumps({
        "status":    "updated",
        "unique_id": unique_id,
        "name":      t.Name,
        "active":    active,
    }, indent=2)


@mcp.tool()
def dry_run_bulk_update(updates_json: str) -> str:
    """
    Preview bulk changes without modifying anything — the enterprise safety net.
    Shows what would change for each task without actually applying updates.

    Args:
        updates_json: JSON string — list of objects with fields:
            unique_id (required), name, start, finish, duration_days,
            percent_complete, rag, text2, text3, notes.
            Same format as bulk_update_tasks.
    """
    items = json.loads(updates_json)
    app   = get_app()
    proj  = get_proj(app)
    mpd   = _get_mpd(proj)

    changes   = []
    not_found = []
    no_change = []

    for item in items:
        uid = item.get("unique_id")
        if uid is None:
            continue

        t = _find_task(proj, uid)
        if t is None:
            not_found.append(uid)
            continue

        field_changes = []

        if "name" in item and item["name"] and item["name"] != t.Name:
            field_changes.append({"field": "name", "old": t.Name, "new": item["name"]})
        if "start" in item and item["start"]:
            old_start = _fmt_date(t.Start)
            if old_start != item["start"]:
                field_changes.append({"field": "start", "old": old_start, "new": item["start"]})
        if "finish" in item and item["finish"]:
            old_finish = _fmt_date(t.Finish)
            if old_finish != item["finish"]:
                field_changes.append({"field": "finish", "old": old_finish, "new": item["finish"]})
        if "duration_days" in item and item["duration_days"] is not None:
            old_dur = round(t.Duration / mpd, 2) if t.Duration else 0
            if old_dur != item["duration_days"]:
                field_changes.append({"field": "duration_days", "old": old_dur, "new": item["duration_days"]})
        if "percent_complete" in item and item["percent_complete"] is not None:
            if t.PercentComplete != item["percent_complete"]:
                field_changes.append({"field": "percent_complete", "old": t.PercentComplete, "new": item["percent_complete"]})
        if "rag" in item and item["rag"]:
            old_rag = (t.Text1 or "").strip()
            if old_rag != item["rag"]:
                field_changes.append({"field": "rag", "old": old_rag, "new": item["rag"]})
        if "text2" in item and item["text2"]:
            old = (t.Text2 or "").strip()
            if old != item["text2"]:
                field_changes.append({"field": "text2", "old": old, "new": item["text2"]})
        if "text3" in item and item["text3"]:
            old = (t.Text3 or "").strip()
            if old != item["text3"]:
                field_changes.append({"field": "text3", "old": old, "new": item["text3"]})
        if "notes" in item and item["notes"]:
            old = (t.Notes or "").strip()
            if old != item["notes"]:
                field_changes.append({"field": "notes", "old": old[:50], "new": item["notes"][:50]})

        if field_changes:
            changes.append({
                "unique_id": uid,
                "name":      t.Name,
                "fields":    field_changes,
            })
        else:
            no_change.append(uid)

    total_changes = sum(len(c["fields"]) for c in changes)

    return json.dumps({
        "preview":              True,
        "changes":              changes,
        "not_found":            not_found,
        "no_change":            no_change,
        "total_changes":        total_changes,
        "total_tasks_affected": len(changes),
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 4, Tier 1: High Impact
# ---------------------------------------------------------------------------


@mcp.tool()
def move_task(unique_id: int, after_unique_id: int) -> str:
    """
    Reposition a task to appear after another task.

    Args:
        unique_id:       Task UniqueID to move (required).
        after_unique_id: Place moved task after this task's UniqueID (required).
    """
    app  = get_app()
    proj = get_proj(app)

    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    after_t = _find_task(proj, after_unique_id)
    if after_t is None:
        return json.dumps({"error": f"Target task UniqueID {after_unique_id} not found."})

    old_id = t.ID
    old_level = t.OutlineLevel

    # Select and cut the source task row
    app.SelectRow(t.ID, False)
    app.EditCut()

    # Re-find the after task (IDs may have shifted after cut)
    after_t = _find_task(proj, after_unique_id)
    if after_t is None:
        # Paste back at original position as fallback
        app.EditPaste()
        return json.dumps({"error": "Target task not found after cut. Task restored."})

    # Select row after the target and paste
    target_row = after_t.ID + 1
    if target_row > proj.Tasks.Count:
        target_row = proj.Tasks.Count
    app.SelectRow(target_row, False)
    app.EditPaste()

    # Re-find moved task and verify
    moved = _find_task(proj, unique_id)
    new_id = moved.ID if moved else None

    return json.dumps({
        "status":    "moved",
        "unique_id": unique_id,
        "name":      moved.Name if moved else "(unknown)",
        "old_id":    old_id,
        "new_id":    new_id,
    }, indent=2)


@mcp.tool()
def get_progress_by_wbs(max_level: int = 2) -> str:
    """
    Rolled-up % complete per WBS branch — the PMO dashboard.
    Returns summary tasks at or below max_level with their percent complete.

    Args:
        max_level: Maximum outline level to report (default 2).
    """
    app  = get_app()
    proj = get_proj(app)

    branches = []
    tasks_list = [t for t in proj.Tasks if t is not None]

    for i, t in enumerate(tasks_list):
        if not t.Summary:
            continue
        if t.OutlineLevel > max_level:
            continue

        # Count children
        child_count = 0
        milestones_complete = 0
        milestones_total = 0
        for j in range(i + 1, len(tasks_list)):
            child = tasks_list[j]
            if child.OutlineLevel <= t.OutlineLevel:
                break
            if not child.Summary:
                child_count += 1
                if child.Milestone:
                    milestones_total += 1
                    if child.PercentComplete >= 100:
                        milestones_complete += 1

        branches.append({
            "unique_id":          t.UniqueID,
            "name":               t.Name,
            "level":              t.OutlineLevel,
            "percent_complete":   t.PercentComplete,
            "start":              _fmt_date(t.Start),
            "finish":             _fmt_date(t.Finish),
            "child_count":        child_count,
            "milestones_complete": milestones_complete,
            "milestones_total":   milestones_total,
        })

    return json.dumps({
        "max_level": max_level,
        "branches":  branches,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 4, Tier 2: Nice to Have
# ---------------------------------------------------------------------------

@mcp.tool()
def copy_task_structure(source_unique_id: int, copies: int = 1) -> str:
    """
    Duplicate a task subtree (reusable programme templates).

    Args:
        source_unique_id: UniqueID of the task to copy (and its children if summary).
        copies:           Number of copies to create (default 1).
    """
    app  = get_app()
    proj = get_proj(app)

    source = _find_task(proj, source_unique_id)
    if source is None:
        return json.dumps({"error": f"Task UniqueID {source_unique_id} not found."})

    # Determine range: source + all children (higher outline level)
    source_id = source.ID
    source_level = source.OutlineLevel
    end_id = source_id

    for t in proj.Tasks:
        if t is None:
            continue
        if t.ID > source_id:
            if t.OutlineLevel > source_level:
                end_id = t.ID
            else:
                break

    all_copied = []
    for _copy_num in range(copies):
        # Select the range
        app.SelectRow(source_id, False)
        app.SelectRow(end_id, True)  # extend selection
        app.EditCopy()

        # Paste at end
        last_id = proj.Tasks.Count
        app.SelectRow(last_id, False)
        app.EditPaste()

        # Collect newly created tasks
        new_count = proj.Tasks.Count
        for t in proj.Tasks:
            if t is not None and t.ID > last_id:
                all_copied.append({
                    "unique_id": t.UniqueID,
                    "name":      t.Name,
                })

    return json.dumps({
        "status":       "copied",
        "source_name":  source.Name,
        "copied_tasks": all_copied,
    }, indent=2)


@mcp.tool()
def export_csv(output_path: str, columns_json: str = "", filters_json: str = "") -> str:
    """
    Export filtered task data to CSV for PowerBI / Excel dashboards.

    Args:
        output_path:  Full path for the output CSV file (required).
        columns_json: JSON list of column names to include (optional, default: standard set).
                      Available: any key from task_to_dict output.
        filters_json: JSON object with filter criteria (same format as filter_tasks).
    """
    import csv

    app  = get_app()
    proj = get_proj(app)

    # Default columns
    default_cols = ["unique_id", "name", "outline_level", "start", "finish",
                    "duration_days", "percent_complete", "resource_names", "rag", "critical"]

    columns = default_cols
    if columns_json:
        columns = json.loads(columns_json)

    # Apply filters if provided
    tasks = []
    if filters_json:
        f = json.loads(filters_json)
        # Reuse filter_tasks logic inline
        result = json.loads(filter_tasks(json.dumps(f)))
        tasks = result.get("tasks", [])
    else:
        for t in proj.Tasks:
            if t is not None:
                tasks.append(task_to_dict(t, proj))

    # Write CSV
    with open(output_path, "w", newline="", encoding="utf-8") as fp:
        writer = csv.writer(fp)
        writer.writerow(columns)
        for task in tasks:
            row = [task.get(col, "") for col in columns]
            writer.writerow(row)

    return json.dumps({
        "status":  "exported",
        "path":    output_path,
        "rows":    len(tasks),
        "columns": columns,
    }, indent=2)


@mcp.tool()
def bulk_set_deadlines(deadlines_json: str) -> str:
    """
    Set deadlines on multiple tasks at once.

    Args:
        deadlines_json: JSON string — list of {unique_id, deadline_date}.
            deadline_date can be 'clear' to remove the deadline.
            Example: '[{"unique_id": 42, "deadline_date": "2026-06-01"},
                       {"unique_id": 55, "deadline_date": "clear"}]'
    """
    items = json.loads(deadlines_json)
    app   = get_app()
    proj  = get_proj(app)

    set_count = 0
    cleared = 0
    not_found = []
    errors = []

    for item in items:
        uid = item["unique_id"]
        dd  = item["deadline_date"]

        t = _find_task(proj, uid)
        if t is None:
            not_found.append(uid)
            continue

        try:
            if dd.lower() == "clear":
                t.Deadline = "NA"
                cleared += 1
            else:
                t.Deadline = _parse_date(dd)
                set_count += 1
        except Exception as e:
            errors.append({"unique_id": uid, "error": str(e)})

    return json.dumps({
        "set":       set_count,
        "cleared":   cleared,
        "not_found": not_found,
        "errors":    errors,
    }, indent=2)


@mcp.tool()
def find_available_slack(min_days: int = 5) -> str:
    """
    Find tasks with positive float — where can we absorb delay?

    Args:
        min_days: Minimum total slack in working days (default 5).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    tasks = []
    for t in proj.Tasks:
        if t is None or t.Summary:
            continue

        try:
            ts = t.TotalSlack
            if ts is None:
                continue
            ts_days = round(ts / mpd, 2)
            if ts_days < min_days:
                continue

            fs = 0
            try:
                fs = round(t.FreeSlack / mpd, 2) if t.FreeSlack else 0
            except Exception:
                pass

            tasks.append({
                "unique_id":        t.UniqueID,
                "name":             t.Name,
                "total_slack_days": ts_days,
                "free_slack_days":  fs,
                "start":            _fmt_date(t.Start),
                "finish":           _fmt_date(t.Finish),
                "resource_names":   t.ResourceNames or "",
            })
        except Exception:
            continue

    tasks.sort(key=lambda x: x["total_slack_days"], reverse=True)

    return json.dumps({
        "min_days": min_days,
        "count":    len(tasks),
        "tasks":    tasks,
    }, indent=2)


# ---------------------------------------------------------------------------
# TOOLS — Phase 4, Tier 3: Power User
# ---------------------------------------------------------------------------


@mcp.tool()
def apply_filter(filter_name: str) -> str:
    """
    Apply MS Project's built-in or custom named filter to the Gantt view.
    Use 'All Tasks' to clear any active filter.

    Args:
        filter_name: Name of the filter (e.g. 'Critical', 'Incomplete Tasks', 'All Tasks').
    """
    app = get_app()

    try:
        app.FilterApply(filter_name)
    except Exception as e:
        return json.dumps({"error": f"Failed to apply filter '{filter_name}': {e}"})

    return json.dumps({
        "status": "applied",
        "filter": filter_name,
    }, indent=2)


# ---------------------------------------------------------------------------
# Phase 5 — get_constraints, delete_resource, get_actual_work
# ---------------------------------------------------------------------------

@mcp.tool()
def get_constraints() -> str:
    """Return all tasks with non-default (non-ASAP) scheduling constraints."""
    CONSTRAINT_NAMES = {
        0: "ASAP", 1: "ALAP", 2: "MSO", 3: "MFO",
        4: "SNET", 5: "SNLT", 6: "FNET", 7: "FNLT",
    }
    app  = get_app()
    proj = get_proj(app)

    results = []
    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        try:
            ct = t.ConstraintType
            if ct != 0:  # 0 = ASAP (default)
                results.append({
                    "unique_id":       t.UniqueID,
                    "name":            t.Name,
                    "constraint_type": CONSTRAINT_NAMES.get(ct, f"Unknown({ct})"),
                    "constraint_date": _fmt_date(t.ConstraintDate),
                })
        except Exception:
            continue

    return json.dumps({"count": len(results), "tasks": results}, indent=2)


@mcp.tool()
def get_actual_work() -> str:
    """
    Return actual vs remaining work per task and project totals.
    Work values are converted from COM minutes to hours.
    """
    app  = get_app()
    proj = get_proj(app)

    tasks = []
    total_work = 0.0
    total_actual = 0.0
    total_remaining = 0.0

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        try:
            w = (t.Work or 0) / 60.0
            a = (t.ActualWork or 0) / 60.0
            r = (t.RemainingWork or 0) / 60.0
            total_work += w
            total_actual += a
            total_remaining += r
            tasks.append({
                "unique_id":      t.UniqueID,
                "name":           t.Name,
                "work_hours":     round(w, 2),
                "actual_hours":   round(a, 2),
                "remaining_hours": round(r, 2),
                "pct_work_complete": t.PercentWorkComplete,
            })
        except Exception:
            continue

    pct = round(total_actual / total_work * 100, 1) if total_work else 0.0

    return json.dumps({
        "totals": {
            "work_hours":        round(total_work, 2),
            "actual_hours":      round(total_actual, 2),
            "remaining_hours":   round(total_remaining, 2),
            "pct_work_complete": pct,
        },
        "count": len(tasks),
        "tasks": tasks,
    }, indent=2)


# ---------------------------------------------------------------------------
# Phase 6 — Gap Analysis & Missing Features
# ---------------------------------------------------------------------------


@mcp.tool()
def calculate_project() -> str:
    """
    Recalculate the active project schedule.
    Use after bulk manual changes to ensure dates, slack, and critical path are up to date.
    """
    app = get_app()
    app.CalculateProject()
    return json.dumps({"status": "calculated", "project": app.ActiveProject.Name})


@mcp.tool()
def update_project(complete_through: str, set_0_or_100: bool = False) -> str:
    """
    Mark all tasks complete through a given date (the weekly PMO ritual).
    Tasks that should have finished by the date get their % complete updated.

    Args:
        complete_through: Date as YYYY-MM-DD — tasks scheduled through this date are updated.
        set_0_or_100:     If True, tasks are set to 0% or 100% only (no partial). Default False.
    """
    app  = get_app()
    proj = get_proj(app)
    dt   = _parse_date(complete_through)
    if dt is None:
        return json.dumps({"error": "complete_through date is required (YYYY-MM-DD)."})

    # COM VBA signature: UpdateProject(All, UpdateDate, Action)
    # All = True (entire project), UpdateDate = date, Action:
    #   pjUpdateProjectStatusPctComplete = 0
    #   pjUpdateProject0or100 = 1
    action = 1 if set_0_or_100 else 0
    try:
        app.UpdateProject(True, dt, action)
    except Exception:
        try:
            # Alternative: just date
            app.UpdateProject(True, dt)
        except Exception:
            app.UpdateProject(dt)
    app.FileSave()

    return json.dumps({
        "status": "updated",
        "complete_through": complete_through,
        "set_0_or_100": set_0_or_100,
        "project": proj.Name,
    }, indent=2)


@mcp.tool()
def reschedule_incomplete_work(reschedule_from: str = "") -> str:
    """
    Move remaining work on incomplete tasks to start after the given date
    (or the project status date if not specified).

    Args:
        reschedule_from: Date as YYYY-MM-DD. Empty = use project status date.
    """
    app  = get_app()
    proj = get_proj(app)

    if reschedule_from:
        dt = _parse_date(reschedule_from)
    else:
        dt = proj.StatusDate
        if dt is None or str(dt) == "NA":
            dt = datetime.datetime.now()

    # Set status date, then use UpdateProject to reschedule
    # The reschedule action = updating incomplete tasks from the status date
    try:
        proj.StatusDate = dt
        # UpdateProject with All=True, date, action=0 to push remaining work
        app.UpdateProject(True, dt, 0)
    except Exception:
        try:
            app.UpdateProject(True, dt)
        except Exception:
            proj.StatusDate = dt  # At minimum set the status date
    app.FileSave()

    return json.dumps({
        "status": "rescheduled",
        "reschedule_from": str(dt)[:10],
        "project": proj.Name,
    }, indent=2)


@mcp.tool()
def get_timephased_data(
    unique_id:  int,
    start_date: str,
    end_date:   str,
    timescale:  str = "weekly",
    data_type:  str = "work",
) -> str:
    """
    Get period-by-period timephased data for a task. Essential for S-curves,
    resource loading charts, and cash flow forecasts.

    Args:
        unique_id:  Task UniqueID.
        start_date: Period start as YYYY-MM-DD.
        end_date:   Period end as YYYY-MM-DD.
        timescale:  'daily', 'weekly', or 'monthly' (default 'weekly').
        data_type:  'work', 'cost', 'actual_work', 'actual_cost',
                    'remaining_work', 'baseline_work', 'baseline_cost' (default 'work').
    """
    app  = get_app()
    proj = get_proj(app)

    TIMESCALE_MAP = {"daily": 3, "weekly": 4, "monthly": 5}
    ts = TIMESCALE_MAP.get(timescale.lower())
    if ts is None:
        return json.dumps({"error": f"Unknown timescale '{timescale}'. Use: daily, weekly, monthly."})

    # pjTaskTimescaledWork=1, Cost=2, ActualWork=3, ActualCost=4,
    # RemainingWork=9, BaselineWork=22, BaselineCost=23
    TYPE_MAP = {
        "work": 1, "cost": 2, "actual_work": 3, "actual_cost": 4,
        "remaining_work": 9, "baseline_work": 22, "baseline_cost": 23,
    }
    dt = TYPE_MAP.get(data_type.lower())
    if dt is None:
        return json.dumps({"error": f"Unknown data_type '{data_type}'. Use: {list(TYPE_MAP.keys())}."})

    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    sd = _parse_date(start_date)
    ed = _parse_date(end_date)
    if sd is None or ed is None:
        return json.dumps({"error": "Both start_date and end_date are required (YYYY-MM-DD)."})

    periods = []
    try:
        tsd = t.TimeScaleData(sd, ed, dt, ts)
        for item in tsd:
            val = item.Value
            periods.append({
                "start": _fmt_date(item.StartDate),
                "end":   _fmt_date(item.EndDate),
                "value": float(val) if val else 0.0,
            })
    except Exception as e:
        return json.dumps({"error": f"TimeScaleData failed: {e}"})

    return json.dumps({
        "unique_id": unique_id,
        "name":      t.Name,
        "data_type": data_type,
        "timescale": timescale,
        "periods":   periods,
    }, indent=2)


@mcp.tool()
def set_task_hyperlink(unique_id: int, url: str, text: str = "", sub_address: str = "") -> str:
    """
    Set a hyperlink on a task.

    Args:
        unique_id:   Task UniqueID.
        url:         The hyperlink URL or file path.
        text:        Display text for the hyperlink (optional).
        sub_address: Sub-address / bookmark within the target (optional).
    """
    app  = get_app()
    proj = get_proj(app)
    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    t.HyperlinkAddress    = url
    t.HyperlinkScreenTip  = text or url
    if text:
        t.Hyperlink = text
    if sub_address:
        t.HyperlinkSubAddress = sub_address

    app.FileSave()
    return json.dumps({
        "status":    "updated",
        "unique_id": unique_id,
        "name":      t.Name,
        "hyperlink": url,
        "text":      text or url,
    }, indent=2)


@mcp.tool()
def add_recurring_task(
    name:            str,
    recurrence_type: str = "weekly",
    start_date:      str = "",
    end_date:        str = "",
    duration_days:   float = 1,
    day_of_week:     int = 2,
) -> str:
    """
    Add a recurring task to the project.

    Args:
        name:            Task name.
        recurrence_type: 'daily', 'weekly', or 'monthly' (default 'weekly').
        start_date:      Recurrence range start (YYYY-MM-DD).
        end_date:        Recurrence range end (YYYY-MM-DD).
        duration_days:   Duration of each occurrence in days (default 1).
        day_of_week:     For weekly: 1=Sun, 2=Mon, ..., 7=Sat (default 2=Monday).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    sd = _parse_date(start_date)
    ed = _parse_date(end_date)
    if sd is None or ed is None:
        return json.dumps({"error": "Both start_date and end_date are required (YYYY-MM-DD)."})

    dur = int(duration_days * mpd)

    # pjRecurType: 0=daily, 1=weekly, 2=monthly, 3=yearly
    RECUR_MAP = {"daily": 0, "weekly": 1, "monthly": 2, "yearly": 3}
    rt = RECUR_MAP.get(recurrence_type.lower())
    if rt is None:
        return json.dumps({"error": f"Unknown recurrence_type '{recurrence_type}'. Use: daily, weekly, monthly, yearly."})

    try:
        # MS Project's RecurringTaskInsert is a dialog-only COM method;
        # it cannot be driven programmatically. Instead, we create individual
        # task occurrences under a summary task to simulate recurrence.

        from dateutil.rrule import rrule, DAILY, WEEKLY, MONTHLY, YEARLY
        FREQ_MAP = {"daily": DAILY, "weekly": WEEKLY, "monthly": MONTHLY, "yearly": YEARLY}
        freq = FREQ_MAP.get(recurrence_type.lower(), WEEKLY)

        # Map day_of_week (1=Sun..7=Sat) to dateutil byweekday (0=Mon..6=Sun)
        WEEKDAY_MAP = {1: 6, 2: 0, 3: 1, 4: 2, 5: 3, 6: 4, 7: 5}
        byday = WEEKDAY_MAP.get(day_of_week, 0)

        if freq == WEEKLY:
            dates = list(rrule(freq, dtstart=sd, until=ed, byweekday=byday))
        else:
            dates = list(rrule(freq, dtstart=sd, until=ed))

        if not dates:
            return json.dumps({"error": "No occurrences generated for the given range."})

        # Create summary task
        summary = proj.Tasks.Add(name)
        summary_uid = summary.UniqueID

        # Create each occurrence as a subtask
        occurrence_uids = []
        for i, dt_occ in enumerate(dates, 1):
            occ = proj.Tasks.Add(f"{name} #{i}")
            occ.OutlineIndent()
            occ.Start = dt_occ
            occ.Duration = dur
            occurrence_uids.append(occ.UniqueID)

        app.CalculateProject()
        app.FileSave()

        return json.dumps({
            "status":          "created",
            "name":            name,
            "unique_id":       summary_uid,
            "recurrence_type": recurrence_type,
            "occurrences":     len(dates),
            "start":           start_date,
            "end":             end_date,
        }, indent=2)

    except ImportError:
        return json.dumps({"error": "python-dateutil is required for recurring tasks. Install with: pip install python-dateutil"})
    except Exception as e:
        return json.dumps({"error": f"Failed to create recurring task: {e}"})


# ---------------------------------------------------------------------------
# TOOLS — Phase 7: Critical Path Intelligence
# ---------------------------------------------------------------------------

@mcp.tool()
def get_critical_path_sequence() -> str:
    """
    Return the critical path as an ordered chain from project start to finish.
    Unlike get_critical_path (flat list), this shows the exact sequence of tasks
    that drives the project end date, connected by their dependency links.

    Returns the longest path through the network with total duration,
    each task's contribution, and the driving relationships.
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    # Build adjacency graph of critical tasks only
    critical_tasks = {}
    for t in proj.Tasks:
        if t is not None and t.Critical and not t.Summary:
            critical_tasks[t.UniqueID] = t

    if not critical_tasks:
        return json.dumps({"error": "No critical tasks found. Ensure the project has tasks with dependencies."})

    # Build forward adjacency: uid -> [(successor_uid, link_type, lag_days)]
    forward = {uid: [] for uid in critical_tasks}
    incoming = {uid: 0 for uid in critical_tasks}

    LINK_NAMES = {0: "FF", 1: "FS", 2: "SF", 3: "SS"}

    for uid, t in critical_tasks.items():
        try:
            for dep in t.TaskDependencies:
                if dep.From.UniqueID == uid:
                    succ_uid = dep.To.UniqueID
                    if succ_uid in critical_tasks:
                        lag = round(dep.Lag / mpd, 2) if dep.Lag else 0
                        link = LINK_NAMES.get(dep.Type, "FS")
                        forward[uid].append((succ_uid, link, lag))
                        incoming[succ_uid] = incoming.get(succ_uid, 0) + 1
        except Exception:
            pass

    # Find start nodes (no critical predecessors)
    start_nodes = [uid for uid, count in incoming.items() if count == 0]

    # Find the longest path using DFS (by finish date)
    best_path = []

    def dfs(uid, path):
        nonlocal best_path
        path.append(uid)
        successors = forward.get(uid, [])
        critical_successors = [s for s in successors if s[0] in critical_tasks]

        if not critical_successors:
            # End of chain — check if this is the longest
            if not best_path or len(path) > len(best_path):
                best_path = list(path)
            else:
                # Tie-break by finish date of last task
                last_current = critical_tasks[path[-1]]
                last_best = critical_tasks[best_path[-1]]
                try:
                    if _to_naive(last_current.Finish) > _to_naive(last_best.Finish):
                        best_path = list(path)
                except Exception:
                    if len(path) > len(best_path):
                        best_path = list(path)
        else:
            for succ_uid, _, _ in critical_successors:
                if succ_uid not in path:  # avoid cycles
                    dfs(succ_uid, path)

        path.pop()

    for start in start_nodes:
        dfs(start, [])

    # If no path found from start nodes, fall back to all critical tasks sorted by start
    if not best_path:
        sorted_critical = sorted(critical_tasks.values(), key=lambda t: _to_naive(t.Start) or datetime.datetime.max)
        best_path = [t.UniqueID for t in sorted_critical]

    # Build the ordered sequence with link info
    sequence = []
    total_duration = 0
    for i, uid in enumerate(best_path):
        t = critical_tasks[uid]
        dur = round(t.Duration / mpd, 2) if t.Duration else 0
        total_duration += dur

        entry = {
            "step":            i + 1,
            "unique_id":       t.UniqueID,
            "name":            t.Name,
            "start":           _fmt_date(t.Start),
            "finish":          _fmt_date(t.Finish),
            "duration_days":   dur,
            "milestone":       bool(t.Milestone),
            "percent_complete": t.PercentComplete,
            "resource_names":  t.ResourceNames or "",
        }

        # Add link info to next task
        if i < len(best_path) - 1:
            next_uid = best_path[i + 1]
            for succ_uid, link, lag in forward.get(uid, []):
                if succ_uid == next_uid:
                    entry["link_to_next"] = link
                    entry["lag_days"] = lag
                    break

        sequence.append(entry)

    # Project dates
    proj_start = _fmt_date(proj.ProjectStart)
    proj_finish = _fmt_date(proj.ProjectFinish)

    return json.dumps({
        "project_start":       proj_start,
        "project_finish":      proj_finish,
        "critical_path_length": len(sequence),
        "total_duration_days":  total_duration,
        "total_critical_tasks": len(critical_tasks),
        "sequence":            sequence,
    }, indent=2)


@mcp.tool()
def get_critical_tasks_for_period(
    start_date: str,
    end_date: str,
    include_milestones: bool = True,
    include_non_critical_milestones: bool = False,
) -> str:
    """
    Return critical tasks and key milestones that fall within a date range.
    Perfect for period-focused reporting: 'What's critical in Q2?'

    Args:
        start_date:  Period start (YYYY-MM-DD, required).
        end_date:    Period end (YYYY-MM-DD, required).
        include_milestones: Include critical milestones in results (default True).
        include_non_critical_milestones: Also include non-critical milestones in the period (default False).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    period_start = _parse_date(start_date)
    period_end   = _parse_date(end_date)
    if not period_start or not period_end:
        return json.dumps({"error": "Both start_date and end_date are required (YYYY-MM-DD)."})

    critical_tasks = []
    critical_milestones = []
    other_milestones = []

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue

        try:
            t_start  = _to_naive(t.Start)
            t_finish = _to_naive(t.Finish)
        except Exception:
            continue

        if not t_start or not t_finish:
            continue

        # Check overlap: task intersects the period
        overlaps = t_start <= period_end and t_finish >= period_start

        if not overlaps:
            continue

        dur = round(t.Duration / mpd, 2) if t.Duration else 0

        # Calculate how much of the task falls within the period
        overlap_start = max(t_start, period_start)
        overlap_end   = min(t_finish, period_end)
        overlap_days  = max(0, (overlap_end - overlap_start).days)

        entry = {
            "unique_id":        t.UniqueID,
            "name":             t.Name,
            "start":            _fmt_date(t_start),
            "finish":           _fmt_date(t_finish),
            "duration_days":    dur,
            "percent_complete": t.PercentComplete,
            "milestone":        bool(t.Milestone),
            "critical":         bool(t.Critical),
            "resource_names":   t.ResourceNames or "",
            "overlap_days":     overlap_days,
        }

        # Baseline variance
        try:
            bf = _to_naive(t.BaselineFinish)
            if bf and t_finish:
                entry["finish_variance_days"] = (t_finish - bf).days
        except Exception:
            pass

        if t.Critical:
            if t.Milestone and include_milestones:
                critical_milestones.append(entry)
            elif not t.Milestone:
                critical_tasks.append(entry)
        elif t.Milestone and include_non_critical_milestones:
            other_milestones.append(entry)

    # Sort by start date
    critical_tasks.sort(key=lambda x: x["start"] or "")
    critical_milestones.sort(key=lambda x: x["finish"] or "")
    other_milestones.sort(key=lambda x: x["finish"] or "")

    return json.dumps({
        "period":              {"start": start_date, "end": end_date},
        "critical_task_count": len(critical_tasks),
        "critical_milestone_count": len(critical_milestones),
        "other_milestone_count": len(other_milestones),
        "critical_tasks":      critical_tasks,
        "critical_milestones": critical_milestones,
        "other_milestones":    other_milestones,
    }, indent=2)


@mcp.tool()
def what_if_delay(
    unique_id: int,
    delay_days: int = 5,
) -> str:
    """
    What-if analysis: 'If I delay task X by N days, what happens to the schedule?'
    Simulates the delay WITHOUT modifying the project — read-only analysis.

    Shows:
    - Whether the project end date would change (and by how much)
    - Which tasks would become newly critical
    - Which tasks would lose their slack
    - Downstream tasks affected

    Args:
        unique_id:  Task UniqueID to simulate delaying.
        delay_days: Number of working days to simulate (default 5).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    target = _find_task(proj, unique_id)
    if target is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    if target.Summary:
        return json.dumps({"error": "Cannot simulate delay on a summary task. Choose a work task or milestone."})

    delay_minutes = delay_days * mpd

    # Gather current state of all tasks
    task_data = {}
    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        try:
            ts = t.TotalSlack if t.TotalSlack is not None else 0
            fs = t.FreeSlack if t.FreeSlack is not None else 0
        except Exception:
            ts = 0
            fs = 0

        task_data[t.UniqueID] = {
            "name":         t.Name,
            "start":        _to_naive(t.Start),
            "finish":       _to_naive(t.Finish),
            "total_slack":  ts,
            "free_slack":   fs,
            "critical":     bool(t.Critical),
            "milestone":    bool(t.Milestone),
            "duration":     t.Duration or 0,
        }

    target_info = task_data[unique_id]
    target_total_slack = target_info["total_slack"]
    target_free_slack  = target_info["free_slack"]

    # Project end date impact
    project_finish = _to_naive(proj.ProjectFinish)
    slack_days = round(target_total_slack / mpd, 2)
    excess_delay = delay_days - slack_days  # days beyond available slack

    project_delay_days = max(0, round(excess_delay, 2)) if excess_delay > 0 else 0

    new_project_finish = None
    if project_delay_days > 0 and project_finish:
        new_project_finish = _fmt_date(project_finish + datetime.timedelta(days=int(project_delay_days * 1.4)))  # rough calendar conversion

    # Walk downstream tasks (successors chain)
    downstream_affected = []
    newly_critical = []
    slack_consumed = []
    visited = set()

    def walk_successors(uid, accumulated_delay_min):
        if uid in visited:
            return
        visited.add(uid)

        t_info = task_data.get(uid)
        if not t_info:
            return

        for t in proj.Tasks:
            if t is None or t.UniqueID != uid:
                continue
            try:
                for dep in t.TaskDependencies:
                    if dep.From.UniqueID == uid:
                        succ_uid = dep.To.UniqueID
                        succ_info = task_data.get(succ_uid)
                        if succ_info and succ_uid not in visited:
                            succ_free_slack = succ_info["free_slack"]
                            succ_total_slack = succ_info["total_slack"]
                            effective_delay = max(0, accumulated_delay_min - succ_free_slack)

                            entry = {
                                "unique_id":         succ_uid,
                                "name":              succ_info["name"],
                                "current_start":     _fmt_date(succ_info["start"]),
                                "current_finish":    _fmt_date(succ_info["finish"]),
                                "current_slack_days": round(succ_total_slack / mpd, 2),
                                "remaining_slack_days": round(max(0, succ_total_slack - accumulated_delay_min) / mpd, 2),
                                "would_shift_days":  round(effective_delay / mpd, 2),
                                "milestone":         succ_info["milestone"],
                            }

                            downstream_affected.append(entry)

                            # Check if task becomes newly critical
                            if not succ_info["critical"] and accumulated_delay_min >= succ_total_slack:
                                newly_critical.append({
                                    "unique_id": succ_uid,
                                    "name":      succ_info["name"],
                                    "was_slack_days": round(succ_total_slack / mpd, 2),
                                })

                            # Check if slack is significantly consumed
                            if succ_total_slack > 0 and accumulated_delay_min > 0:
                                pct_consumed = min(100, round(accumulated_delay_min / succ_total_slack * 100))
                                if pct_consumed >= 50:
                                    slack_consumed.append({
                                        "unique_id":      succ_uid,
                                        "name":           succ_info["name"],
                                        "original_slack_days": round(succ_total_slack / mpd, 2),
                                        "percent_consumed":    pct_consumed,
                                    })

                            walk_successors(succ_uid, effective_delay)
            except Exception:
                pass
            break

    walk_successors(unique_id, delay_minutes)

    # Build severity assessment
    if project_delay_days > 0:
        severity = "HIGH"
        summary = f"Project end date would slip by ~{project_delay_days} days. {len(newly_critical)} task(s) become newly critical."
    elif len(newly_critical) > 0:
        severity = "MEDIUM"
        summary = f"No project delay, but {len(newly_critical)} task(s) would become critical (slack fully consumed)."
    elif len(slack_consumed) > 0:
        severity = "LOW"
        summary = f"No project delay, but slack reduced on {len(slack_consumed)} downstream task(s)."
    else:
        severity = "NONE"
        summary = f"Task has {slack_days} days of slack. A {delay_days}-day delay is fully absorbed."

    return json.dumps({
        "task":               {"unique_id": unique_id, "name": target_info["name"]},
        "simulated_delay_days": delay_days,
        "severity":           severity,
        "summary":            summary,
        "current_slack_days": slack_days,
        "project_impact": {
            "current_finish":     _fmt_date(project_finish),
            "estimated_new_finish": new_project_finish,
            "delay_days":         project_delay_days,
        },
        "downstream_affected":   len(downstream_affected),
        "newly_critical_count":  len(newly_critical),
        "newly_critical":        newly_critical,
        "slack_consumed":        slack_consumed,
        "downstream_tasks":      downstream_affected,
    }, indent=2)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # stdout is the MCP stdio transport -- anything written here corrupts
    # the protocol. Diagnostics go to stderr.
    print("Starting MS Project MCP Server...", file=sys.stderr)
    print("MS Project must be running with a file open before using tools.", file=sys.stderr)
    mcp.run()


# ---------------------------------------------------------------------------
# REGISTRATION — add this to claude_desktop_config.json:
#
# {
#   "mcpServers": {
#     "msproject": {
#       "command": "python",
#       "args": ["/path/to/msproject/server.py"]
#     }
#   }
# }
# ---------------------------------------------------------------------------
