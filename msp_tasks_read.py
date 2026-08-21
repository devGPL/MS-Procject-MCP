"""
Reading tasks: queries, filters, roll-ups and exports.

Boundary against msp_tasks_write is mechanical, not taxonomic: does the tool
write to the project? Nothing here calls FileSave or assigns a Task property.
That same property is what makes these tools the ones a read-only fake could
exercise offline.

Two placements are deliberate and would otherwise look wrong:

  apply_filter changes the VIEW, not the data, so it reads as a mutation and
  is not one.

  export_csv sits here rather than beside export_xml in msp_projects, because
  it calls filter_tasks and that edge has to stay inside one module. This
  trades thematic coherence for a contained call graph -- worth knowing before
  hunting for it next to the other export tools.

get_constraints still carries its own local copy of CONSTRAINT_NAMES,
duplicating the one in msp_core. Removing it is a body edit and belongs to the
later consolidation step, not to a cut.
"""

import json
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
    mpd = _get_mpd(proj)

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
        results.append(task_to_dict(t, mpd))

    return json.dumps({"count": len(results), "tasks": results}, indent=2)


@mcp.tool()
def get_task(unique_id: int) -> str:
    """Get full details for a single task by its UniqueID."""
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            return json.dumps(task_to_dict(t, mpd), indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


@mcp.tool()
def get_tasks_by_rag(rag: str = "Red") -> str:
    """
    Return tasks filtered by RAG status stored in the Text1 custom field.
    rag: 'Red', 'Amber', or 'Green'
    """
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    results = []
    for t in proj.Tasks:
        if t is not None and not t.Summary:
            if (t.Text1 or "").strip().lower() == rag.strip().lower():
                results.append(task_to_dict(t, mpd))

    return json.dumps({"rag": rag, "count": len(results), "tasks": results}, indent=2)


@mcp.tool()
def get_overdue_tasks() -> str:
    """Return incomplete tasks whose Finish date is in the past."""
    import datetime
    today = datetime.datetime.now()
    app   = get_app()
    proj  = get_proj(app)
    mpd  = _get_mpd(proj)

    results = []
    for t in proj.Tasks:
        if t is None or t.Summary or t.Milestone:
            continue
        if t.PercentComplete >= 100:
            continue
        try:
            finish = _to_naive(t.Finish)
            if finish and finish < today:
                results.append(task_to_dict(t, mpd))
        except Exception:
            continue

    return json.dumps({"count": len(results), "tasks": results}, indent=2)


@mcp.tool()
def get_tasks_by_resource(resource_name: str) -> str:
    """Return all tasks assigned to a named resource (case-insensitive substring match)."""
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    results = []
    name_lower = resource_name.lower()
    for t in proj.Tasks:
        if t is not None and not t.Summary:
            if name_lower in (t.ResourceNames or "").lower():
                results.append(task_to_dict(t, mpd))

    return json.dumps({
        "resource": resource_name,
        "count":    len(results),
        "tasks":    results,
    }, indent=2)


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
    mpd = _get_mpd(proj)
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
    mpd = _get_mpd(proj)

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
            matched.append(task_to_dict(t, mpd))

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
    mpd = _get_mpd(proj)

    groups = {}
    total  = 0

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        total += 1

        td = task_to_dict(t, mpd) if include_tasks else None

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
    mpd = _get_mpd(proj)

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
                tasks.append(task_to_dict(t, mpd))

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
