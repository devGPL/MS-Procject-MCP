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

Twelve of these tools read through a VIEW rather than through COM directly:
they take whatever msp_fast.varredura hands them -- dicts parsed from the
saved .mpp, or COM tasks wrapped so they answer to the same key names -- and
the body cannot tell which. Every one of them reports which backend served it
under "source". The four that remain on COM alone need fields no view carries:
baselines (get_milestone_report), work (get_actual_work), timephased series,
or the view itself (apply_filter).
"""

import json
import datetime

import msp_fast

from msp_core import (
    TIMESCALE_MAP,
    mcp,
    get_app,
    get_proj,
    tarefa_completa,
    vista_com,
    _dia,
    _dt_de_vista,
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

    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)

    alvo = keyword.lower()
    results = []
    for v in fonte:
        if not include_summary and v["summary"]:
            continue
        if outline_level > 0 and v["outline_level"] != outline_level:
            continue
        if alvo and alvo not in (v["name"] or "").lower():
            continue
        results.append(tarefa_completa(v))

    return json.dumps({"count": len(results), "source": origem,
                       "tasks": results}, indent=2)


@mcp.tool()
def get_task(unique_id: int) -> str:
    """Get full details for a single task by its UniqueID."""
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    # Reads like a cheap lookup and is not: there is no index by UniqueID on
    # either side, so both backends scan. The scan is what the fast path pays
    # for -- a task near the end of an 8,000-row schedule is 8,000 COM object
    # fetches away.
    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)

    for v in fonte:
        if v["unique_id"] == unique_id:
            saida = tarefa_completa(v)
            saida["source"] = origem
            return json.dumps(saida, indent=2)

    return json.dumps({"error": f"Task UniqueID {unique_id} not found.",
                       "source": origem})


@mcp.tool()
def get_tasks_by_rag(rag: str = "Red") -> str:
    """
    Return tasks filtered by RAG status stored in the Text1 custom field.
    rag: 'Red', 'Amber', or 'Green'
    """
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)

    alvo = rag.strip().lower()
    results = []
    for v in fonte:
        if v["summary"]:
            continue
        if (v["text1"] or "").strip().lower() == alvo:
            results.append(tarefa_completa(v))

    return json.dumps({"rag": rag, "count": len(results), "source": origem,
                       "tasks": results}, indent=2)


@mcp.tool()
def get_overdue_tasks() -> str:
    """Return incomplete tasks whose Finish date is in the past."""
    today = datetime.datetime.now()
    app   = get_app()
    proj  = get_proj(app)
    mpd  = _get_mpd(proj)

    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)

    results = []
    for v in fonte:
        if v["summary"] or v["milestone"]:
            continue
        if (v["percent_complete"] or 0) >= 100:
            continue
        finish = _dt_de_vista(v["finish"])
        if finish and finish < today:
            results.append(tarefa_completa(v))

    return json.dumps({"count": len(results), "source": origem,
                       "tasks": results}, indent=2)


@mcp.tool()
def get_tasks_by_resource(resource_name: str) -> str:
    """Return all tasks assigned to a named resource (case-insensitive substring match)."""
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)

    name_lower = resource_name.lower()
    results = []
    for v in fonte:
        if v["summary"]:
            continue
        if name_lower in (v["resource_names"] or "").lower():
            results.append(tarefa_completa(v))

    return json.dumps({
        "resource": resource_name,
        "count":    len(results),
        "source":   origem,
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
    today = datetime.datetime.now()
    app   = get_app()
    proj  = get_proj(app)
    mpd   = _get_mpd(proj)

    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)

    not_started = in_progress = complete = 0
    rag_counts  = {"Red": 0, "Amber": 0, "Green": 0, "Other": 0}
    overdue     = critical = 0

    for v in fonte:
        if v["summary"]:
            continue

        pct = v["percent_complete"] or 0
        if pct == 0:
            not_started += 1
        elif pct < 100:
            in_progress += 1
        else:
            complete += 1

        rag = (v["text1"] or "").strip()
        if rag in rag_counts:
            rag_counts[rag] += 1
        elif rag:
            rag_counts["Other"] += 1

        if v["critical"]:
            critical += 1

        fin = _dt_de_vista(v["finish"])
        if fin and fin < today and pct < 100:
            overdue += 1

    total = not_started + in_progress + complete

    return json.dumps({
        "project":     proj.Name,
        "source":      origem,
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

    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)

    # Build flat list first. Order carries the hierarchy here -- the tree below
    # is built by walking this list against a stack of parents -- which is why
    # the fast path sorts by ID before returning, the order COM enumerates in.
    flat = []
    for v in fonte:
        if max_level > 0 and v["outline_level"] > max_level:
            continue
        flat.append({
            "unique_id":     v["unique_id"],
            "id":            v["id"],
            "name":          v["name"],
            "level":         v["outline_level"],
            "summary":       bool(v["summary"]),
            "milestone":     bool(v["milestone"]),
            "start":         _dia(v["start"]),
            "finish":        _dia(v["finish"]),
            "duration_days": v["duration_days"],
            "children":      [],
        })

    # Build tree using stack
    root = {"name": proj.Name, "level": 0, "source": origem, "children": []}
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

    # Predicates read a view by published key name, so the same predicate runs
    # against a dict parsed from the file and against a COM task -- and the COM
    # side still reads only the fields a filter actually mentions, one round
    # trip each, cached.
    predicates = []

    def texto_igual(chave, valor):
        alvo = valor.lower()
        return lambda v: (v[chave] or "").strip().lower() == alvo

    if "rag" in f:
        predicates.append(texto_igual("text1", f["rag"]))
    if "resource" in f:
        alvo = f["resource"].lower()
        predicates.append(lambda v, _a=alvo: _a in (v["resource_names"] or "").lower())
    if "start_after" in f:
        d = _parse_date(f["start_after"])
        predicates.append(lambda v, _d=d: (_dt_de_vista(v["start"]) or datetime.datetime.min) >= _d)
    if "start_before" in f:
        d = _parse_date(f["start_before"])
        predicates.append(lambda v, _d=d: (_dt_de_vista(v["start"]) or datetime.datetime.max) <= _d)
    if "finish_after" in f:
        d = _parse_date(f["finish_after"])
        predicates.append(lambda v, _d=d: (_dt_de_vista(v["finish"]) or datetime.datetime.min) >= _d)
    if "finish_before" in f:
        d = _parse_date(f["finish_before"])
        predicates.append(lambda v, _d=d: (_dt_de_vista(v["finish"]) or datetime.datetime.max) <= _d)
    if "min_pct" in f:
        predicates.append(lambda v, _x=f["min_pct"]: (v["percent_complete"] or 0) >= _x)
    if "max_pct" in f:
        predicates.append(lambda v, _x=f["max_pct"]: (v["percent_complete"] or 0) <= _x)
    if "outline_level" in f:
        predicates.append(lambda v, _x=f["outline_level"]: v["outline_level"] == _x)
    if "critical" in f:
        predicates.append(lambda v, _x=f["critical"]: bool(v["critical"]) == _x)
    if "milestone" in f:
        predicates.append(lambda v, _x=f["milestone"]: bool(v["milestone"]) == _x)
    if "active" in f:
        predicates.append(lambda v, _x=f["active"]: bool(v["active"]) == _x)
    if "summary" in f:
        predicates.append(lambda v, _x=f["summary"]: bool(v["summary"]) == _x)
    if "name_contains" in f:
        alvo = f["name_contains"].lower()
        predicates.append(lambda v, _a=alvo: _a in (v["name"] or "").lower())
    if "text1" in f:
        predicates.append(texto_igual("text1", f["text1"]))
    if "text2" in f:
        predicates.append(texto_igual("text2", f["text2"]))
    if "text3" in f:
        predicates.append(texto_igual("text3", f["text3"]))
    if "flag1" in f:
        predicates.append(lambda v, _x=f["flag1"]: bool(v["flag1"]) == _x)
    if "flag2" in f:
        predicates.append(lambda v, _x=f["flag2"]: bool(v["flag2"]) == _x)

    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)

    # Collect matching tasks
    matched = []
    for v in fonte:
        if all(p(v) for p in predicates):
            matched.append(tarefa_completa(v))

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
        "source":         origem,
        "tasks":          matched,
    }, indent=2)


@mcp.tool()
def group_tasks_by(field: str, include_tasks: bool = False) -> str:
    """
    Group non-summary tasks by a field and return counts per group.

    Args:
        field:         Field to group by: 'rag', 'resource', 'outline_level', 'critical',
                       'milestone', 'percent_complete', 'text1', 'text2', 'text3',
                       'flag1', 'flag2'. Any other name is looked up among the
                       fields a task is published with (the keys get_task
                       returns), and yields '(unknown)' when there is none.
        include_tasks: If true, include task list per group (default false).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)

    groups = {}
    total  = 0

    for v in fonte:
        if v["summary"]:
            continue
        total += 1

        td = tarefa_completa(v) if include_tasks else None

        if field == "resource":
            # Split comma-separated resource names
            names = [n.strip() for n in (v["resource_names"] or "").split(",") if n.strip()]
            if not names:
                names = ["(unassigned)"]
            keys = names
        elif field == "percent_complete":
            pct = v["percent_complete"] or 0
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
            keys = [(v["text1"] or "").strip() or "(blank)"]
        elif field == "text2":
            keys = [(v["text2"] or "").strip() or "(blank)"]
        elif field == "text3":
            keys = [(v["text3"] or "").strip() or "(blank)"]
        elif field == "outline_level":
            keys = [str(v["outline_level"])]
        elif field == "critical":
            keys = [str(bool(v["critical"]))]
        elif field == "milestone":
            keys = [str(bool(v["milestone"]))]
        elif field == "flag1":
            keys = [str(bool(v["flag1"]))]
        elif field == "flag2":
            keys = [str(bool(v["flag2"]))]
        else:
            keys = [str(v.get(field, "(unknown)"))]

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
        "source":      origem,
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
    mpd  = _get_mpd(proj)

    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)

    branches = []
    # Materialised because every branch walks forward from its own position to
    # find its children -- the list is read many times over, not once.
    tasks_list = list(fonte)

    for i, t in enumerate(tasks_list):
        if not t["summary"]:
            continue
        if t["outline_level"] > max_level:
            continue

        # Count children
        child_count = 0
        milestones_complete = 0
        milestones_total = 0
        for j in range(i + 1, len(tasks_list)):
            child = tasks_list[j]
            if child["outline_level"] <= t["outline_level"]:
                break
            if not child["summary"]:
                child_count += 1
                if child["milestone"]:
                    milestones_total += 1
                    if (child["percent_complete"] or 0) >= 100:
                        milestones_complete += 1

        branches.append({
            "unique_id":          t["unique_id"],
            "name":               t["name"],
            "level":              t["outline_level"],
            "percent_complete":   t["percent_complete"],
            "start":              _dia(t["start"]),
            "finish":             _dia(t["finish"]),
            "child_count":        child_count,
            "milestones_complete": milestones_complete,
            "milestones_total":   milestones_total,
        })

    return json.dumps({
        "max_level": max_level,
        "source":    origem,
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
        origem = result.get("source")
    else:
        lidas, origem = msp_fast.varredura(proj)
        fonte = lidas if lidas is not None else vista_com(proj, mpd)
        tasks = [tarefa_completa(v) for v in fonte]

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
        "source":  origem,
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
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)

    results = []
    for v in fonte:
        if v["summary"]:
            continue
        # Both backends publish the constraint already named, and both map an
        # unrecognised value to "ASAP" -- so a constraint neither reader knows
        # is now reported as absent rather than as "Unknown(9)". The eight
        # values are the whole of the COM enum; a ninth would be a new
        # Microsoft Project.
        if v["constraint_type"] != "ASAP":
            results.append({
                "unique_id":       v["unique_id"],
                "name":            v["name"],
                "constraint_type": v["constraint_type"],
                "constraint_date": _dia(v["constraint_date"]),
            })

    return json.dumps({"count": len(results), "source": origem,
                       "tasks": results}, indent=2)


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
