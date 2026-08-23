"""
Resources: the pool, assignments, availability and cost rates.

Reunites a CRUD that the original single file had split three ways --
add_resource near line 1234, update_resource near 3068 and delete_resource
near 3870.

Boundary: a tool belongs here when it manipulates a Resource object and does
not touch a Calendar. set_resource_calendar is the deliberate exception and
lives in msp_calendars: its subject is the calendar, the resource is only the
target. The rate-table tools belong here rather than with cost reporting
because they write to r.CostRateTables -- they edit the resource, they do not
report on money.

bulk_assign_resources brackets its loop with app.Calculation = 0 and restores
it in a finally. That pair is wholly inside the function, so moving the tool
carries it; it must never be sliced apart.
"""

import json

from msp_core import (
    _uid_map,
    _find_resource,
    TIMESCALE_MAP,
    mcp,
    get_app,
    get_proj,
    _find_task,
    _fmt_date,
    _parse_date,
)


@mcp.tool()
def get_resources() -> str:
    """Get all resources in the active project."""
    app  = get_app()
    proj = get_proj(app)

    results = []
    for r in proj.Resources:
        if r is not None:
            results.append({
                "unique_id":  r.UniqueID,
                "id":         r.ID,
                "name":       r.Name,
                "initials":   r.Initials,
                "type":       r.Type,
                "max_units":  r.MaxUnits,
                "cost":       r.StandardRate,
                "task_count": r.Assignments.Count,
            })

    return json.dumps({"count": len(results), "resources": results}, indent=2)


@mcp.tool()
def add_resource(
    name:          str,
    type:          int   = 0,
    max_units:     float = 1.0,
    standard_rate: str   = "",
    cost_per_use:  float = 0.0,
) -> str:
    """
    Add a resource to the project resource pool.

    Args:
        name:          Resource name (required).
        type:          0=Work (default), 1=Material, 2=Cost.
        max_units:     Maximum allocation units (default 1.0 = 100%).
        standard_rate: Standard rate as string (e.g. "50/h", "100/d").
        cost_per_use:  Fixed cost per use (default 0).
    """
    app  = get_app()
    proj = get_proj(app)

    r = proj.Resources.Add(name)
    r.Type = type
    if type == 0:  # MaxUnits only valid for Work resources
        r.MaxUnits = max_units
    if standard_rate:
        r.StandardRate = standard_rate
    if cost_per_use > 0:
        r.CostPerUse = cost_per_use

    app.FileSave()
    return json.dumps({
        "status":    "created",
        "unique_id": r.UniqueID,
        "id":        r.ID,
        "name":      r.Name,
        "type":      r.Type,
        "max_units": r.MaxUnits,
    }, indent=2)


@mcp.tool()
def assign_resource(task_unique_id: int, resource_name: str, units: float = 1.0) -> str:
    """
    Assign a resource to a task. If the resource doesn't exist, it is created.
    If the task already has resources, the new one is appended.

    Args:
        task_unique_id: Task UniqueID (required).
        resource_name:  Resource name to assign (required).
        units:          Allocation units, e.g. 1.0 = 100% (default 1.0).
    """
    app  = get_app()
    proj = get_proj(app)

    # Find task
    task = None
    for t in proj.Tasks:
        if t is not None and t.UniqueID == task_unique_id:
            task = t
            break
    if task is None:
        return json.dumps({"error": f"Task UniqueID {task_unique_id} not found."})

    # Check if resource exists, create if not
    res_exists = _find_resource(proj, resource_name) is not None
    if not res_exists:
        proj.Resources.Add(resource_name)

    # Append to ResourceNames (handles existing assignments)
    existing = (task.ResourceNames or "").strip()
    if existing:
        # Check if already assigned
        existing_names = [n.strip().lower() for n in existing.split(",")]
        if resource_name.lower() not in existing_names:
            task.ResourceNames = existing + "," + resource_name
    else:
        task.ResourceNames = resource_name

    app.FileSave()
    return json.dumps({
        "status":         "assigned",
        "task_unique_id": task_unique_id,
        "task_name":      task.Name,
        "resource_name":  resource_name,
        "resource_names": task.ResourceNames,
    }, indent=2)


@mcp.tool()
def get_resource_workload(resource_name: str, start_date: str = "", end_date: str = "") -> str:
    """
    Resource allocation view with conflict detection.
    Shows all assignments for a resource and identifies overlapping assignments.

    Args:
        resource_name: Resource name (case-insensitive exact match).
        start_date:    Filter assignments starting after this date (YYYY-MM-DD, optional).
        end_date:      Filter assignments ending before this date (YYYY-MM-DD, optional).
    """
    app  = get_app()
    proj = get_proj(app)

    # Find resource
    resource = _find_resource(proj, resource_name)

    if resource is None:
        # List available resources
        avail = []
        for r in proj.Resources:
            if r is not None:
                avail.append(r.Name)
        return json.dumps({"error": f"Resource '{resource_name}' not found. Available: {avail}"})

    filter_start = _parse_date(start_date) if start_date else None
    filter_end   = _parse_date(end_date) if end_date else None

    assignments = []
    for a in resource.Assignments:
        try:
            a_start  = a.Start
            a_finish = a.Finish
            task     = a.Task

            if filter_start and a_finish and a_finish < filter_start:
                continue
            if filter_end and a_start and a_start > filter_end:
                continue

            work_hours = 0
            try:
                work_hours = round(a.Work / 60, 2)  # minutes to hours
            except Exception:
                pass

            assignments.append({
                "task_unique_id": task.UniqueID if task else None,
                "task_name":      task.Name if task else "(unknown)",
                "start":          _fmt_date(a_start),
                "finish":         _fmt_date(a_finish),
                "units":          a.Units,
                "work_hours":     work_hours,
            })
        except Exception:
            continue

    # Conflict detection: find overlapping date ranges
    conflicts = []
    for i in range(len(assignments)):
        for j in range(i + 1, len(assignments)):
            a = assignments[i]
            b = assignments[j]
            if a["start"] and a["finish"] and b["start"] and b["finish"]:
                if a["start"] <= b["finish"] and b["start"] <= a["finish"]:
                    overlap_start = max(a["start"], b["start"])
                    overlap_finish = min(a["finish"], b["finish"])
                    combined = (a.get("units") or 0) + (b.get("units") or 0)
                    conflicts.append({
                        "task_a":          a["task_name"],
                        "task_b":          b["task_name"],
                        "overlap_start":   overlap_start,
                        "overlap_finish":  overlap_finish,
                        "combined_units":  combined,
                    })

    overallocated = False
    try:
        overallocated = bool(resource.Overallocated)
    except Exception:
        pass

    max_units = 1.0
    try:
        max_units = resource.MaxUnits
    except Exception:
        pass

    return json.dumps({
        "resource":      resource.Name,
        "overallocated": overallocated,
        "max_units":     max_units,
        "assignments":   assignments,
        "conflicts":     conflicts,
    }, indent=2)


@mcp.tool()
def bulk_assign_resources(assignments_json: str) -> str:
    """
    Assign resources to multiple tasks in one call.

    Args:
        assignments_json: JSON string — list of {task_unique_id, resource_name, units (optional)}.
            Example: '[{"task_unique_id": 42, "resource_name": "Alice"},
                       {"task_unique_id": 55, "resource_name": "Bob", "units": 0.5}]'
    """
    items = json.loads(assignments_json)
    app   = get_app()
    proj  = get_proj(app)

    app.Calculation = 0
    try:
        uid_map = _uid_map(proj)
        # Get existing resource names
        existing_resources = set()
        for r in proj.Resources:
            if r is not None:
                existing_resources.add(r.Name.lower())

        assigned = 0
        errors = []
        created_resources = []

        for item in items:
            task_uid = item["task_unique_id"]
            res_name = item["resource_name"]

            t = uid_map.get(task_uid)
            if t is None:
                errors.append({"task_unique_id": task_uid, "error": "task not found"})
                continue

            # Create resource if needed
            if res_name.lower() not in existing_resources:
                proj.Resources.Add(res_name)
                existing_resources.add(res_name.lower())
                created_resources.append(res_name)

            # Append to ResourceNames
            existing = (t.ResourceNames or "").strip()
            if existing:
                existing_names = [n.strip().lower() for n in existing.split(",")]
                if res_name.lower() not in existing_names:
                    t.ResourceNames = existing + "," + res_name
            else:
                t.ResourceNames = res_name

            assigned += 1
    finally:
        app.CalculateProject()
        app.Calculation = -1

    try:
        app.FileSave()
    except Exception:
        pass

    return json.dumps({
        "assigned":          assigned,
        "errors":            errors,
        "created_resources": created_resources,
    }, indent=2)


@mcp.tool()
def remove_resource_assignment(task_unique_id: int, resource_name: str) -> str:
    """
    Remove a specific resource from a task.

    Args:
        task_unique_id: Task UniqueID (required).
        resource_name:  Resource name to remove (required).
    """
    app  = get_app()
    proj = get_proj(app)

    t = _find_task(proj, task_unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {task_unique_id} not found."})

    existing = (t.ResourceNames or "").strip()
    if not existing:
        return json.dumps({"error": f"Task '{t.Name}' has no resources assigned."})

    names = [n.strip() for n in existing.split(",")]
    filtered = [n for n in names if n.lower() != resource_name.lower()]

    if len(filtered) == len(names):
        return json.dumps({"error": f"Resource '{resource_name}' not assigned to task '{t.Name}'. Current: {existing}"})

    t.ResourceNames = ",".join(filtered) if filtered else ""

    try:
        app.FileSave()
    except Exception:
        pass

    return json.dumps({
        "status":           "removed",
        "task_name":        t.Name,
        "removed":          resource_name,
        "resource_names_now": t.ResourceNames,
    }, indent=2)


@mcp.tool()
def update_resource(resource_name: str, new_name: str = "", max_units: float = -1, standard_rate: str = "", cost_per_use: float = -1) -> str:
    """
    Modify an existing resource's properties.

    Args:
        resource_name: Current resource name (required, case-insensitive).
        new_name:      New name for the resource (optional).
        max_units:     Maximum allocation units, e.g. 2.0 = 200% (optional, -1 = no change).
        standard_rate: Standard rate as string, e.g. '50/h' (optional).
        cost_per_use:  Fixed cost per use (optional, -1 = no change).
    """
    app  = get_app()
    proj = get_proj(app)

    resource = _find_resource(proj, resource_name)

    if resource is None:
        avail = [r.Name for r in proj.Resources if r is not None]
        return json.dumps({"error": f"Resource '{resource_name}' not found. Available: {avail}"})

    changed = []

    if new_name:
        resource.Name = new_name
        changed.append("name")
    if max_units >= 0:
        resource.MaxUnits = max_units
        changed.append("max_units")
    if standard_rate:
        resource.StandardRate = standard_rate
        changed.append("standard_rate")
    if cost_per_use >= 0:
        resource.CostPerUse = cost_per_use
        changed.append("cost_per_use")

    try:
        app.FileSave()
    except Exception:
        pass

    return json.dumps({
        "status":  "updated",
        "name":    resource.Name,
        "changed": changed,
    }, indent=2)


@mcp.tool()
def delete_resource(resource_name: str) -> str:
    """
    Delete a resource from the project pool (case-insensitive match).
    All task assignments referencing this resource are cleared first.
    """
    app  = get_app()
    proj = get_proj(app)

    target = _find_resource(proj, resource_name)

    if target is None:
        return json.dumps({"error": f"Resource '{resource_name}' not found."})

    # Clear assignments first
    assignments_cleared = 0
    # Iterate in reverse to avoid index shifting
    for i in range(target.Assignments.Count, 0, -1):
        try:
            target.Assignments(i).Delete()
            assignments_cleared += 1
        except Exception:
            pass

    name = target.Name
    try:
        target.Delete()
    except Exception as e:
        return json.dumps({"error": f"Failed to delete resource: {e}"})

    app.FileSave()
    return json.dumps({
        "status":              "deleted",
        "name":                name,
        "assignments_cleared": assignments_cleared,
    }, indent=2)


@mcp.tool()
def get_resource_availability(
    resource_name: str,
    start_date:    str,
    end_date:      str,
    timescale:     str = "weekly",
) -> str:
    """
    Show resource allocation vs capacity per period. Shows max units, allocated
    work, and free capacity windows.

    Args:
        resource_name: Name of the resource.
        start_date:    Period start as YYYY-MM-DD.
        end_date:      Period end as YYYY-MM-DD.
        timescale:     'daily', 'weekly', or 'monthly' (default 'weekly').
    """
    app  = get_app()
    proj = get_proj(app)

    ts = TIMESCALE_MAP.get(timescale.lower())
    if ts is None:
        return json.dumps({"error": f"Unknown timescale '{timescale}'. Use: daily, weekly, monthly."})

    res = _find_resource(proj, resource_name)
    if res is None:
        return json.dumps({"error": f"Resource '{resource_name}' not found."})

    sd = _parse_date(start_date)
    ed = _parse_date(end_date)

    max_units = res.MaxUnits  # e.g. 1.0 = 100%

    periods = []
    try:
        # Resource TimeScaleData types differ from Task types:
        # Type 13 = pjResourceTimescaledWork (minutes), Type 4 = Availability (units)
        tsd = res.TimeScaleData(sd, ed, 13, ts)
        for item in tsd:
            try:
                val = item.Value
                allocated_hrs = float(val) / 60.0 if val else 0.0
            except Exception:
                allocated_hrs = 0.0
            periods.append({
                "start":          _fmt_date(item.StartDate),
                "end":            _fmt_date(item.EndDate),
                "allocated_hours": round(allocated_hrs, 2),
            })
    except Exception as e:
        return json.dumps({"error": f"TimeScaleData failed: {e}"})

    return json.dumps({
        "resource":  res.Name,
        "max_units": max_units,
        "timescale": timescale,
        "periods":   periods,
    }, indent=2)


@mcp.tool()
def get_resource_rate_tables(resource_name: str) -> str:
    """
    Get cost rate tables (A through E) for a resource.

    Args:
        resource_name: Name of the resource.
    """
    app  = get_app()
    proj = get_proj(app)

    res = _find_resource(proj, resource_name)
    if res is None:
        return json.dumps({"error": f"Resource '{resource_name}' not found."})

    tables = {}
    TABLE_NAMES = ["A", "B", "C", "D", "E"]

    for idx, tname in enumerate(TABLE_NAMES):
        try:
            table = res.CostRateTables(idx + 1)
            rates = []
            for pay_rate in table.PayRates:
                rates.append({
                    "effective_date":  _fmt_date(pay_rate.EffectiveDate),
                    "standard_rate":   str(pay_rate.StandardRate),
                    "overtime_rate":   str(pay_rate.OvertimeRate),
                    "cost_per_use":    float(pay_rate.CostPerUse) if pay_rate.CostPerUse else 0.0,
                })
            tables[tname] = rates
        except Exception:
            tables[tname] = []

    return json.dumps({
        "resource": res.Name,
        "tables":   tables,
    }, indent=2)


@mcp.tool()
def set_resource_rate_table(
    resource_name: str,
    table:         str = "A",
    standard_rate: str = "",
    overtime_rate: str = "",
    cost_per_use:  float = -1,
    effective_date: str = "",
) -> str:
    """
    Set or add a cost rate entry in a resource's rate table.

    Args:
        resource_name:  Name of the resource.
        table:          Rate table letter: A, B, C, D, or E (default A).
        standard_rate:  Standard rate as string, e.g. '50/h' or '400/d'.
        overtime_rate:  Overtime rate as string, e.g. '75/h'.
        cost_per_use:   Per-use cost (default -1 = don't change).
        effective_date: When this rate takes effect (YYYY-MM-DD). Empty = first entry.
    """
    app  = get_app()
    proj = get_proj(app)

    TABLE_MAP = {"A": 1, "B": 2, "C": 3, "D": 4, "E": 5}
    tbl_idx = TABLE_MAP.get(table.upper())
    if tbl_idx is None:
        return json.dumps({"error": f"Invalid table '{table}'. Use A-E."})

    res = _find_resource(proj, resource_name)
    if res is None:
        return json.dumps({"error": f"Resource '{resource_name}' not found."})

    try:
        rate_table = res.CostRateTables(tbl_idx)
        pay_rates  = rate_table.PayRates

        if effective_date:
            # Add a new rate entry with effective date
            ed = _parse_date(effective_date)
            new_rate = pay_rates.Add(ed)
            if standard_rate:
                new_rate.StandardRate = standard_rate
            if overtime_rate:
                new_rate.OvertimeRate = overtime_rate
            if cost_per_use >= 0:
                new_rate.CostPerUse = cost_per_use
        else:
            # Update the first (default) rate entry
            first = pay_rates(1)
            if standard_rate:
                first.StandardRate = standard_rate
            if overtime_rate:
                first.OvertimeRate = overtime_rate
            if cost_per_use >= 0:
                first.CostPerUse = cost_per_use

        app.FileSave()
        return json.dumps({
            "status":   "updated",
            "resource": res.Name,
            "table":    table.upper(),
            "standard_rate": standard_rate or "(unchanged)",
            "overtime_rate": overtime_rate or "(unchanged)",
            "cost_per_use":  cost_per_use if cost_per_use >= 0 else "(unchanged)",
        }, indent=2)

    except Exception as e:
        return json.dumps({"error": f"Failed to update rate table: {e}"})
