"""
Writing tasks: creation, update, deletion, scheduling mode and bulk edits.

Boundary against msp_tasks_read is mechanical: everything here writes to the
project. dry_run_bulk_update is the deliberate exception -- it mirrors the
write logic without executing it, so it belongs beside the code it simulates
rather than with the read tools.

REQUIRES WINDOWS
----------------
Five tools here drive the MS Project window rather than the object model,
using SelectRow plus EditCut/EditCopy/EditPaste or OutlineIndent/OutlineOutdent:

    add_recurring_task, copy_task_structure, delete_task, indent_task,
    move_task

They depend on selection state inside the running application, which no fake
reproduces. An offline test that exercises them will pass while doing nothing,
so they must be verified against a real MS Project. Four more tools elsewhere
carry the same warning: undo_last and insert_subproject in msp_projects,
apply_filter in msp_tasks_read, set_project_calendar in msp_calendars.

Four tools bracket their loops with app.Calculation = 0 and restore it in a
finally. Each pair is wholly inside one function; never slice one apart.
"""

import json

from msp_core import (
    responder,
    calculo_suspenso,
    _uid_map,
    mcp,
    get_app,
    get_proj,
    _find_task,
    _fmt_date,
    _get_mpd,
    _parse_date,
)


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
    Update a task by UniqueID. Only the arguments you pass are changed.
    
    Sentinels mean "leave alone": -1 for numbers, "" for strings, None for bools --
    so these cannot set a field to 0, "" or False.
    Dates YYYY-MM-DD. percent_complete 0-100. duration_days in working days.
    priority 0-1000. rag: 'Red'/'Amber'/'Green' (stored in Text1).
    task_type: 'FixedUnits', 'FixedDuration' or 'FixedWork'.
    manual: True = manually scheduled, False = auto.
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
        return responder({
            "status":   "updated",
            "unique_id": unique_id,
            "name":     t.Name,
            "changed":  changed,
        })

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

    uid_map = _uid_map(proj)

    results = []
    with calculo_suspenso(app):
        for item in items:
            uid = item["unique_id"]
            rag = item["rag"]
            if uid in uid_map:
                uid_map[uid].Text1 = rag
                results.append({"unique_id": uid, "status": "updated", "rag": rag})
            else:
                results.append({"unique_id": uid, "status": "not_found"})

    app.FileSave()
    return responder({"updated": len([r for r in results if r["status"] == "updated"]),
                       "results": results})


@mcp.tool()
def bulk_update_tasks(updates_json: str) -> str:
    """
    Update many tasks in one call, with auto-calculation suspended.
    
    updates_json: JSON list of objects, each with unique_id plus any of name,
    percent_complete, start, finish (YYYY-MM-DD), duration_days, notes, rag,
    text2, text3, flag1, flag2, priority, manual.
    Example: '[{"unique_id":12,"percent_complete":50,"rag":"Amber"}]'
    """
    items = json.loads(updates_json)
    app   = get_app()
    proj  = get_proj(app)
    mpd   = _get_mpd(proj)

    app.Calculation = 0
    try:
        uid_map = _uid_map(proj)
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
    return responder({
        "updated":   updated,
        "not_found": not_found,
    })


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
    Add a task to the active project.
    
    outline_level is the WBS depth (1 = top level). Dates YYYY-MM-DD, empty for
    none. duration_days in working days. rag is stored in Text1.
    after_unique_id is accepted and IGNORED -- tasks always append at the end.
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
    return responder({
        "status":    "created",
        "unique_id": task.UniqueID,
        "id":        task.ID,
        "name":      task.Name,
    })


@mcp.tool()
def bulk_add_tasks(tasks_json: str) -> str:
    """
    Add many tasks in one call, in order, with auto-calculation suspended.
    
    tasks_json: JSON list of objects. Fields: name (required), outline_level
    (default 1), start, finish (YYYY-MM-DD), duration_days, milestone, resource,
    rag, text2, text3, notes, manual.
    Example: '[{"name":"Phase 1","outline_level":1},{"name":"Task A",
    "outline_level":2,"start":"2026-04-01"}]'
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
    return responder({
        "created": len(created),
        "tasks":   created,
    })


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
            return responder({
                "status":    "deleted",
                "unique_id": unique_id,
                "name":      task_name,
            })

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


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
            return responder({
                "status":    "updated",
                "unique_id": unique_id,
                "name":      t.Name,
                "manual":    manual,
            })

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


@mcp.tool()
def bulk_set_task_mode(updates_json: str) -> str:
    """
    Switch tasks between manual and automatic scheduling in one call.
    
    updates_json: either a JSON list of {"unique_id":N,"manual":bool}, or a scope
    object -- {"manual":false,"all":true} for every task, or
    {"manual":false,"outline_level":4} for one level.
    Converting a large schedule to automatic reschedules the whole network: dates
    WILL move.
    """
    data = json.loads(updates_json)
    app  = get_app()
    proj = get_proj(app)

    updated = 0

    with calculo_suspenso(app):
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
            uid_map = _uid_map(proj)
            for item in items:
                uid = item["unique_id"]
                t = uid_map.get(uid)
                if t is not None:
                    t.Manual = item.get("manual", True)
                    updated += 1

    app.FileSave()
    return responder({"updated": updated})


@mcp.tool()
def set_constraint(unique_id: int, constraint_type: str = "SNET", constraint_date: str = "") -> str:
    """
    Set a scheduling constraint on a task.
    
    constraint_type: ASAP, ALAP, MSO, MFO, SNET (default), SNLT, FNET, FNLT.
    constraint_date YYYY-MM-DD, required for every type except ASAP and ALAP.
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
            return responder({
                "status":          "updated",
                "unique_id":       unique_id,
                "name":            t.Name,
                "constraint_type": ct,
                "constraint_date": constraint_date or "N/A",
            })

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
    return responder({"status": "cleared", "tasks_updated": count})


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
            return responder({
                "status":    "updated",
                "unique_id": unique_id,
                "name":      t.Name,
                "old_level": old_level,
                "new_level": t.OutlineLevel,
            })

    return json.dumps({"error": f"Task UniqueID {unique_id} not found."})


@mcp.tool()
def set_deadline(unique_id: int, deadline_date: str) -> str:
    """
    Set a deadline on a task, or pass 'clear' to remove it.
    
    A deadline is an indicator, not a constraint: it flags a late finish and does
    not move dates. deadline_date YYYY-MM-DD.
    """
    app  = get_app()
    proj = get_proj(app)

    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    if deadline_date.lower() == "clear":
        t.Deadline = "NA"
        app.FileSave()
        return responder({
            "status":    "cleared",
            "unique_id": unique_id,
            "name":      t.Name,
            "deadline":  None,
        })

    dl = _parse_date(deadline_date)
    t.Deadline = dl
    app.FileSave()

    deadline_missed = False
    try:
        if t.Finish and t.Finish > dl:
            deadline_missed = True
    except Exception:
        pass

    return responder({
        "status":          "set",
        "unique_id":       unique_id,
        "name":            t.Name,
        "deadline":        deadline_date,
        "finish":          _fmt_date(t.Finish),
        "deadline_missed": deadline_missed,
    })


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

    return responder({
        "status":    "updated",
        "unique_id": unique_id,
        "name":      t.Name,
        "active":    active,
    })


@mcp.tool()
def dry_run_bulk_update(updates_json: str) -> str:
    """
    Preview what bulk_update_tasks would change. Mutates nothing.
    
    updates_json takes the same shape bulk_update_tasks does: a JSON list of
    objects, each with unique_id plus the fields to change.
    """
    items = json.loads(updates_json)
    app   = get_app()
    proj  = get_proj(app)
    mpd   = _get_mpd(proj)

    changes   = []
    not_found = []
    no_change = []

    # One traversal for the whole batch instead of one per item.
    uid_map = _uid_map(proj)

    for item in items:
        uid = item.get("unique_id")
        if uid is None:
            continue

        t = uid_map.get(uid)
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

    return responder({
        "preview":              True,
        "changes":              changes,
        "not_found":            not_found,
        "no_change":            no_change,
        "total_changes":        total_changes,
        "total_tasks_affected": len(changes),
    })


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

    return responder({
        "status":    "moved",
        "unique_id": unique_id,
        "name":      moved.Name if moved else "(unknown)",
        "old_id":    old_id,
        "new_id":    new_id,
    })


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

    return responder({
        "status":       "copied",
        "source_name":  source.Name,
        "copied_tasks": all_copied,
    })


@mcp.tool()
def bulk_set_deadlines(deadlines_json: str) -> str:
    """
    Set deadlines on many tasks in one call.
    
    deadlines_json: JSON list of {"unique_id":N,"deadline":"YYYY-MM-DD"}.
    An empty deadline clears it. A deadline is an indicator, not a constraint: it
    does not move dates.
    Does NOT save -- call save_project if it must persist.
    """
    items = json.loads(deadlines_json)
    app   = get_app()
    proj  = get_proj(app)

    set_count = 0
    cleared = 0
    not_found = []
    errors = []

    # One traversal for the whole batch instead of one per item.
    uid_map = _uid_map(proj)

    with calculo_suspenso(app):
        for item in items:
            uid = item["unique_id"]
            dd  = item["deadline_date"]

            t = uid_map.get(uid)
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

    try:
        app.FileSave()
    except Exception:
        pass

    return responder({
        "set":       set_count,
        "cleared":   cleared,
        "not_found": not_found,
        "errors":    errors,
    })


@mcp.tool()
def set_task_hyperlink(unique_id: int, url: str, text: str = "", sub_address: str = "") -> str:
    """
    Set a hyperlink on a task. url may be a web address or a file path;
    sub_address is a bookmark within the target.
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
    return responder({
        "status":    "updated",
        "unique_id": unique_id,
        "name":      t.Name,
        "hyperlink": url,
        "text":      text or url,
    })


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
    Create a summary task with one subtask per occurrence.
    
    Simulated rather than native: Microsoft Project's recurring-task API is
    dialog-only. Requires python-dateutil.
    Dates YYYY-MM-DD. pattern: 'daily', 'weekly' or 'monthly'.
    weekday for weekly (0 = Monday), day_of_month for monthly.
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

        return responder({
            "status":          "created",
            "name":            name,
            "unique_id":       summary_uid,
            "recurrence_type": recurrence_type,
            "occurrences":     len(dates),
            "start":           start_date,
            "end":             end_date,
        })

    except ImportError:
        return json.dumps({"error": "python-dateutil is required for recurring tasks. Install with: pip install python-dateutil"})
    except Exception as e:
        return json.dumps({"error": f"Failed to create recurring task: {e}"})
