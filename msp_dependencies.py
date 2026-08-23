"""
Task precedence network.

Owns the links between tasks: creating and removing FS/SS/FF/SF dependencies
with lag, and walking the predecessor/successor graph.

Boundary: a tool belongs here when its subject is the LINK between tasks. A
tool that changes a single task's own scheduling -- its constraint, its mode,
its calendar -- belongs to its own domain even though the schedule shifts as a
result.
"""

import json

from msp_core import (
    responder,
    calculo_suspenso,
    _uid_map,
    _uid_to_id_map,
    mcp,
    get_app,
    get_proj,
    _find_task,
    _fmt_date,
    _get_mpd,
)


@mcp.tool()
def add_predecessor(
    successor_unique_id:   int,
    predecessor_unique_id: int,
    link_type:             str = "FS",
    lag_days:              int = 0,
) -> str:
    """
    Link a predecessor to a task.
    
    link_type: 'FS' (default), 'SS', 'FF' or 'SF'. lag_days may be negative.
    Both ids are UniqueIDs, not the row numbers that appear inside predecessor
    strings.
    """
    app  = get_app()
    proj = get_proj(app)

    uid_to_id = _uid_to_id_map(proj)

    if successor_unique_id not in uid_to_id:
        return json.dumps({"error": f"Successor UniqueID {successor_unique_id} not found."})
    if predecessor_unique_id not in uid_to_id:
        return json.dumps({"error": f"Predecessor UniqueID {predecessor_unique_id} not found."})

    pred_id = uid_to_id[predecessor_unique_id]
    succ_task = None
    for t in proj.Tasks:
        if t is not None and t.UniqueID == successor_unique_id:
            succ_task = t
            break

    lag_str = ""
    if lag_days > 0:
        lag_str = f"+{lag_days}d"
    elif lag_days < 0:
        lag_str = f"{lag_days}d"

    existing = succ_task.Predecessors.strip()
    new_pred  = f"{pred_id}{link_type}{lag_str}"

    if existing:
        succ_task.Predecessors = existing + "," + new_pred
    else:
        succ_task.Predecessors = new_pred

    app.FileSave()
    return responder({
        "status":       "linked",
        "successor":    successor_unique_id,
        "predecessor":  predecessor_unique_id,
        "link":         new_pred,
        "predecessors": succ_task.Predecessors,
    })


@mcp.tool()
def bulk_add_predecessors(links_json: str) -> str:
    """
    Add many predecessor links in one call.
    
    links_json: JSON list of {"successor_unique_id":N,"predecessor_unique_id":N,
    "link_type":"FS","lag_days":0}. link_type: FS (default), SS, FF, SF.
    """
    links = json.loads(links_json)
    app   = get_app()
    proj  = get_proj(app)

    uid_to_id = _uid_to_id_map(proj)
    uid_to_task = _uid_map(proj)

    linked = 0
    errors = []

    with calculo_suspenso(app):
        for link in links:
            succ_uid = link["successor_unique_id"]
            pred_uid = link["predecessor_unique_id"]
            lt       = link.get("link_type", "FS")
            lag      = link.get("lag_days", 0)

            if succ_uid not in uid_to_id:
                errors.append({"successor_unique_id": succ_uid, "error": "not found"})
                continue
            if pred_uid not in uid_to_id:
                errors.append({"predecessor_unique_id": pred_uid, "error": "not found"})
                continue

            pred_id   = uid_to_id[pred_uid]
            succ_task = uid_to_task[succ_uid]

            lag_str = ""
            if lag > 0:
                lag_str = f"+{lag}d"
            elif lag < 0:
                lag_str = f"{lag}d"

            new_pred = f"{pred_id}{lt}{lag_str}"
            existing = succ_task.Predecessors.strip()

            if existing:
                succ_task.Predecessors = existing + "," + new_pred
            else:
                succ_task.Predecessors = new_pred

            linked += 1

    app.FileSave()
    return responder({
        "linked": linked,
        "errors": errors,
    })


@mcp.tool()
def remove_predecessor(
    successor_unique_id:   int,
    predecessor_unique_id: int,
) -> str:
    """Remove a specific predecessor link from a task."""
    app  = get_app()
    proj = get_proj(app)

    uid_to_id = _uid_to_id_map(proj)

    if successor_unique_id not in uid_to_id:
        return json.dumps({"error": f"Successor UniqueID {successor_unique_id} not found."})

    pred_id   = uid_to_id.get(predecessor_unique_id)
    succ_task = None
    for t in proj.Tasks:
        if t is not None and t.UniqueID == successor_unique_id:
            succ_task = t
            break

    existing = succ_task.Predecessors.strip()
    if not existing:
        return json.dumps({"status": "no_change", "message": "Task has no predecessors."})

    parts     = [p.strip() for p in existing.split(",")]
    filtered  = [p for p in parts if not p.startswith(str(pred_id))]
    succ_task.Predecessors = ",".join(filtered)

    app.FileSave()
    return responder({
        "status":              "unlinked",
        "successor":           successor_unique_id,
        "removed_predecessor": predecessor_unique_id,
        "predecessors_now":    succ_task.Predecessors,
    })


@mcp.tool()
def get_task_dependencies(unique_id: int) -> str:
    """Get all predecessor and successor dependencies for a task."""
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    target = None
    for t in proj.Tasks:
        if t is not None and t.UniqueID == unique_id:
            target = t
            break

    if target is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    preds = []
    try:
        for dep in target.TaskDependencies:
            if dep.To.UniqueID == unique_id:
                preds.append({
                    "unique_id": dep.From.UniqueID,
                    "name":      dep.From.Name,
                    "type":      dep.Type,
                    "lag_days":  round(dep.Lag / mpd, 2),
                })
    except Exception:
        pass

    succs = []
    try:
        for dep in target.TaskDependencies:
            if dep.From.UniqueID == unique_id:
                succs.append({
                    "unique_id": dep.To.UniqueID,
                    "name":      dep.To.Name,
                    "type":      dep.Type,
                    "lag_days":  round(dep.Lag / mpd, 2),
                })
    except Exception:
        pass

    return responder({
        "task":        {"unique_id": unique_id, "name": target.Name},
        "predecessors": preds,
        "successors":   succs,
    })


@mcp.tool()
def get_dependency_chain(unique_id: int, direction: str = "successors", max_depth: int = 50) -> str:
    """
    Walk the dependency chain from a task -- what is downstream if this slips.
    
    direction: 'successors' (default) or 'predecessors'. max_depth caps the walk
    (default 50).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    root = _find_task(proj, unique_id)
    if root is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    visited = set()
    chain = []
    queue = [(root, 0)]  # (task, depth)
    visited.add(root.UniqueID)

    while queue:
        current, depth = queue.pop(0)
        if depth > 0:
            chain.append({
                "unique_id": current.UniqueID,
                "name":      current.Name,
                "depth":     depth,
                "start":     _fmt_date(current.Start),
                "finish":    _fmt_date(current.Finish),
                "critical":  bool(current.Critical),
            })
        if depth >= max_depth:
            continue

        try:
            for dep in current.TaskDependencies:
                if direction.lower() == "successors":
                    if dep.From.UniqueID == current.UniqueID:
                        next_task = dep.To
                    else:
                        continue
                else:
                    if dep.To.UniqueID == current.UniqueID:
                        next_task = dep.From
                    else:
                        continue

                if next_task.UniqueID not in visited:
                    visited.add(next_task.UniqueID)
                    # Add link info to the chain entry
                    link_type = dep.Type
                    lag_days = round(dep.Lag / mpd, 2) if dep.Lag else 0
                    queue.append((next_task, depth + 1))
                    # Update last chain entry with link info when it's added
        except Exception:
            pass

    return responder({
        "root":          {"unique_id": unique_id, "name": root.Name},
        "direction":     direction,
        "depth_reached": max(e["depth"] for e in chain) if chain else 0,
        "chain":         chain,
    })
