"""
Project lifecycle, multi-project navigation, interchange and infrastructure.

Owns the .mpp file itself: opening, creating, saving, closing, describing;
moving between open projects; XML/JSON interchange and subproject insertion;
plus undo and the COM connectivity probe.

Boundary: a tool belongs here when its subject is the project FILE or the
server's connection to it. Tools whose subject is a Task, Resource or
Calendar live in their own module even when they save the file as a side
effect.
"""

import json

from msp_core import (
    _com_retry,
    mcp,
    get_app,
    get_proj,
    task_to_dict,
    _count_resources,
    _find_task,
    _fmt_date,
    _get_mpd,
    _parse_date,
)


@mcp.tool()
def open_project(file_path: str) -> str:
    """
    Open a Microsoft Project file (.mpp or .xml).
    MS Project must already be running (it is launched automatically if not).
    """
    import win32com.client
    try:
        app = win32com.client.GetActiveObject("MSProject.Application")
    except Exception:
        app = win32com.client.Dispatch("MSProject.Application")
        app.Visible = True
        app.DisplayAlerts = False

    app.FileOpen(file_path)
    proj = app.ActiveProject
    return json.dumps({
        "status":     "opened",
        "name":       proj.Name,
        "full_path":  proj.FullName,
        "task_count": proj.Tasks.Count,
        "start":      str(proj.ProjectStart)[:10],
        "finish":     str(proj.ProjectFinish)[:10],
    }, indent=2)


@mcp.tool()
def new_project(title: str = "New Project", start: str = "") -> str:
    """
    Create a new blank project without needing a file on disk.

    Args:
        title: Project title (default "New Project").
        start: Project start date as YYYY-MM-DD (optional).
    """
    import win32com.client
    try:
        app = win32com.client.GetActiveObject("MSProject.Application")
    except Exception:
        app = win32com.client.Dispatch("MSProject.Application")
        app.Visible = True
        app.DisplayAlerts = False

    app.FileNew()
    proj = app.ActiveProject
    proj.Title = title
    if start:
        proj.ProjectStart = _parse_date(start)

    return json.dumps({
        "status": "created",
        "title":  proj.Title,
        "name":   proj.Name,
        "start":  str(proj.ProjectStart)[:10],
    }, indent=2)


@mcp.tool()
def get_project_info() -> str:
    """Get summary information about the currently active project."""
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    # One pass, four counters.
    #
    # These were four separate comprehensions, which read well and cost four
    # full traversals of the task collection. Enumerating that collection is
    # the expensive part of any COM scan -- measured at 2.6 ms per task against
    # 0.24 ms for reading a property -- so writing the counters as four
    # one-liners multiplied precisely the part that costs. On an 8,429-task
    # project that was roughly 90 seconds where 22 will do.
    task_count = summary_count = mile_count = critical = 0
    for t in proj.Tasks:
        if t is None:
            continue
        task_count += 1
        eh_resumo = False
        try:
            eh_resumo = bool(t.Summary)
        except Exception:
            pass
        if eh_resumo:
            summary_count += 1
        try:
            if t.Milestone:
                mile_count += 1
        except Exception:
            pass
        if not eh_resumo:
            try:
                if t.Critical:
                    critical += 1
            except Exception:
                pass

    # Safe reads for optional metadata
    def safe_read(attr):
        try:
            return getattr(proj, attr, "") or ""
        except Exception:
            return ""

    return json.dumps({
        "name":            proj.Name,
        "full_path":       proj.FullName,
        "title":           safe_read("Title"),
        "manager":         safe_read("Manager"),
        "company":         safe_read("Company"),
        "author":          safe_read("Author"),
        "subject":         safe_read("Subject"),
        "start":           _fmt_date(proj.ProjectStart),
        "finish":          _fmt_date(proj.ProjectFinish),
        "status_date":     _fmt_date(getattr(proj, "StatusDate", None)),
        "calendar":        str(proj.Calendar) if proj.Calendar else "",
        "tasks_total":     task_count,
        "summary_tasks":   summary_count,
        "milestones":      mile_count,
        "critical_tasks":  critical,
        "resources":       _count_resources(proj),
        "minutes_per_day": mpd,
    }, indent=2)


@mcp.tool()
def set_project_properties(properties_json: str) -> str:
    """
    Set project metadata properties.

    Args:
        properties_json: JSON string with fields to set. All optional:
            title, manager, company, author, subject, status_date (YYYY-MM-DD),
            start (YYYY-MM-DD).
            Example: '{"title": "EXPO 2030", "manager": "John", "company": "ERC"}'
    """
    props = json.loads(properties_json)
    app   = get_app()
    proj  = get_proj(app)

    changed = []
    if "title" in props:
        proj.Title = props["title"];       changed.append("title")
    if "manager" in props:
        proj.Manager = props["manager"];   changed.append("manager")
    if "company" in props:
        proj.Company = props["company"];   changed.append("company")
    if "author" in props:
        proj.Author = props["author"];     changed.append("author")
    if "subject" in props:
        proj.Subject = props["subject"];   changed.append("subject")
    if "start" in props and props["start"]:
        proj.ProjectStart = _parse_date(props["start"]); changed.append("start")
    if "status_date" in props and props["status_date"]:
        proj.StatusDate = _parse_date(props["status_date"]); changed.append("status_date")

    app.FileSave()
    return json.dumps({"status": "updated", "changed": changed}, indent=2)


@mcp.tool()
def save_project() -> str:
    """Save the active project (in place)."""
    app = get_app()
    app.FileSave()
    return "Project saved."


@mcp.tool()
def save_project_as(file_path: str, format: str = "mpp") -> str:
    """
    Save the active project to a new path.
    format: 'mpp' (default), 'xml', 'csv'
    """
    fmt_map = {"mpp": 0, "xml": 22, "csv": 23}
    fmt_id  = fmt_map.get(format.lower(), 0)
    app     = get_app()
    app.FileSaveAs(Name=file_path, Format=fmt_id, Backup=False, ReadOnly=False)
    return f"Project saved as: {file_path}"


@mcp.tool()
def close_project(save: bool = False) -> str:
    """Close the active project. Set save=True to save before closing."""
    app = get_app()
    app.FileClose(Save=1 if save else 0)
    return "Project closed."


@mcp.tool()
def import_xml(file_path: str) -> str:
    """
    Open an MS Project XML file (e.g. the consolidated EXPO 2030 roadmap).
    MS Project must be running.
    """
    return open_project(file_path)


@mcp.tool()
def export_xml(output_path: str) -> str:
    """Export the active project to MS Project XML format."""
    return save_project_as(output_path, format="xml")


@mcp.tool()
def list_projects() -> str:
    """
    List all open projects in MS Project with metadata.
    Returns project count, which is active, and details for each.
    """
    app = get_app(require_project=False)

    projects = []
    active_name = ""
    try:
        active_name = app.ActiveProject.Name
    except Exception:
        pass

    for i in range(1, app.Projects.Count + 1):
        p = app.Projects(i)
        projects.append({
            "index":      i,
            "name":       p.Name,
            "full_path":  p.FullName,
            "task_count": p.Tasks.Count,
            "start":      _fmt_date(p.ProjectStart),
            "finish":     _fmt_date(p.ProjectFinish),
            "is_active":  p.Name == active_name,
        })

    return json.dumps({
        "count":    len(projects),
        "active":   active_name,
        "projects": projects,
    }, indent=2)


@mcp.tool()
def switch_project(name_or_index: str) -> str:
    """
    Switch the active project by name (substring match) or 1-based index.

    Args:
        name_or_index: Project name (case-insensitive substring) or numeric index.
    """
    app = get_app(require_project=False)

    if app.Projects.Count == 0:
        return json.dumps({"error": "No projects are open."})

    # Try as index first
    try:
        idx = int(name_or_index)
        if 1 <= idx <= app.Projects.Count:
            p = app.Projects(idx)
            p.Activate()
            return json.dumps({
                "status":     "switched",
                "name":       p.Name,
                "full_path":  p.FullName,
                "task_count": p.Tasks.Count,
                "start":      _fmt_date(p.ProjectStart),
                "finish":     _fmt_date(p.ProjectFinish),
            }, indent=2)
        else:
            return json.dumps({"error": f"Index {idx} out of range. Projects: 1-{app.Projects.Count}."})
    except ValueError:
        pass

    # Substring match on name
    query = name_or_index.lower()
    matches = []
    for i in range(1, app.Projects.Count + 1):
        p = app.Projects(i)
        if query in p.Name.lower():
            matches.append((i, p))

    if len(matches) == 0:
        names = [app.Projects(i).Name for i in range(1, app.Projects.Count + 1)]
        return json.dumps({"error": f"No project matching '{name_or_index}'. Open: {names}"})
    if len(matches) > 1:
        names = [m[1].Name for m in matches]
        return json.dumps({"error": f"Multiple matches for '{name_or_index}': {names}. Be more specific."})

    _, p = matches[0]
    try:
        active = app.ActiveProject.Name
        if active == p.Name:
            return json.dumps({
                "status":     "already_active",
                "name":       p.Name,
                "full_path":  p.FullName,
                "task_count": p.Tasks.Count,
                "start":      _fmt_date(p.ProjectStart),
                "finish":     _fmt_date(p.ProjectFinish),
            }, indent=2)
    except Exception:
        pass

    p.Activate()
    return json.dumps({
        "status":     "switched",
        "name":       p.Name,
        "full_path":  p.FullName,
        "task_count": p.Tasks.Count,
        "start":      _fmt_date(p.ProjectStart),
        "finish":     _fmt_date(p.ProjectFinish),
    }, indent=2)


@mcp.tool()
def cross_project_link(source_project: str, source_unique_id: int, target_project: str, target_unique_id: int, link_type: str = "FS") -> str:
    """
    Create a dependency link across open projects.

    Args:
        source_project:    Name of the predecessor's project.
        source_unique_id:  UniqueID of the predecessor task.
        target_project:    Name of the successor's project.
        target_unique_id:  UniqueID of the successor task.
        link_type:         'FS' (default), 'SS', 'FF', or 'SF'.
    """
    app = get_app(require_project=False)

    # Find source project and task
    src_proj = None
    for i in range(1, app.Projects.Count + 1):
        p = app.Projects(i)
        if p.Name.lower() == source_project.lower() or source_project.lower() in p.Name.lower():
            src_proj = p
            break
    if src_proj is None:
        return json.dumps({"error": f"Source project '{source_project}' not found."})

    src_task = _find_task(src_proj, source_unique_id)
    if src_task is None:
        return json.dumps({"error": f"Source task UniqueID {source_unique_id} not found in '{src_proj.Name}'."})

    # Find target project and task
    tgt_proj = None
    for i in range(1, app.Projects.Count + 1):
        p = app.Projects(i)
        if p.Name.lower() == target_project.lower() or target_project.lower() in p.Name.lower():
            tgt_proj = p
            break
    if tgt_proj is None:
        return json.dumps({"error": f"Target project '{target_project}' not found."})

    tgt_task = _find_task(tgt_proj, target_unique_id)
    if tgt_task is None:
        return json.dumps({"error": f"Target task UniqueID {target_unique_id} not found in '{tgt_proj.Name}'."})

    # Set cross-project predecessor using "ProjectName\TaskID" format
    pred_str = f"{src_proj.Name}\\{src_task.ID}{link_type}"
    existing = (tgt_task.Predecessors or "").strip()
    if existing:
        tgt_task.Predecessors = existing + "," + pred_str
    else:
        tgt_task.Predecessors = pred_str

    try:
        app.FileSave()
    except Exception:
        pass

    return json.dumps({
        "status":  "linked",
        "source":  {"project": src_proj.Name, "task": src_task.Name, "unique_id": source_unique_id},
        "target":  {"project": tgt_proj.Name, "task": tgt_task.Name, "unique_id": target_unique_id},
        "link_type": link_type,
    }, indent=2)


@mcp.tool()
def undo_last(count: int = 1) -> str:
    """
    Safety net — undo last N operations in MS Project.

    Args:
        count: Number of undo steps (default 1, max 10).
    """
    if count < 1:
        count = 1
    if count > 10:
        count = 10

    app = get_app()
    for _ in range(count):
        try:
            app.EditUndo()
        except Exception:
            break

    return json.dumps({"status": "undone", "undo_count": count}, indent=2)


@mcp.tool()
def insert_subproject(file_path: str, after_unique_id: int = 0) -> str:
    """
    Insert an external .mpp file as a subproject.

    Args:
        file_path:       Full path to the .mpp file to insert (required).
        after_unique_id: Insert after this task's UniqueID (0 = insert at end).
    """
    import os
    if not os.path.exists(file_path):
        return json.dumps({"error": f"File not found: {file_path}"})

    app  = get_app()
    proj = get_proj(app)

    count_before = proj.Tasks.Count

    # Insert by adding a task and setting its SubProject property.
    # (SubprojectInsert is not reliably exposed via COM in all versions.)
    try:
        import os as _os
        basename = _os.path.splitext(_os.path.basename(file_path))[0]
        if after_unique_id > 0:
            t = _find_task(proj, after_unique_id)
            if t is None:
                return json.dumps({"error": f"Task UniqueID {after_unique_id} not found."})
            app.SelectRow(t.ID + 1, False)
            new_t = proj.Tasks.Add(basename)
        else:
            new_t = proj.Tasks.Add(basename)
        new_t.SubProject = file_path
    except Exception as e:
        return json.dumps({"error": f"Failed to insert subproject: {e}"})

    count_after = proj.Tasks.Count

    try:
        app.FileSave()
    except Exception:
        pass

    return json.dumps({
        "status":           "inserted",
        "file_path":        file_path,
        "inserted_after":   after_unique_id or "end",
        "task_count_before": count_before,
        "task_count_after":  count_after,
    }, indent=2)


@mcp.tool()
def snapshot_to_json(output_path: str, include_resources: bool = True) -> str:
    """
    Full project state dump for version control / diff.
    Exports all tasks and optionally resources to a JSON file.

    Args:
        output_path:       Full path for the output JSON file (required).
        include_resources: Include resource data (default True).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    tasks = []
    for t in proj.Tasks:
        if t is not None:
            tasks.append(task_to_dict(t, mpd))

    resources = []
    if include_resources:
        try:
            for r in proj.Resources:
                if r is not None:
                    resources.append({
                        "unique_id":  r.UniqueID,
                        "id":         r.ID,
                        "name":       r.Name,
                        "initials":   r.Initials,
                        "type":       r.Type,
                        "max_units":  r.MaxUnits,
                        "cost":       r.StandardRate,
                        "task_count": r.Assignments.Count,
                    })
        except Exception:
            pass

    project_meta = {
        "name":   proj.Name,
        "start":  _fmt_date(proj.ProjectStart),
        "finish": _fmt_date(proj.ProjectFinish),
    }
    try:
        project_meta["title"]   = proj.Title or ""
        project_meta["manager"] = proj.Manager or ""
    except Exception:
        pass

    snapshot = {
        "project":   project_meta,
        "tasks":     tasks,
        "resources": resources,
    }

    with open(output_path, "w", encoding="utf-8") as fp:
        json.dump(snapshot, fp, indent=2, default=str)

    return json.dumps({
        "status":    "exported",
        "path":      output_path,
        "tasks":     len(tasks),
        "resources": len(resources),
        "project":   project_meta["name"],
    }, indent=2)


@mcp.tool()
def health_check() -> str:
    """
    Lightweight connectivity test. Returns MS Project version, whether a project
    is open, and basic project info if available.
    """
    import win32com.client
    try:
        app = _com_retry(
            lambda: win32com.client.GetActiveObject("MSProject.Application"))
    except Exception:
        return json.dumps({"status": "disconnected", "error": "MS Project is not running."})

    # Everything past the attach goes through _com_retry as well. A busy
    # application accepts the attach and then rejects the first property read,
    # which is exactly how this tool failed with RPC_E_CALL_REJECTED on a
    # project that was open and reachable.
    try:
        result = {
            "status":  "connected",
            "version": str(_com_retry(lambda: app.Version)),
        }

        if _com_retry(lambda: app.Projects.Count) > 0:
            proj = app.ActiveProject
            result["project_open"] = True
            result["project_name"] = _com_retry(lambda: proj.Name)
            result["task_count"]   = _com_retry(lambda: proj.Tasks.Count)
        else:
            result["project_open"] = False
    except Exception as exc:
        # Reached the application, could not talk to it. Say which, rather than
        # letting a Windows-language COM code reach the client bare.
        return json.dumps({
            "status": "busy",
            "error": "MS Project answered but refused the call: %s" % str(exc)[:160],
            "hint": "It is usually recalculating or showing a dialog. Retried "
                    "for about 3 seconds before giving up; try again shortly, "
                    "and check the application for an open dialog box.",
        }, indent=2)

    return json.dumps(result, indent=2)


@mcp.tool()
def snapshot_diff(path_a: str, path_b: str) -> str:
    """
    Compare two JSON snapshot files (from snapshot_to_json) and return
    additions, deletions, and changes.

    Args:
        path_a: Path to the earlier snapshot JSON.
        path_b: Path to the later snapshot JSON.
    """
    import os

    for p in (path_a, path_b):
        if not os.path.exists(p):
            return json.dumps({"error": f"File not found: {p}"})

    with open(path_a, "r", encoding="utf-8") as f:
        snap_a = json.load(f)
    with open(path_b, "r", encoding="utf-8") as f:
        snap_b = json.load(f)

    tasks_a = {t["unique_id"]: t for t in snap_a.get("tasks", [])}
    tasks_b = {t["unique_id"]: t for t in snap_b.get("tasks", [])}

    ids_a = set(tasks_a.keys())
    ids_b = set(tasks_b.keys())

    added   = [tasks_b[uid] for uid in sorted(ids_b - ids_a)]
    deleted = [tasks_a[uid] for uid in sorted(ids_a - ids_b)]

    changed = []
    for uid in sorted(ids_a & ids_b):
        a, b = tasks_a[uid], tasks_b[uid]
        diffs = {}
        for key in set(list(a.keys()) + list(b.keys())):
            va, vb = a.get(key), b.get(key)
            if va != vb:
                diffs[key] = {"from": va, "to": vb}
        if diffs:
            changed.append({"unique_id": uid, "name": b.get("name", a.get("name")), "changes": diffs})

    return json.dumps({
        "added_count":   len(added),
        "deleted_count": len(deleted),
        "changed_count": len(changed),
        "added":         added,
        "deleted":       deleted,
        "changed":       changed,
    }, indent=2)
