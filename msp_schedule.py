"""
Schedule computation: critical path, slack, leveling and status rolls.

Owns what is CALCULATED over the network of tasks rather than stored on any
one of them -- critical path and its ordered sequence, float analysis,
validation, leveling, recalculation, and the status-date rituals that move
remaining work.

Boundary: a tool belongs here when its subject is the schedule as a whole. A
tool that sets one task's constraint or mode edits that task and lives with
tasks, even though the schedule shifts as a result.

This cut finally puts get_critical_path next to get_critical_path_sequence.
In the original file they sat 4,382 lines apart, though the second names the
first in its own docstring.
"""

import json
import datetime

import msp_fast

from msp_core import (
    mcp,
    responder,
    enxugar,
    recortar,
    get_app,
    get_proj,
    tarefa_completa,
    vista_com,
    VistaCOM,
    _dia,
    _dt_de_vista,
    _find_task,
    _fmt_date,
    _get_mpd,
    _parse_date,
    _to_naive,
)


# ---------------------------------------------------------------------------
# Dependency-network probe
#
# Every tool in this module computes over the network of links between tasks.
# A project whose detail tasks carry no predecessors has no network, and CPM is
# undefined on it -- but the tools still return numbers, and those numbers look
# like measurements. Slack comes back as zero for every task, which reads as
# "everything is critical" rather than "nothing is computed".
#
# Measured on a real 7,985-task schedule exported from a non-CPM planning tool:
# 7,583 detail tasks, none with a predecessor, and get_schedule_analysis
# reporting a project where every task has zero float.
# ---------------------------------------------------------------------------

# CPM needs tasks that are BOTH linked AND automatically scheduled. Once this
# many such tasks are seen the project can demonstrably be computed over, and
# further probing is wasted COM reads. A healthy project stops early; one that
# cannot be computed pays for the full scan, which is where the answer matters.
_PROBE_CEILING = 200


def _new_probe():
    return {"detail": 0, "linked": 0, "manual": 0, "auto_linked": 0,
            "probing": True}


def _probe_vista(probe, v, teto=True):
    """Count one detail task: is it linked, is it manually scheduled.

    Both matter, and checking only the first is what an earlier version of this
    probe got wrong. A project can be 92% linked and still yield no critical
    path, because Microsoft Project excludes manually scheduled tasks from the
    calculation -- the network exists and CPM does not traverse it. Measured on
    a real 8,429-task schedule: 2,758 of 3,000 detail tasks linked, all 3,000
    manual, zero marked critical.

    `teto` applies the sampling ceiling, and belongs to the COM backend only:
    the ceiling exists to stop paying two COM reads per task once the answer
    is settled. Reading two keys of a parsed dict is free, so that side counts
    everything and reports exact numbers instead of ">= 200".
    """
    probe["detail"] += 1
    if teto and not probe["probing"]:
        return
    try:
        tem_pred = bool(v["predecessors"])
    except Exception:
        tem_pred = False
    try:
        eh_manual = bool(v["manual"])
    except Exception:
        eh_manual = False

    if tem_pred:
        probe["linked"] += 1
    if eh_manual:
        probe["manual"] += 1
    if tem_pred and not eh_manual:
        probe["auto_linked"] += 1
        if teto and probe["auto_linked"] >= _PROBE_CEILING:
            probe["probing"] = False


def _probe_report(probe):
    """Diagnostics block for the response."""
    parcial = not probe["probing"]
    def conta(chave):
        return ">= %d (stopped counting)" % _PROBE_CEILING if parcial else probe[chave]
    return {
        "detail_tasks_scanned": probe["detail"],
        "with_predecessor": conta("linked"),
        "manually_scheduled": conta("manual"),
        "auto_and_linked": conta("auto_linked"),
    }


def _network_warning(probe, o_que):
    """Warning when `o_que` cannot be computed on this project, else None.

    Two distinct causes, and the remedy differs, so they get distinct
    messages: a project with no links needs dependencies; a linked project of
    manual tasks needs those tasks switched to automatic scheduling.
    """
    if probe["detail"] == 0 or probe["auto_linked"] > 0:
        return None

    comum = (
        "%s is computed over the dependency network, so on this project it has "
        "nothing to compute and the numbers above should not be read as a "
        "result. Every CPM-based tool here -- critical path, slack, what-if "
        "delay, schedule analysis -- is equally undefined on it." % o_que
    )

    if probe["linked"] == 0:
        return (
            "No dependency network found: none of the %d detail tasks has a "
            "predecessor. %s This happens when a schedule carries fixed dates "
            "instead of links, or was exported from a planning method that "
            "does not use CPM."
            % (probe["detail"], comum)
        )

    return (
        "The links exist but Microsoft Project cannot compute over them: %d of "
        "the %d detail tasks have a predecessor, and every task that has one is "
        "MANUALLY SCHEDULED. Manual tasks are excluded from critical path "
        "calculation, so the network is there and unused. %s To get a critical "
        "path, switch the linked tasks to automatic scheduling."
        % (probe["linked"], probe["detail"], comum)
    )


@mcp.tool()
def get_critical_path() -> str:
    """Return all tasks on the critical path (non-summary)."""
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    # Fast path first: parsing the saved file walks 8,000+ tasks in under a
    # second where COM needs half a minute. varredura() decides per call and
    # explains itself; anything it cannot verify sends us to COM.
    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)
    teto = lidas is None

    results = []
    probe = _new_probe()
    for v in fonte:
        if v["summary"]:
            continue
        _probe_vista(probe, v, teto)
        if v["critical"]:
            results.append(tarefa_completa(v))

    fatia, pagina = recortar([enxugar(t) for t in results])
    saida = {
        "count": len(results),
        "source": origem,
        "diagnostics": _probe_report(probe),
        "page": pagina,
        "tasks": fatia,
    }
    aviso = _network_warning(probe, "The critical path")
    if aviso:
        saida["warning"] = aviso
    return responder(saida)


@mcp.tool()
def get_schedule_analysis() -> str:
    """
    Return float/slack metrics and schedule health for all non-summary tasks.
    TotalSlack and FreeSlack are converted from minutes to working days.
    """
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)
    teto = lidas is None

    tasks = []
    zero_float = 0
    negative_float = 0
    total_slack_sum = 0
    count = 0

    probe = _new_probe()
    for v in fonte:
        if v["summary"]:
            continue
        count += 1
        _probe_vista(probe, v, teto)

        ts = v["total_slack_days"] or 0
        fs = v["free_slack_days"] or 0

        if ts == 0:
            zero_float += 1
        if ts < 0:
            negative_float += 1
        total_slack_sum += ts

        tasks.append({
            "unique_id":       v["unique_id"],
            "name":            v["name"],
            "total_slack_days": ts,
            "free_slack_days":  fs,
            "critical":        bool(v["critical"]),
            "start":           _dia(v["start"]),
            "finish":          _dia(v["finish"]),
        })

    fatia, pagina = recortar(tasks)
    saida = {
        "summary": {
            "total_tasks":     count,
            "zero_float":      zero_float,
            "negative_float":  negative_float,
            "avg_total_slack":  round(total_slack_sum / count, 2) if count else 0,
        },
        "source": origem,
        "diagnostics": _probe_report(probe),
        "page": pagina,
        "tasks": fatia,
    }
    aviso = _network_warning(probe, "Float analysis")
    if aviso:
        # Without links every task reports zero slack, so "zero_float" above
        # equals the task count and reads as a project entirely on the critical
        # path. It is the opposite: nothing was computed.
        saida["warning"] = aviso
    return responder(saida)


def _ids_de_predecessores(texto):
    """The row IDs named by a predecessor string, as strings.

    Parses the same shape the old nested loop did -- the leading digits of
    each entry, so "5FS+2d" yields "5" -- and accepts both separators,
    because the list separator is locale-dependent: an English install of
    Microsoft Project joins with "," and a Portuguese one with ";", and the
    parsed-file backend joins with ";" regardless. Splitting on only one of
    them finds no successors at all on the other, which reads as "every task
    is an orphan" rather than as a parse failure.
    """
    for parte in (texto or "").replace(";", ",").split(","):
        parte = parte.strip()
        num = ""
        for ch in parte:
            if ch.isdigit():
                num += ch
            else:
                break
        if num:
            yield num


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

    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)

    issues = {
        "orphan_tasks":       {"count": 0, "tasks": []},
        "no_resources":       {"count": 0, "tasks": []},
        "past_due_no_progress": {"count": 0, "tasks": []},
        "empty_summaries":    {"count": 0, "tasks": []},
        "missing_dates":      {"count": 0, "tasks": []},
        "negative_slack":     {"count": 0, "tasks": []},
    }

    # Materialised because "empty summary" is decided by looking at the NEXT
    # task, and because the orphan pass below needs two walks over the same
    # tasks. A COM view holds no field until asked, so a list of them costs
    # the enumeration, not 35 reads each.
    tarefas = list(fonte)
    total_tasks = len(tarefas)

    # Who is somebody else's predecessor -- built ONCE, in a single pass.
    #
    # This used to be a nested loop: for every detail task, walk the whole
    # task list re-reading Predecessors to see whether this task's row ID
    # appeared in it. That is O(n^2) COM property reads -- on the reference
    # 8,429-task schedule about 71 million of them at 0.240 ms each, which is
    # not slow but unusable. One pass building the set of referenced row IDs
    # answers the same question by membership.
    #
    # Values are the UniqueIDs that named that row, not just a count, so the
    # exclusion the nested loop made explicit ("skip the task itself") still
    # holds: a task that somehow lists its own row is not its own successor.
    referenciados = {}
    for v in tarefas:
        for rid in _ids_de_predecessores(v["predecessors"]):
            referenciados.setdefault(rid, set()).add(v["unique_id"])

    for i, v in enumerate(tarefas):
        tid = {"unique_id": v["unique_id"], "name": v["name"]}

        # Empty summaries: summary with no children
        if v["summary"]:
            has_child = False
            if i + 1 < len(tarefas):
                if tarefas[i + 1]["outline_level"] > v["outline_level"]:
                    has_child = True
            if not has_child:
                issues["empty_summaries"]["count"] += 1
                issues["empty_summaries"]["tasks"].append(tid)
            continue  # Skip non-leaf checks for summaries

        # Orphan tasks: no predecessors AND no successors
        preds = (v["predecessors"] or "").strip()
        has_successor = bool(referenciados.get(str(v["id"]), set())
                             - {v["unique_id"]})

        if not preds and not has_successor:
            issues["orphan_tasks"]["count"] += 1
            issues["orphan_tasks"]["tasks"].append(tid)

        # No resources (non-milestone)
        if not v["milestone"] and not (v["resource_names"] or "").strip():
            issues["no_resources"]["count"] += 1
            issues["no_resources"]["tasks"].append(tid)

        # Past due, zero progress
        fin = _dt_de_vista(v["finish"])
        if fin and fin < today and (v["percent_complete"] or 0) == 0:
            issues["past_due_no_progress"]["count"] += 1
            issues["past_due_no_progress"]["tasks"].append(tid)

        # Missing dates
        if not v["start"] or not v["finish"]:
            issues["missing_dates"]["count"] += 1
            issues["missing_dates"]["tasks"].append(tid)

        # Negative slack
        if (v["total_slack_days"] or 0) < 0:
            issues["negative_slack"]["count"] += 1
            issues["negative_slack"]["tasks"].append(tid)

    # Every category can name thousands of tasks. The counts stay exact --
    # the health score is computed from them -- and only the lists are cut,
    # each one saying so.
    for categoria in issues.values():
        fatia, pagina = recortar(categoria["tasks"])
        categoria["tasks"] = fatia
        if pagina.get("truncated"):
            categoria["page"] = pagina

    total_issues = sum(cat["count"] for cat in issues.values())
    health_score = max(0, round(100 - (total_issues / total_tasks * 100))) if total_tasks else 0

    return responder({
        "project":      proj.Name,
        "health_score": health_score,
        "source":       origem,
        "issues":       issues,
        "summary": {
            "total_tasks":  total_tasks,
            "total_issues": total_issues,
        },
    })


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

    return responder({
        "status":  "leveled",
        "project": proj.Name,
    })


@mcp.tool()
def find_available_slack(min_days: int = 5) -> str:
    """
    Find tasks with positive float — where can we absorb delay?

    Args:
        min_days: Minimum total slack in working days (default 5).
    """
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    lidas, origem = msp_fast.varredura(proj)
    fonte = lidas if lidas is not None else vista_com(proj, mpd)
    teto = lidas is None

    tasks = []
    probe = _new_probe()
    for v in fonte:
        if v["summary"]:
            continue
        _probe_vista(probe, v, teto)

        ts_days = v["total_slack_days"] or 0
        if ts_days < min_days:
            continue

        tasks.append({
            "unique_id":        v["unique_id"],
            "name":             v["name"],
            "total_slack_days": ts_days,
            "free_slack_days":  v["free_slack_days"] or 0,
            "start":            _dia(v["start"]),
            "finish":           _dia(v["finish"]),
            "resource_names":   v["resource_names"] or "",
        })

    tasks.sort(key=lambda x: x["total_slack_days"], reverse=True)

    fatia, pagina = recortar(tasks)
    saida = {
        "min_days":    min_days,
        "count":       len(tasks),
        "source":      origem,
        "diagnostics": _probe_report(probe),
        "page":        pagina,
        "tasks":       fatia,
    }
    aviso = _network_warning(probe, "Available slack")
    if aviso:
        saida["warning"] = aviso
    return responder(saida)


@mcp.tool()
def calculate_project() -> str:
    """
    Recalculate the active project schedule.
    Use after bulk manual changes to ensure dates, slack, and critical path are up to date.
    """
    app = get_app()
    app.CalculateProject()
    try:
        app.FileSave()
    except Exception:
        pass

    return json.dumps({"status": "calculated", "project": app.ActiveProject.Name})


@mcp.tool()
def update_project(complete_through: str, set_0_or_100: bool = False) -> str:
    """
    Mark progress complete through a date -- the weekly status ritual.
    
    complete_through: YYYY-MM-DD. set_0_or_100 restricts the result to 0% or 100%
    instead of partial progress. Changes percent complete on every task scheduled
    through that date.
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

    return responder({
        "status": "updated",
        "complete_through": complete_through,
        "set_0_or_100": set_0_or_100,
        "project": proj.Name,
    })


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

    return responder({
        "status": "rescheduled",
        "reschedule_from": str(dt)[:10],
        "project": proj.Name,
    })


@mcp.tool()
def get_critical_path_sequence() -> str:
    """
    The critical path as an ordered chain, each step linked to the next.
    
    Reports link type and lag between steps. Warns when the critical tasks are not
    linked to each other -- that is the absence of a critical path, not a path of
    one step.
    """
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    # Build adjacency graph of critical tasks only.
    #
    # The counters ride along on the scan that has to happen anyway. They cost
    # one extra property read per critical task and they are what turns a
    # degenerate answer into a diagnosable one: a schedule with no dependency
    # network still yields "critical" tasks -- the ones that happen to end on
    # the project finish date -- and reporting a one-step path without saying
    # so reads as data rather than as the absence of an answer.
    critical_tasks = {}
    scanned = 0
    critical_total = 0
    critical_summary = 0
    for t in proj.Tasks:
        if t is None:
            continue
        scanned += 1
        try:
            if not t.Critical:
                continue
        except Exception:
            continue
        critical_total += 1
        if t.Summary:
            critical_summary += 1
        else:
            critical_tasks[t.UniqueID] = t

    if not critical_tasks:
        return responder({
            "error": "No critical non-summary tasks found.",
            "tasks_scanned": scanned,
            "critical_total": critical_total,
            "critical_summary": critical_summary,
            "hint": "Every critical task in this project is a summary. That "
                    "usually means the detail tasks are manually scheduled or "
                    "unlinked, so MS Project computes no critical path.",
        })

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

    linked = sum(1 for uid in critical_tasks if forward.get(uid))

    fatia, pagina = recortar(sequence)
    resultado = {
        "project_start":       proj_start,
        "project_finish":      proj_finish,
        "critical_path_length": len(sequence),
        "total_duration_days":  total_duration,
        "total_critical_tasks": len(critical_tasks),
        "diagnostics": {
            "tasks_scanned":     scanned,
            "critical_total":    critical_total,
            "critical_summary":  critical_summary,
            "critical_detail":   len(critical_tasks),
            "linked_to_another_critical_task": linked,
        },
        "page":                pagina,
        "sequence":            fatia,
    }

    # A chain of one is not a critical path, it is the absence of one. Say so
    # rather than letting a 1.5-day "path" stand for a three-year project.
    if linked == 0 and len(sequence) <= 1:
        resultado["warning"] = (
            "No critical path could be traced. The critical tasks found are not "
            "linked to each other, so this is not a chain -- it is the single "
            "task that happens to finish last. A schedule whose detail tasks are "
            "manually scheduled or carry no predecessors has no dependency "
            "network for MS Project to compute a critical path over, and every "
            "CPM-based tool here (slack, what-if delay, schedule analysis) is "
            "equally undefined on it. Check whether the tasks came from a "
            "planning method that does not use CPM."
        )

    return responder(resultado)


@mcp.tool()
def get_critical_tasks_for_period(
    start_date: str,
    end_date: str,
    include_milestones: bool = True,
    include_non_critical_milestones: bool = False,
) -> str:
    """
    Critical tasks and milestones overlapping a date range.
    
    Dates YYYY-MM-DD, both required. include_milestones covers critical ones;
    include_non_critical_milestones adds the rest.
    """
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

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

    tarefas_pg, pagina = recortar(critical_tasks)
    marcos_pg, pagina_marcos = recortar(critical_milestones)
    outros_pg, pagina_outros = recortar(other_milestones)
    return responder({
        "period":              {"start": start_date, "end": end_date},
        "critical_task_count": len(critical_tasks),
        "critical_milestone_count": len(critical_milestones),
        "other_milestone_count": len(other_milestones),
        "page":                pagina,
        "critical_tasks":      tarefas_pg,
        "critical_milestones": marcos_pg,
        "critical_milestones_page": pagina_marcos,
        "other_milestones":    outros_pg,
        "other_milestones_page": pagina_outros,
    })


@mcp.tool()
def what_if_delay(
    unique_id: int,
    delay_days: int = 5,
) -> str:
    """
    Simulate delaying a task by N working days WITHOUT changing the project.
    
    Reports the project-end impact, tasks that would become critical, slack
    consumed and downstream tasks affected. An estimate from slack arithmetic,
    not a real reschedule.
    """
    app  = get_app()
    proj = get_proj(app)
    mpd = _get_mpd(proj)

    target = _find_task(proj, unique_id)
    if target is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    if target.Summary:
        return json.dumps({"error": "Cannot simulate delay on a summary task. Choose a work task or milestone."})

    delay_minutes = delay_days * mpd

    # Gather current state of all tasks
    task_data = {}
    probe = _new_probe()
    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        _probe_vista(probe, VistaCOM(t, mpd))
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

    saida = {
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
        "diagnostics":           _probe_report(probe),
        "downstream_tasks":      downstream_affected,
    }
    aviso = _network_warning(probe, "Delay propagation")
    if aviso:
        # The most dangerous of the four when it goes unflagged: it reports a
        # severity and a new finish date for a delay that cannot propagate,
        # because there are no links to propagate along. Downstream is always
        # empty and the impact always reads as none.
        saida["warning"] = aviso
    return responder(saida)
