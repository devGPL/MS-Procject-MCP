"""
Planned versus actual: baselines, earned value, cost and variance.

Owns the comparison between what was planned and what happened -- saving and
clearing baselines, comparing them, and the reports derived from that gap
(BCWS/BCWP/ACWP/SPI/CPI, cost roll-up, schedule and cost variance).

Boundary: a tool belongs here when it REPORTS on planned-versus-actual. A tool
that edits the numbers themselves belongs to its own domain -- the rate-table
tools write to r.CostRateTables and live in msp_resources, even though they
decide what things cost.

compare_baselines defines _bl_attr as a nested helper. It travels with the
function; Gate 1 is what would catch it being left behind.
"""

import json

from msp_core import (
    responder,
    mcp,
    get_app,
    get_proj,
    _fmt_date,
    _get_mpd,
    _to_naive,
)


@mcp.tool()
def save_baseline(baseline_number: int = 0, all_tasks: bool = True) -> str:
    """
    Save a baseline for variance and earned-value tracking.
    
    baseline_number 0-10 (0 = the main Baseline). all_tasks False baselines only
    the current selection in the Microsoft Project window.
    """
    if baseline_number < 0 or baseline_number > 10:
        return json.dumps({"error": "baseline_number must be 0-10."})

    app  = get_app()
    proj = get_proj(app)

    task_count = sum(1 for t in proj.Tasks if t is not None)

    app.BaselineSave(All=all_tasks, Copy=baseline_number, Into=baseline_number)
    app.FileSave()

    return responder({
        "status":          "saved",
        "baseline_number": baseline_number,
        "all_tasks":       all_tasks,
        "tasks_baselined": task_count if all_tasks else "selected",
    })


@mcp.tool()
def clear_baseline(baseline_number: int = 0, all_tasks: bool = True) -> str:
    """
    Clear a previously saved baseline.

    Args:
        baseline_number: 0 to 10. Default 0.
        all_tasks:       True to clear for all tasks (default).
    """
    if baseline_number < 0 or baseline_number > 10:
        return json.dumps({"error": "baseline_number must be 0-10."})

    app = get_app()
    app.BaselineClear(All=all_tasks, From=baseline_number)
    app.FileSave()

    return responder({
        "status":          "cleared",
        "baseline_number": baseline_number,
    })


@mcp.tool()
def get_earned_value() -> str:
    """
    Return earned value metrics for all non-summary tasks.
    Requires a saved baseline and progress (% complete) to return meaningful data.
    Fields: BCWS (PV), BCWP (EV), ACWP (AC), SV, CV, plus SPI and CPI.
    """
    app  = get_app()
    proj = get_proj(app)

    tasks = []
    totals = {"bcws": 0, "bcwp": 0, "acwp": 0, "sv": 0, "cv": 0}

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue

        try:
            bcws = float(t.BCWS) if t.BCWS else 0
        except Exception:
            bcws = 0
        try:
            bcwp = float(t.BCWP) if t.BCWP else 0
        except Exception:
            bcwp = 0
        try:
            acwp = float(t.ACWP) if t.ACWP else 0
        except Exception:
            acwp = 0
        try:
            sv = float(t.SV) if t.SV else 0
        except Exception:
            sv = 0
        try:
            cv = float(t.CV) if t.CV else 0
        except Exception:
            cv = 0

        totals["bcws"] += bcws
        totals["bcwp"] += bcwp
        totals["acwp"] += acwp
        totals["sv"]   += sv
        totals["cv"]   += cv

        tasks.append({
            "unique_id": t.UniqueID,
            "name":      t.Name,
            "bcws":      bcws,
            "bcwp":      bcwp,
            "acwp":      acwp,
            "sv":        sv,
            "cv":        cv,
        })

    # Derived indices
    spi = round(totals["bcwp"] / totals["bcws"], 3) if totals["bcws"] else 0
    cpi = round(totals["bcwp"] / totals["acwp"], 3) if totals["acwp"] else 0

    if totals["bcws"] == 0 and totals["bcwp"] == 0:
        warning = "No earned value data. Ensure a baseline is saved and progress is entered."
    else:
        warning = None

    result = {
        "project_totals": {
            "bcws": totals["bcws"],
            "bcwp": totals["bcwp"],
            "acwp": totals["acwp"],
            "sv":   totals["sv"],
            "cv":   totals["cv"],
            "spi":  spi,
            "cpi":  cpi,
        },
        "tasks": tasks,
    }
    if warning:
        result["warning"] = warning

    return responder(result)


@mcp.tool()
def compare_baselines(baseline_a: int = 0, baseline_b: int = -1) -> str:
    """
    Variance report between two baselines, or baseline vs current schedule.

    Args:
        baseline_a: First baseline number (0-10). Default 0.
        baseline_b: Second baseline number (0-10), or -1 for current schedule (default).
    """
    if baseline_a < 0 or baseline_a > 10:
        return json.dumps({"error": "baseline_a must be 0-10."})
    if baseline_b < -1 or baseline_b > 10:
        return json.dumps({"error": "baseline_b must be -1 to 10 (-1 = current schedule)."})

    app  = get_app()
    proj = get_proj(app)

    def _bl_attr(n, suffix):
        if n == 0:
            return f"Baseline{suffix}"
        return f"Baseline{n}{suffix}"

    tasks = []
    tasks_with_variance = 0
    total_finish_variance = 0
    max_slippage = 0
    total_tasks = 0

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        total_tasks += 1

        try:
            a_start  = _to_naive(getattr(t, _bl_attr(baseline_a, "Start"), None))
            a_finish = _to_naive(getattr(t, _bl_attr(baseline_a, "Finish"), None))
        except Exception:
            a_start = a_finish = None

        if baseline_b == -1:
            b_start  = _to_naive(t.Start)
            b_finish = _to_naive(t.Finish)
        else:
            try:
                b_start  = _to_naive(getattr(t, _bl_attr(baseline_b, "Start"), None))
                b_finish = _to_naive(getattr(t, _bl_attr(baseline_b, "Finish"), None))
            except Exception:
                b_start = b_finish = None

        start_delta = finish_delta = None
        if a_start and b_start:
            try:
                start_delta = (b_start - a_start).days
            except Exception:
                pass
        if a_finish and b_finish:
            try:
                finish_delta = (b_finish - a_finish).days
            except Exception:
                pass

        if finish_delta is not None and finish_delta != 0:
            tasks_with_variance += 1
            total_finish_variance += finish_delta
            if finish_delta > max_slippage:
                max_slippage = finish_delta

        tasks.append({
            "unique_id":    t.UniqueID,
            "name":         t.Name,
            "a_start":      _fmt_date(a_start),
            "a_finish":     _fmt_date(a_finish),
            "b_start":      _fmt_date(b_start),
            "b_finish":     _fmt_date(b_finish),
            "start_delta":  start_delta,
            "finish_delta": finish_delta,
        })

    # Sort by finish variance desc (worst slippages first)
    tasks.sort(key=lambda x: x["finish_delta"] if x["finish_delta"] is not None else 0, reverse=True)

    return responder({
        "baseline_a": baseline_a,
        "baseline_b": baseline_b if baseline_b >= 0 else "current",
        "summary": {
            "total_tasks":          total_tasks,
            "tasks_with_variance":  tasks_with_variance,
            "avg_finish_variance":  round(total_finish_variance / tasks_with_variance, 1) if tasks_with_variance else 0,
            "max_slippage":         max_slippage,
        },
        "tasks": tasks,
    })


@mcp.tool()
def get_cost_summary() -> str:
    """
    Budget rollup with cost fields.
    Returns project cost totals, cost breakdown by resource, and tasks with cost data.
    Requires cost data to be entered on tasks or resources.
    """
    app  = get_app()
    proj = get_proj(app)

    totals = {"cost": 0, "actual_cost": 0, "remaining_cost": 0, "baseline_cost": 0}
    by_resource = {}
    tasks_with_cost = []

    for t in proj.Tasks:
        if t is None or t.Summary:
            continue

        cost = actual = remaining = baseline = 0
        try:
            cost = float(t.Cost) if t.Cost else 0
        except Exception:
            pass
        try:
            actual = float(t.ActualCost) if t.ActualCost else 0
        except Exception:
            pass
        try:
            remaining = float(t.RemainingCost) if t.RemainingCost else 0
        except Exception:
            pass
        try:
            baseline = float(t.BaselineCost) if t.BaselineCost else 0
        except Exception:
            pass

        totals["cost"] += cost
        totals["actual_cost"] += actual
        totals["remaining_cost"] += remaining
        totals["baseline_cost"] += baseline

        if cost > 0:
            tasks_with_cost.append({
                "unique_id":      t.UniqueID,
                "name":           t.Name,
                "cost":           cost,
                "actual_cost":    actual,
                "remaining_cost": remaining,
                "baseline_cost":  baseline,
            })

    # Cost by resource
    try:
        for r in proj.Resources:
            if r is not None:
                r_cost = 0
                r_actual = 0
                try:
                    r_cost = float(r.Cost) if r.Cost else 0
                except Exception:
                    pass
                try:
                    r_actual = float(r.ActualCost) if r.ActualCost else 0
                except Exception:
                    pass
                if r_cost > 0 or r_actual > 0:
                    by_resource[r.Name] = {"cost": r_cost, "actual_cost": r_actual}
    except Exception:
        pass

    totals["variance"] = totals["baseline_cost"] - totals["cost"]

    return responder({
        "project": proj.Name,
        "totals":  totals,
        "by_resource":     list({"name": k, **v} for k, v in by_resource.items()),
        "tasks_with_cost": tasks_with_cost,
    })


@mcp.tool()
def get_variance_report(baseline: int = 0) -> str:
    """
    Schedule and cost variance per task compared to a baseline.

    Args:
        baseline: Baseline number (0-10). Default 0.
    """
    app  = get_app()
    proj = get_proj(app)
    mpd  = _get_mpd(proj)

    # Baseline property name mapping
    BL_START  = f"Baseline{'' if baseline == 0 else baseline}Start"
    BL_FINISH = f"Baseline{'' if baseline == 0 else baseline}Finish"
    BL_COST   = f"Baseline{'' if baseline == 0 else baseline}Cost"

    tasks = []
    for t in proj.Tasks:
        if t is None or t.Summary:
            continue
        try:
            bl_start  = getattr(t, BL_START, None)
            bl_finish = getattr(t, BL_FINISH, None)
            bl_cost   = getattr(t, BL_COST, 0) or 0

            sv_days = round(t.StartVariance / mpd, 2) if t.StartVariance else 0
            fv_days = round(t.FinishVariance / mpd, 2) if t.FinishVariance else 0
            cv      = round((t.Cost or 0) - bl_cost, 2)

            tasks.append({
                "unique_id":        t.UniqueID,
                "name":             t.Name,
                "start":            _fmt_date(t.Start),
                "finish":           _fmt_date(t.Finish),
                "baseline_start":   _fmt_date(bl_start),
                "baseline_finish":  _fmt_date(bl_finish),
                "start_variance_days":  sv_days,
                "finish_variance_days": fv_days,
                "cost":             round(t.Cost or 0, 2),
                "baseline_cost":    round(bl_cost, 2),
                "cost_variance":    cv,
            })
        except Exception:
            continue

    # Filter to only tasks with actual variance
    with_variance = [t for t in tasks if t["start_variance_days"] != 0 or t["finish_variance_days"] != 0 or t["cost_variance"] != 0]

    return responder({
        "baseline":       baseline,
        "total_tasks":    len(tasks),
        "with_variance":  len(with_variance),
        "tasks":          with_variance if with_variance else tasks[:50],
    })
