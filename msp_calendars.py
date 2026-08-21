"""
Calendars: base calendars, exceptions and working hours.

Boundary: a tool belongs here when it manipulates a Calendar object, even when
the thing it is attached to is a Task or a Resource. That is why
set_task_calendar and set_resource_calendar live here rather than with tasks
or resources -- the subject is the calendar, the task is only the target.

This is the most physically scattered domain in the original single file,
drawn from seven separate blocks between lines 1517 and 4365.

set_project_calendar carries the project's only proj._oleobj_.InvokeTypes
call, the third of three cascading fallbacks, with its own deferred
'import pythoncom'. The whole try block travels together; slicing it would
strand the fallback.
"""

import json

from msp_core import (
    mcp,
    get_app,
    get_proj,
    _find_task,
    _fmt_date,
    _parse_date,
)


@mcp.tool()
def get_calendars() -> str:
    """List all base calendars in the project and which one is active."""
    app  = get_app()
    proj = get_proj(app)

    calendars = []
    try:
        for cal in proj.BaseCalendars:
            if cal is not None:
                calendars.append(str(cal.Name))
    except Exception:
        pass

    active = ""
    try:
        active = str(proj.Calendar)
    except Exception:
        pass

    return json.dumps({
        "active_calendar": active,
        "calendars":       calendars,
    }, indent=2)


@mcp.tool()
def set_calendar_exception(
    calendar_name: str,
    name: str,
    start: str,
    finish: str,
    working: bool = False,
) -> str:
    """
    Add a holiday or exception to a base calendar.

    Args:
        calendar_name: Name of the base calendar (e.g. 'Standard').
        name:          Exception name (e.g. 'National Day').
        start:         Start date as YYYY-MM-DD.
        finish:        End date as YYYY-MM-DD (same as start for single day).
        working:       True for a working exception, False for non-working/holiday (default).
    """
    app  = get_app()
    proj = get_proj(app)

    # Validate calendar exists
    valid_cals = []
    try:
        for cal in proj.BaseCalendars:
            if cal is not None:
                valid_cals.append(str(cal.Name))
    except Exception:
        pass

    # MS Project treats calendar names case-insensitively; match the same
    # way and canonicalise, so later comparisons against cal.Name hold.
    _match = next((c for c in valid_cals if c.lower() == calendar_name.lower()), None)
    if _match is None:
        return json.dumps({"error": f"Calendar '{calendar_name}' not found. Available: {valid_cals}"})
    calendar_name = _match

    try:
        # Use the Calendar.Exceptions collection for date-range exceptions
        cal = None
        for c in proj.BaseCalendars:
            if c is not None and str(c.Name) == calendar_name:
                cal = c
                break

        # pjCalendarExceptionDaily = 1
        start_dt  = _parse_date(start)
        finish_dt = _parse_date(finish)
        exc = cal.Exceptions.Add(1, start_dt, finish_dt, name)
    except Exception as e:
        return json.dumps({"error": f"Failed to set exception: {e}"})

    app.FileSave()
    return json.dumps({
        "status":   "created",
        "calendar": calendar_name,
        "exception": name,
        "start":    start,
        "finish":   finish,
        "working":  working,
    }, indent=2)


@mcp.tool()
def set_project_calendar(calendar_name: str) -> str:
    """
    Switch the active project's base calendar.

    Args:
        calendar_name: Name of the base calendar to use (e.g. '24 Hours', 'Standard').
    """
    app  = get_app()
    proj = get_proj(app)

    # Validate
    valid_cals = []
    try:
        for cal in proj.BaseCalendars:
            if cal is not None:
                valid_cals.append(str(cal.Name))
    except Exception:
        pass

    # MS Project treats calendar names case-insensitively; match the same
    # way and canonicalise, so later comparisons against cal.Name hold.
    _match = next((c for c in valid_cals if c.lower() == calendar_name.lower()), None)
    if _match is None:
        return json.dumps({"error": f"Calendar '{calendar_name}' not found. Available: {valid_cals}"})
    calendar_name = _match

    previous = ""
    try:
        previous = str(proj.Calendar)
    except Exception:
        pass

    # proj.Calendar is read-only in some COM bindings;
    # try multiple approaches to set it
    set_ok = False
    errors = []

    # Approach 1: direct property set
    try:
        proj.Calendar = calendar_name
        set_ok = True
    except Exception as e:
        errors.append(f"direct: {e}")

    # Approach 2: use the Calendar object from BaseCalendars
    if not set_ok:
        try:
            for cal in proj.BaseCalendars:
                if cal is not None and str(cal.Name) == calendar_name:
                    proj.Calendar = cal
                    set_ok = True
                    break
        except Exception as e:
            errors.append(f"object: {e}")

    # Approach 3: use _oleobj_ InvokeTypes to force property set
    if not set_ok:
        try:
            import pythoncom
            # Calendar property dispid — try to find via QueryInterface
            proj._oleobj_.InvokeTypes(
                0x30, 0, pythoncom.DISPATCH_PROPERTYPUT,
                (24, 0),  # VT_VOID return
                ((8, 1),),  # VT_BSTR input
                calendar_name
            )
            set_ok = True
        except Exception as e:
            errors.append(f"oleobj: {e}")

    if not set_ok:
        return json.dumps({"error": f"Could not set calendar. Tried: {errors}"})

    app.FileSave()

    return json.dumps({
        "status":   "updated",
        "calendar": calendar_name,
        "previous": previous,
    }, indent=2)


@mcp.tool()
def set_task_calendar(unique_id: int, calendar_name: str) -> str:
    """
    Set a task-level calendar override (e.g. 24/7 for commissioning phase).

    Args:
        unique_id:     Task UniqueID (required).
        calendar_name: Calendar name to apply, or empty string '' to clear.
    """
    app  = get_app()
    proj = get_proj(app)

    t = _find_task(proj, unique_id)
    if t is None:
        return json.dumps({"error": f"Task UniqueID {unique_id} not found."})

    # Validate calendar exists (if not clearing)
    if calendar_name:
        valid_cals = []
        try:
            for cal in proj.BaseCalendars:
                if cal is not None:
                    valid_cals.append(str(cal.Name))
        except Exception:
            pass
        # MS Project treats calendar names case-insensitively; match the same
        # way and canonicalise, so later comparisons against cal.Name hold.
        _match = next((c for c in valid_cals if c.lower() == calendar_name.lower()), None)
        if _match is None:
            return json.dumps({"error": f"Calendar '{calendar_name}' not found. Available: {valid_cals}"})
        calendar_name = _match

    previous = ""
    try:
        previous = str(t.Calendar) if t.Calendar else ""
    except Exception:
        pass

    try:
        if calendar_name:
            t.Calendar = calendar_name
        else:
            t.Calendar = ""
    except Exception as e:
        return json.dumps({"error": f"Failed to set task calendar: {e}"})

    return json.dumps({
        "status":    "updated",
        "unique_id": unique_id,
        "name":      t.Name,
        "calendar":  calendar_name or "(cleared)",
        "previous":  previous,
    }, indent=2)


@mcp.tool()
def create_calendar(name: str, copy_from: str = "Standard") -> str:
    """
    Create a new base calendar, optionally copying from an existing one.

    Args:
        name:      Name for the new calendar (required).
        copy_from: Existing calendar to copy from (default 'Standard').
    """
    app  = get_app()
    proj = get_proj(app)

    # Validate copy_from exists
    valid_cals = []
    try:
        for cal in proj.BaseCalendars:
            if cal is not None:
                valid_cals.append(str(cal.Name))
    except Exception:
        pass

    # MS Project treats calendar names case-insensitively; match the same
    # way and canonicalise, so later comparisons against cal.Name hold.
    _match = next((c for c in valid_cals if c.lower() == copy_from.lower()), None)
    if _match is None:
        return json.dumps({"error": f"Calendar '{copy_from}' not found. Available: {valid_cals}"})
    copy_from = _match

    if any(c.lower() == name.lower() for c in valid_cals):
        return json.dumps({"error": f"Calendar '{name}' already exists."})

    try:
        app.BaseCalendarCreate(Name=name, FromName=copy_from)
    except Exception:
        try:
            app.BaseCalendarCreate(name, copy_from)
        except Exception as e:
            return json.dumps({"error": f"Failed to create calendar: {e}"})

    # Re-read calendar list
    calendars = []
    try:
        for cal in proj.BaseCalendars:
            if cal is not None:
                calendars.append(str(cal.Name))
    except Exception:
        pass

    return json.dumps({
        "status":     "created",
        "name":       name,
        "copied_from": copy_from,
        "calendars":  calendars,
    }, indent=2)


@mcp.tool()
def list_calendar_exceptions(calendar_name: str = "") -> str:
    """
    List all exceptions (holidays, non-working days) defined on a calendar.
    If calendar_name is empty, uses the project calendar.

    Args:
        calendar_name: Name of the base calendar. Empty = project calendar.
    """
    app  = get_app()
    proj = get_proj(app)

    if not calendar_name:
        try:
            calendar_name = str(proj.Calendar)
        except Exception:
            calendar_name = "Standard"

    cal = None
    for c in proj.BaseCalendars:
        if c is not None and str(c.Name).lower() == str(calendar_name).lower():
            cal = c
            break
    if cal is None:
        return json.dumps({"error": f"Calendar '{calendar_name}' not found."})

    exceptions = []
    try:
        for exc in cal.Exceptions:
            exceptions.append({
                "name":   str(exc.Name),
                "start":  _fmt_date(exc.Start),
                "finish": _fmt_date(exc.Finish),
                "type":   exc.Type,
            })
    except Exception:
        pass  # Some calendars have no Exceptions collection

    return json.dumps({
        "calendar":   str(cal.Name),
        "count":      len(exceptions),
        "exceptions": exceptions,
    }, indent=2)


@mcp.tool()
def delete_calendar(calendar_name: str) -> str:
    """
    Delete a base calendar by name. Cannot delete the project calendar.

    Args:
        calendar_name: Name of the calendar to delete.
    """
    app  = get_app()
    proj = get_proj(app)

    try:
        proj_cal = str(proj.Calendar).lower()
    except Exception:
        proj_cal = "standard"
    if proj_cal == calendar_name.lower():
        return json.dumps({"error": "Cannot delete the active project calendar."})

    for c in proj.BaseCalendars:
        if c is not None and str(c.Name).lower() == calendar_name.lower():
            c.Delete()
            app.FileSave()
            return json.dumps({"status": "deleted", "calendar": calendar_name})

    return json.dumps({"error": f"Calendar '{calendar_name}' not found."})


@mcp.tool()
def delete_calendar_exception(calendar_name: str, exception_name: str) -> str:
    """
    Remove a specific exception (holiday/non-working day) from a calendar.

    Args:
        calendar_name:  Name of the base calendar.
        exception_name: Name of the exception to remove.
    """
    app  = get_app()
    proj = get_proj(app)

    cal = None
    for c in proj.BaseCalendars:
        if c is not None and str(c.Name).lower() == calendar_name.lower():
            cal = c
            break
    if cal is None:
        return json.dumps({"error": f"Calendar '{calendar_name}' not found."})

    try:
        for exc in cal.Exceptions:
            if str(exc.Name).lower() == exception_name.lower():
                exc.Delete()
                app.FileSave()
                return json.dumps({
                    "status":    "deleted",
                    "calendar":  calendar_name,
                    "exception": exception_name,
                })
    except Exception as e:
        return json.dumps({"error": f"Failed to access exceptions: {e}"})

    return json.dumps({"error": f"Exception '{exception_name}' not found in calendar '{calendar_name}'."})


@mcp.tool()
def set_resource_calendar(resource_name: str, calendar_name: str) -> str:
    """
    Assign a specific base calendar to a resource (e.g., part-time, different timezone).

    Args:
        resource_name: Name of the resource.
        calendar_name: Name of the base calendar to assign.
    """
    app  = get_app()
    proj = get_proj(app)

    # Verify calendar exists
    cal_found = False
    for c in proj.BaseCalendars:
        if c is not None and str(c.Name).lower() == calendar_name.lower():
            cal_found = True
            break
    if not cal_found:
        return json.dumps({"error": f"Calendar '{calendar_name}' not found."})

    for r in proj.Resources:
        if r is not None and r.Name and r.Name.lower() == resource_name.lower():
            r.BaseCalendar = calendar_name
            app.FileSave()
            return json.dumps({
                "status":   "updated",
                "resource": r.Name,
                "calendar": calendar_name,
            }, indent=2)

    return json.dumps({"error": f"Resource '{resource_name}' not found."})


@mcp.tool()
def set_working_hours(calendar_name: str, day: int, shifts_json: str) -> str:
    """
    Modify working hours for a specific day of the week in a calendar.

    Args:
        calendar_name: Name of the base calendar.
        day:           Day number (1=Sunday, 2=Monday, ..., 7=Saturday).
        shifts_json:   JSON array of shifts, e.g. [["08:00","12:00"],["13:00","17:00"]].
                       Empty array [] marks the day as non-working.
    """
    app  = get_app()
    proj = get_proj(app)

    if day < 1 or day > 7:
        return json.dumps({"error": "day must be 1 (Sunday) through 7 (Saturday)."})

    cal = None
    for c in proj.BaseCalendars:
        if c is not None and str(c.Name).lower() == calendar_name.lower():
            cal = c
            break
    if cal is None:
        return json.dumps({"error": f"Calendar '{calendar_name}' not found."})

    shifts = json.loads(shifts_json)

    wd = cal.WeekDays(day)

    if not shifts:
        # Mark as non-working
        wd.Working = False
        app.FileSave()
        return json.dumps({
            "status":   "updated",
            "calendar": cal.Name,
            "day":      day,
            "working":  False,
        }, indent=2)

    wd.Working = True

    # Set shifts using ShiftN Start/Finish properties (1-indexed)
    # Clear all 5 shifts first, then set the provided ones
    shift_attrs = [
        ("Shift1Start", "Shift1Finish"),
        ("Shift2Start", "Shift2Finish"),
        ("Shift3Start", "Shift3Finish"),
        ("Shift4Start", "Shift4Finish"),
        ("Shift5Start", "Shift5Finish"),
    ]

    for i, (s_attr, f_attr) in enumerate(shift_attrs):
        try:
            if i < len(shifts):
                setattr(wd, s_attr, shifts[i][0])
                setattr(wd, f_attr, shifts[i][1])
            else:
                # Clear unused shifts
                setattr(wd, s_attr, "")
                setattr(wd, f_attr, "")
        except Exception:
            pass

    app.FileSave()
    return json.dumps({
        "status":   "updated",
        "calendar": cal.Name,
        "day":      day,
        "working":  True,
        "shifts":   shifts[:5],
    }, indent=2)
