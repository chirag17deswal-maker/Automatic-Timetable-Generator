import os
from collections import defaultdict

import pandas as pd
from ortools.sat.python import cp_model

from services.feasibility import (
    TimetableInputError,
    allowed_slots,
    build_plan,
    check_feasibility,
)
from services.schedule_config import parse_settings


def load_timetable_data(file_path):
    """Read the input Excel workbook."""
    try:
        workbook = pd.ExcelFile(file_path, engine="openpyxl")

        sheets = {}
        for name in ("Classes", "Subjects", "Teachers", "Settings"):
            df = pd.read_excel(workbook, sheet_name=name)
            df.columns = df.columns.astype(str).str.strip()
            # Ignore completely blank rows left behind in Excel.
            sheets[name.lower()] = df.dropna(how="all")

        # Optional sheet: classes that several courses/branches share.
        sheets["common"] = None
        if "Common_Classes" in workbook.sheet_names:
            df = pd.read_excel(workbook, sheet_name="Common_Classes")
            df.columns = df.columns.astype(str).str.strip()
            df = df.dropna(how="all")
            if not df.empty:
                sheets["common"] = df

        return sheets

    except Exception as e:
        raise ValueError(f"Unable to load timetable data: {e}")


def get_schedule_config(file_path):
    """Days, period slots and lunch row taken from the Settings sheet."""
    data = load_timetable_data(file_path)
    config, errors = parse_settings(data["settings"])
    if errors:
        raise TimetableInputError(errors)
    return config


def generate_timetable(file_path, time_limit_seconds=None):
    """Build a timetable from the Excel file.

    Returns a list of scheduled sessions (one per class), or None if the
    solver could not find a timetable. Raises TimetableInputError when the
    data itself makes a timetable impossible (the message explains why).
    """
    if time_limit_seconds is None:
        try:
            time_limit_seconds = int(os.environ.get("SOLVER_TIME_LIMIT", "8"))
        except (ValueError, TypeError):
            time_limit_seconds = 8
    data = load_timetable_data(file_path)

    config, errors = parse_settings(data["settings"])
    if errors:
        raise TimetableInputError(errors)

    plan, teacher_names, problems = build_plan(
        data["classes"], data["subjects"], data["teachers"], data["common"]
    )
    if not problems:
        problems = check_feasibility(plan, teacher_names, config)
    if problems:
        raise TimetableInputError(problems)

    days = config["days"]
    slots = config["slots"]
    periods = config["periods"]
    compact = config["compact"]

    model = cp_model.CpModel()

    # One Boolean per possible lesson: (entry, day, first period, teacher)
    lessons = {}
    # One Boolean per (entry, teacher): "this teacher takes the entry"
    assigned = {}

    # Lookup tables so each constraint only touches the lessons it needs.
    class_period = defaultdict(list)    # (class, day, period) -> lessons
    teacher_period = defaultdict(list)  # (teacher, day, period) -> lessons
    class_day = defaultdict(list)       # (class, day) -> [(lesson, length)]
    slot_vars = defaultdict(list)       # (entry, day, start) -> lessons

    for i, entry in enumerate(plan):
        length = entry["length"]

        # Exactly one teacher teaches this entry all week.
        for t, teacher in enumerate(entry["teachers"]):
            assigned[(i, teacher)] = model.NewBoolVar(f"assigned_{i}_{t}")
        model.Add(
            sum(assigned[(i, teacher)] for teacher in entry["teachers"]) == 1
        )

        entry_lessons = []
        per_day = defaultdict(list)

        for day, start in allowed_slots(entry, config):
            for t, teacher in enumerate(entry["teachers"]):
                var = model.NewBoolVar(f"lesson_{i}_{day}_{start}_{t}")
                lessons[(i, day, start, teacher)] = var

                # A lesson may only use the teacher chosen for this entry.
                model.Add(var <= assigned[(i, teacher)])

                entry_lessons.append(var)
                per_day[day].append(var)
                slot_vars[(i, day, start)].append(var)

                # A common (joint) lesson occupies every one of its classes.
                for class_key in entry["class_keys"]:
                    class_day[(class_key, day)].append((var, length))
                    for period in range(start, start + length):
                        class_period[(class_key, day, period)].append(var)

                # ...but the teacher is only busy once.
                for period in range(start, start + length):
                    teacher_period[(teacher, day, period)].append(var)

        # Schedule the required number of sessions...
        model.Add(sum(entry_lessons) == entry["sessions"])

        # ...with at most one session per day.
        for day_vars in per_day.values():
            model.Add(sum(day_vars) <= 1)

    # Parallel common classes: every class of the group meets at the same
    # time (each class with its own teacher).
    linked = defaultdict(list)
    for i, entry in enumerate(plan):
        if entry["link"] is not None:
            linked[entry["link"]].append(i)

    for members in linked.values():
        first = members[0]
        for other in members[1:]:
            for day, start in allowed_slots(plan[first], config):
                model.Add(
                    sum(slot_vars[(first, day, start)])
                    == sum(slot_vars[(other, day, start)])
                )

    # A class can only be in one lesson at a time.
    for occupying in class_period.values():
        if len(occupying) > 1:
            model.Add(sum(occupying) <= 1)

    # A teacher can only teach one lesson at a time.
    for occupying in teacher_period.values():
        if len(occupying) > 1:
            model.Add(sum(occupying) <= 1)

    # Spread each class's timetable across the week and pack its days.
    class_keys = list(dict.fromkeys(
        key for entry in plan for key in entry["class_keys"]
    ))

    active_days = []
    max_daily_loads = []
    day_ends = []

    for class_key in class_keys:
        daily_loads = []

        for day in days:
            daily = class_day.get((class_key, day))
            if not daily:
                continue

            active = model.NewBoolVar(f"active_{class_key}_{day}")
            active_days.append(active)

            day_total = sum(var for var, _ in daily)
            model.Add(day_total >= active)
            model.Add(day_total <= len(daily) * active)

            daily_loads.append(sum(length * var for var, length in daily))

            if compact != "off":
                # day_end = number of the last period used that day, so
                # minimising it removes free periods before lessons.
                day_end = model.NewIntVar(
                    0, len(slots), f"day_end_{class_key}_{day}"
                )
                for period in periods:
                    occupying = class_period.get((class_key, day, period))
                    if occupying:
                        model.Add(day_end >= period * sum(occupying))
                day_ends.append(day_end)

            if compact == "strict":
                # No free period before a lesson: a period can only be used
                # if the period before it is used.
                for period in periods[:-1]:
                    following = class_period.get((class_key, day, period + 1))
                    if not following:
                        continue
                    current = class_period.get((class_key, day, period))
                    model.Add(
                        sum(following) <= (sum(current) if current else 0)
                    )

        if daily_loads:
            class_max = model.NewIntVar(
                0, len(slots), f"class_max_{class_key}"
            )
            for load in daily_loads:
                model.Add(class_max >= load)
            max_daily_loads.append(class_max)

    # Priorities (highest first): use more days, then pack each day from the
    # first period, then keep the busiest day as light as possible.
    load_bound = len(class_keys) * len(slots)
    compact_weight = load_bound + 1 if day_ends else 0
    day_end_bound = len(class_keys) * len(days) * len(slots)
    active_weight = compact_weight * day_end_bound + load_bound + 1

    model.Maximize(
        active_weight * sum(active_days)
        - compact_weight * sum(day_ends)
        - sum(max_daily_loads)
    )

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_seconds

    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None

    # Collect the generated timetable (one row per class).
    day_order = {day: n for n, day in enumerate(days)}
    result = []

    for (i, day, start, teacher), var in lessons.items():
        if solver.Value(var) != 1:
            continue

        entry = plan[i]
        last_period = start + entry["length"] - 1

        for label in entry["class_labels"]:
            result.append({
                "class_id": label,
                "subject": entry["subject"],
                "teacher": teacher_names[teacher],
                "day": day,
                "period": start,
                "periods_per_session": entry["length"],
                "start_time": slots[start - 1]["start"],
                "end_time": slots[last_period - 1]["end"],
                "shared_with": [
                    other for other in entry["common_with"] if other != label
                ],
            })

    result.sort(key=lambda r: (r["class_id"], day_order[r["day"]], r["period"]))
    return result


if __name__ == "__main__":
    from pathlib import Path

    file_path = (
        Path(__file__).resolve().parent.parent
        / "sample_data"
        / "timetable_input_sample.xlsx"
    )

    print("Excel file exists:", file_path.exists())
    print("Excel file path:", file_path)

    if file_path.exists():
        try:
            timetable = generate_timetable(file_path)
        except TimetableInputError as error:
            print("Cannot build a timetable:")
            for problem in error.problems:
                print(" -", problem)
        else:
            if timetable is None:
                print("No feasible timetable could be generated.")
            else:
                print("TIMETABLE GENERATED SUCCESSFULLY")
                for lesson in timetable:
                    print(lesson)
