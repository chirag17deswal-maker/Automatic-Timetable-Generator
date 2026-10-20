"""Turn the Excel data into a scheduling plan and check it before solving.

The solver can only say "no solution". The checks here find the common
reasons ahead of time and explain them in plain language.

A plan is a list of "entries". An entry is one thing to schedule:

* a normal subject of one class           -> class_keys has one class
* a common (joint) class                  -> class_keys has several classes,
                                             one teacher, one lesson for all
* one class of a common (parallel) group  -> one entry per class; the solver
                                             forces them into the same slot
                                             (each class has its own teacher)
"""

import pandas as pd

from services.schedule_config import (
    is_blank,
    longest_block,
    make_class_id,
    normalize_day,
    norm_key,
    norm_text,
    to_whole_number,
    valid_session_starts,
)

COMMON_MODES = ("joint", "parallel")


class TimetableInputError(ValueError):
    """The data can be read, but a timetable is impossible with it."""

    def __init__(self, problems):
        self.problems = list(problems)
        super().__init__("\n".join(self.problems))


def build_plan(classes, subjects, teachers, common=None):
    """Return (plan, teacher_names, problems).

    plan          list of entries (see module docstring)
    teacher_names teacher key -> name as written in the Teachers sheet
    problems      data errors found while building the plan
    """
    problems = []

    # Canonical class names come from the Classes sheet.
    class_labels = {}
    for _, row in classes.iterrows():
        label = make_class_id(row)
        class_labels.setdefault(norm_key(label), label)

    # Which teachers can teach which subject.
    teacher_names = {}
    subject_teachers = {}
    for _, row in teachers.iterrows():
        name = str(row["teacher_name"]).strip()
        teacher_key = norm_text(name)
        teacher_names.setdefault(teacher_key, name)

        if pd.isna(row["subjects"]):
            continue
        for subject in str(row["subjects"]).split(","):
            subject_key = norm_text(subject)
            if not subject_key:
                continue
            eligible = subject_teachers.setdefault(subject_key, [])
            if teacher_key not in eligible:
                eligible.append(teacher_key)

    plan = []
    used = set()   # (class, subject) already scheduled from the Subjects sheet

    # --- Normal subjects (one class each) ---------------------------------
    for _, row in subjects.iterrows():
        class_key = norm_key(row["class_id"])
        class_label = class_labels.get(class_key, str(row["class_id"]).strip())
        subject = str(row["subject_name"]).strip()
        where = f"{subject} ({class_label})"

        weekly = to_whole_number(row["weekly_periods"])
        length = to_whole_number(row["periods_per_session"])

        if not weekly or weekly < 1 or not length or length < 1:
            problems.append(
                f"{where}: weekly_periods and periods_per_session must be "
                "positive whole numbers."
            )
            continue

        if weekly % length != 0:
            problems.append(
                f"{where}: weekly_periods ({weekly}) must be divisible by "
                f"periods_per_session ({length})."
            )
            continue

        eligible = subject_teachers.get(norm_text(subject), [])
        if not eligible:
            problems.append(f"{where}: no teacher is assigned to this subject.")
            continue

        used.add((class_key, norm_text(subject)))
        plan.append({
            "class_keys": [class_key],
            "class_labels": [class_label],
            "common_with": [],
            "subject": subject,
            "weekly": weekly,
            "length": length,
            "sessions": weekly // length,
            "teachers": eligible,
            "fixed_day": None,
            "fixed_period": None,
            "link": None,
        })

    # --- Common classes (several classes at the same time) ----------------
    if common is not None:
        for group_number, (_, row) in enumerate(common.iterrows()):
            group = str(row["group_name"]).strip()
            subject = str(row["subject_name"]).strip()
            where = f"Common class '{group}'"
            bad = False

            keys, labels = [], []
            for token in str(row["class_ids"]).split(","):
                token = token.strip()
                if not token:
                    continue
                key = norm_key(token)
                if key not in class_labels:
                    problems.append(f"{where}: unknown class '{token}'.")
                    bad = True
                elif key not in keys:
                    keys.append(key)
                    labels.append(class_labels[key])

            if not bad and len(keys) < 2:
                problems.append(
                    f"{where}: list at least two classes in 'class_ids'."
                )
                bad = True

            weekly = to_whole_number(row["weekly_periods"])
            length = to_whole_number(row["periods_per_session"])
            if not weekly or weekly < 1 or not length or length < 1:
                problems.append(
                    f"{where}: weekly_periods and periods_per_session must "
                    "be positive whole numbers."
                )
                bad = True
            elif weekly % length != 0:
                problems.append(
                    f"{where}: weekly_periods ({weekly}) must be divisible "
                    f"by periods_per_session ({length})."
                )
                bad = True

            mode = row.get("mode")
            mode = "joint" if is_blank(mode) else str(mode).strip().casefold()
            if mode not in COMMON_MODES:
                problems.append(
                    f"{where}: 'mode' must be one of: "
                    + ", ".join(COMMON_MODES) + "."
                )
                bad = True

            fixed_day = row.get("fixed_day")
            if is_blank(fixed_day):
                fixed_day = None
            else:
                day = normalize_day(fixed_day)
                if day is None:
                    problems.append(
                        f"{where}: fixed_day '{fixed_day}' is not a day name."
                    )
                    bad = True
                fixed_day = day

            fixed_period = row.get("fixed_period")
            if is_blank(fixed_period):
                fixed_period = None
            else:
                number = to_whole_number(fixed_period)
                if number is None or number < 1:
                    problems.append(
                        f"{where}: fixed_period must be a whole number "
                        "(1 = first period)."
                    )
                    bad = True
                fixed_period = number

            eligible = subject_teachers.get(norm_text(subject), [])
            if not eligible:
                problems.append(
                    f"{where}: no teacher is assigned to subject '{subject}'."
                )
                bad = True

            for key, label in zip(keys, labels):
                if (key, norm_text(subject)) in used:
                    problems.append(
                        f"{where}: {label} also has '{subject}' in the "
                        "Subjects sheet. Remove it from one of the two sheets."
                    )
                    bad = True

            if bad:
                continue

            base = {
                "common_with": labels,
                "subject": subject,
                "weekly": weekly,
                "length": length,
                "sessions": weekly // length,
                "teachers": eligible,
                "fixed_day": fixed_day,
                "fixed_period": fixed_period,
            }

            if mode == "joint":
                # One lesson, one teacher, all classes together.
                plan.append({
                    **base,
                    "class_keys": keys,
                    "class_labels": labels,
                    "link": None,
                })
            else:
                # Same time for everyone, but each class has its own teacher.
                for key, label in zip(keys, labels):
                    plan.append({
                        **base,
                        "class_keys": [key],
                        "class_labels": [label],
                        "link": group_number,
                    })

    return plan, teacher_names, problems


def allowed_slots(entry, config):
    """All (day, first period) pairs where this entry may be scheduled."""
    starts = valid_session_starts(config["slots"], entry["length"])

    if entry["fixed_period"] is not None:
        starts = [s for s in starts if s == entry["fixed_period"]]

    days = config["days"]
    if entry["fixed_day"] is not None:
        days = [d for d in days if d == entry["fixed_day"]]

    return [(day, start) for day in days for start in starts]


def _where(entry):
    names = " + ".join(entry["class_labels"])
    return f"{entry['subject']} ({names})"


def check_feasibility(plan, teacher_names, config):
    """Return a list of plain-language reasons a timetable is impossible."""
    problems = []

    day_count = len(config["days"])
    slots = config["slots"]
    slots_per_day = len(slots)
    weekly_capacity = day_count * slots_per_day
    block = longest_block(slots)

    # 1. Every entry must physically fit into the week.
    for entry in plan:
        where = _where(entry)
        all_starts = valid_session_starts(slots, entry["length"])

        if not all_starts:
            problems.append(
                f"{where}: each session needs {entry['length']} periods in a "
                f"row, but the longest uninterrupted stretch in a day is "
                f"{block} period(s) (lunch splits the day)."
            )
            continue

        if entry["fixed_day"] is not None and entry["fixed_day"] not in config["days"]:
            problems.append(
                f"{where}: fixed_day '{entry['fixed_day']}' is not one of the "
                "working days in Settings."
            )
            continue

        if entry["fixed_period"] is not None and entry["fixed_period"] not in all_starts:
            problems.append(
                f"{where}: fixed_period {entry['fixed_period']} cannot be "
                f"used for a {entry['length']}-period session. Possible "
                "first periods: " + ", ".join(map(str, all_starts)) + "."
            )
            continue

        usable_days = {day for day, _ in allowed_slots(entry, config)}
        if entry["sessions"] > len(usable_days):
            problems.append(
                f"{where}: needs {entry['sessions']} sessions a week, but "
                f"only {len(usable_days)} day(s) are available to it (one "
                "session per subject per day)."
            )

    # 2. A class cannot have more periods than the week offers.
    class_load = {}
    class_name = {}
    for entry in plan:
        for key, label in zip(entry["class_keys"], entry["class_labels"]):
            class_load[key] = class_load.get(key, 0) + entry["weekly"]
            class_name[key] = label

    for key, load in class_load.items():
        if load > weekly_capacity:
            problems.append(
                f"{class_name[key]}: subjects add up to {load} periods a "
                f"week, but the week only has {weekly_capacity} "
                f"({day_count} days x {slots_per_day} periods)."
            )

    # 3. Two lessons pinned to the same slot for the same class.
    pinned = {}
    for entry in plan:
        if entry["fixed_day"] is None or entry["fixed_period"] is None:
            continue
        for key, label in zip(entry["class_keys"], entry["class_labels"]):
            for offset in range(entry["length"]):
                slot = (key, entry["fixed_day"], entry["fixed_period"] + offset)
                pinned.setdefault(slot, []).append((label, entry["subject"]))

    for (key, day, period), items in pinned.items():
        if len(items) > 1:
            subjects = ", ".join(sorted({s for _, s in items}))
            problems.append(
                f"{items[0][0]}: more than one lesson ({subjects}) is fixed "
                f"to {day}, period {period}."
            )

    # 4. Teacher workload. For every group of teachers that some entry can
    #    choose from, the entries that can ONLY be taught by that group must
    #    fit into the group's combined free time.
    teacher_sets = {frozenset(entry["teachers"]) for entry in plan}

    for group in teacher_sets:
        covered = [e for e in plan if frozenset(e["teachers"]) <= group]
        load = sum(e["weekly"] for e in covered)
        capacity = len(group) * weekly_capacity

        if load > capacity:
            names = ", ".join(sorted(teacher_names[t] for t in group))
            subjects = ", ".join(sorted({e["subject"] for e in covered}))
            problems.append(
                f"Teacher workload too high for {names}: {subjects} need "
                f"{load} periods a week in total, but only {capacity} slots "
                "are available to them."
            )

    # 5. Parallel common classes need a different teacher in every class.
    links = {}
    for entry in plan:
        if entry["link"] is not None:
            links.setdefault(entry["link"], []).append(entry)

    for entries in links.values():
        teachers_available = set()
        for entry in entries:
            teachers_available.update(entry["teachers"])
        if len(teachers_available) < len(entries):
            first = entries[0]
            problems.append(
                f"{first['subject']} ({' + '.join(first['common_with'])}): "
                f"runs at the same time in {len(entries)} classes, so it "
                f"needs {len(entries)} different teachers, but only "
                f"{len(teachers_available)} teacher(s) can teach it. Add "
                "teachers, or set mode to 'joint'."
            )

    return problems
