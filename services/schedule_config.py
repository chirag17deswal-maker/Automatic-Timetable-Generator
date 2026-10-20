"""Shared helpers used by the validator, the optimizer and the web app.

Everything that turns the Excel "Settings" sheet into a working schedule
(days, period slots, lunch break) lives here, so the whole project agrees on
what a "period" is.
"""

import math
import re
from datetime import datetime, time

import pandas as pd

ALL_DAYS = [
    "Monday", "Tuesday", "Wednesday", "Thursday",
    "Friday", "Saturday", "Sunday",
]

# Accept "Monday", "monday", "Mon", "MON" ...
_DAY_LOOKUP = {}
for _day in ALL_DAYS:
    _DAY_LOOKUP[_day.casefold()] = _day
    _DAY_LOOKUP[_day[:3].casefold()] = _day

# off       - do not care about free periods
# preferred - avoid free periods where possible (default)
# strict    - lessons must start at period 1 and have no free period between
COMPACT_MODES = ("off", "preferred", "strict")

REQUIRED_SETTINGS = [
    "working_days",
    "college_start_time",
    "college_end_time",
    "period_duration_minutes",
    "lunch_start_time",
    "lunch_duration_minutes",
]


# ---------------------------------------------------------------------------
# Small text / number helpers
# ---------------------------------------------------------------------------

def norm_text(value):
    """Case-insensitive text with single spaces (for names)."""
    return " ".join(str(value).split()).casefold()


def norm_key(value):
    """Case-insensitive text with ALL whitespace removed (for class IDs)."""
    return "".join(str(value).split()).casefold()


def format_semester(value):
    """3 and 3.0 and '3' all become '3' (Excel often stores numbers as floats)."""
    try:
        number = float(value)
        if math.isfinite(number) and number == int(number):
            return str(int(number))
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def make_class_id(row):
    """Build the class ID, e.g. 'B.Tech CSE 3A', from a Classes-sheet row."""
    return (
        f"{str(row['course_name']).strip()} "
        f"{str(row['branch_name']).strip()} "
        f"{format_semester(row['semester'])}"
        f"{str(row['section']).strip()}"
    )


def to_whole_number(value):
    """Return an int for 50, 50.0 or '50'; None for anything else."""
    if isinstance(value, bool):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(number) or number != int(number):
        return None
    return int(number)


def is_blank(value):
    """True for empty cells (None, NaN, or only spaces)."""
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() == ""


_is_blank = is_blank


def normalize_day(value):
    """'mon', 'MONDAY' -> 'Monday'. Returns None if it is not a day."""
    return _DAY_LOOKUP.get(str(value).strip().casefold())


# ---------------------------------------------------------------------------
# Time handling
# ---------------------------------------------------------------------------

_TIME_PATTERN = re.compile(
    r"^(\d{1,2}):(\d{2})(?::\d{2})?\s*([ap]m)?$", re.IGNORECASE
)


def parse_time(value):
    """Return minutes after midnight, or None if the value is not a time.

    Handles text ("09:15", "9:15 AM"), Excel time cells (datetime.time /
    datetime) and Excel's raw day-fraction numbers (0.385 -> 09:15).
    """
    if value is None or _is_blank(value):
        return None

    if isinstance(value, (datetime, time)):
        return value.hour * 60 + value.minute

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 0 <= value < 1:
            return int(round(value * 24 * 60)) % (24 * 60)
        return None

    match = _TIME_PATTERN.match(str(value).strip())
    if not match:
        return None

    hours, minutes, meridian = int(match.group(1)), int(match.group(2)), match.group(3)

    if meridian:
        if not 1 <= hours <= 12:
            return None
        hours = hours % 12 + (12 if meridian.lower() == "pm" else 0)

    if hours > 23 or minutes > 59:
        return None
    return hours * 60 + minutes


def format_minutes(total_minutes):
    return f"{total_minutes // 60:02d}:{total_minutes % 60:02d}"


# ---------------------------------------------------------------------------
# Period slots
# ---------------------------------------------------------------------------

def create_period_slots(start, end, lunch_start, lunch_duration, period_duration):
    """Create the teaching periods of one day (all arguments in minutes).

    A period that would overlap lunch is skipped and the next one starts when
    lunch ends. Each slot keeps its exact start/end so we can later tell
    whether two neighbouring periods are really back-to-back.
    """
    lunch_end = lunch_start + lunch_duration

    slots = []
    current = start
    period_number = 1

    while current + period_duration <= end:
        next_time = current + period_duration

        if lunch_duration > 0 and current < lunch_end and next_time > lunch_start:
            current = lunch_end
            continue

        slots.append({
            "period": period_number,
            "start": format_minutes(current),
            "end": format_minutes(next_time),
            "start_min": current,
            "end_min": next_time,
        })

        period_number += 1
        current = next_time

    return slots


def valid_session_starts(slots, length):
    """Periods at which a session of `length` periods may start.

    A session needs `length` periods in a row with no break (lunch) in
    between, so a 2-period lab can never start right before lunch.
    """
    starts = []
    for i in range(len(slots) - length + 1):
        window = slots[i:i + length]
        back_to_back = all(
            window[j]["end_min"] == window[j + 1]["start_min"]
            for j in range(length - 1)
        )
        if back_to_back:
            starts.append(window[0]["period"])
    return starts


def longest_block(slots):
    """Longest run of back-to-back periods in a day."""
    best = run = 0
    for i, slot in enumerate(slots):
        if i > 0 and slots[i - 1]["end_min"] == slot["start_min"]:
            run += 1
        else:
            run = 1
        best = max(best, run)
    return best


# ---------------------------------------------------------------------------
# Settings sheet -> schedule configuration
# ---------------------------------------------------------------------------

def parse_settings(settings_df):
    """Turn the Settings sheet into a schedule configuration.

    Returns (config, errors). `config` is None when there are errors.
    """
    errors = []

    raw = {}
    for _, row in settings_df.iterrows():
        raw[str(row["setting"]).strip().lower()] = row["value"]

    for key in REQUIRED_SETTINGS:
        if key not in raw or _is_blank(raw[key]):
            errors.append(f"Settings: Missing setting '{key}'.")

    if errors:
        return None, errors

    # Working days
    days = []
    unknown = []
    for token in re.split(r"[,;]", str(raw["working_days"])):
        token = token.strip()
        if not token:
            continue
        day = _DAY_LOOKUP.get(token.casefold())
        if day is None:
            unknown.append(token)
        elif day not in days:
            days.append(day)

    if unknown:
        errors.append(
            "Settings: Unknown working day(s): " + ", ".join(unknown)
            + ". Use names like Monday, Tuesday ..."
        )
    elif not days:
        errors.append("Settings: working_days must list at least one day.")

    # Times
    times = {}
    for key in ("college_start_time", "college_end_time", "lunch_start_time"):
        minutes = parse_time(raw[key])
        if minutes is None:
            errors.append(
                f"Settings: '{key}' must be a time such as 09:15 "
                f"(got '{raw[key]}')."
            )
        else:
            times[key] = minutes

    # Durations
    period_duration = to_whole_number(raw["period_duration_minutes"])
    if period_duration is None or period_duration <= 0:
        errors.append(
            "Settings: 'period_duration_minutes' must be a positive whole number."
        )

    lunch_duration = to_whole_number(raw["lunch_duration_minutes"])
    if lunch_duration is None or lunch_duration < 0:
        errors.append(
            "Settings: 'lunch_duration_minutes' must be a whole number "
            "(0 or more)."
        )

    # Optional: how strongly lessons are packed towards the first period.
    compact = "preferred"
    if "compact_timetable" in raw and not _is_blank(raw["compact_timetable"]):
        compact = str(raw["compact_timetable"]).strip().casefold()
        if compact not in COMPACT_MODES:
            errors.append(
                "Settings: 'compact_timetable' must be one of: "
                + ", ".join(COMPACT_MODES) + "."
            )

    if errors:
        return None, errors

    start = times["college_start_time"]
    end = times["college_end_time"]
    lunch_start = times["lunch_start_time"]

    if end <= start:
        return None, [
            "Settings: 'college_end_time' must be later than "
            "'college_start_time'."
        ]

    slots = create_period_slots(
        start, end, lunch_start, lunch_duration, period_duration
    )

    if not slots:
        return None, [
            "Settings: no teaching period fits between the college start and "
            "end time. Check the period duration and lunch settings."
        ]

    # Rows for display: periods plus a lunch row where lunch happens.
    rows = []
    lunch_shown = False
    lunch_end = lunch_start + lunch_duration

    for slot in slots:
        if (
            not lunch_shown
            and lunch_duration > 0
            and start <= lunch_start
            and slot["start_min"] >= lunch_end
        ):
            rows.append({
                "type": "break",
                "label": "Lunch",
                "start": format_minutes(lunch_start),
                "end": format_minutes(lunch_end),
            })
            lunch_shown = True
        rows.append({"type": "period", **slot})

    config = {
        "days": days,
        "slots": slots,
        "periods": [slot["period"] for slot in slots],
        "rows": rows,
        "compact": compact,
    }
    return config, []
