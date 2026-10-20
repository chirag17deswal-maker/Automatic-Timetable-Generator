import pandas as pd

from services.feasibility import build_plan
from services.schedule_config import (
    make_class_id,
    norm_key,
    norm_text,
    parse_settings,
)

REQUIRED_COLUMNS = {
    "Classes": [
        "course_name",
        "branch_name",
        "semester",
        "section"
    ],
    "Subjects": [
        "class_id",
        "subject_name",
        "weekly_periods",
        "periods_per_session",
        "subject_type"
    ],
    "Teachers": [
        "teacher_name",
        "subjects"
    ],
    "Settings": [
        "setting",
        "value"
    ]
}


# Optional sheet: classes shared by several courses/branches.
COMMON_REQUIRED_COLUMNS = [
    "group_name",
    "class_ids",
    "subject_name",
    "weekly_periods",
    "periods_per_session"
]


def validate_excel(file_path):
    errors = []

    # 1. Read the Excel workbook
    try:
        workbook = pd.ExcelFile(file_path, engine="openpyxl")
    except Exception as e:
        return [f"Unable to read Excel file: {e}"]

    # 2. Check required sheets
    for sheet in REQUIRED_COLUMNS:
        if sheet not in workbook.sheet_names:
            errors.append(f"Missing sheet: {sheet}")

    if errors:
        return errors

    # 3. Read sheets and check columns
    data = {}

    for sheet, required in REQUIRED_COLUMNS.items():
        df = pd.read_excel(workbook, sheet_name=sheet)
        df.columns = df.columns.astype(str).str.strip()

        # Ignore rows that are completely blank. The index is kept, so
        # "Excel row" numbers in the messages below stay correct.
        df = df.dropna(how="all")
        data[sheet] = df

        missing = set(required) - set(df.columns)

        if missing:
            errors.append(
                f"{sheet}: Missing columns: {', '.join(sorted(missing))}"
            )

    if errors:
        return errors

    classes = data["Classes"]
    subjects = data["Subjects"]
    teachers = data["Teachers"]
    settings = data["Settings"]

    # 4. Check for empty sheets
    for name, df in data.items():
        if df.empty:
            errors.append(f"{name}: No data found.")

    if errors:
        return errors

    # 5. Check required fields
    for name, df in data.items():
        required = REQUIRED_COLUMNS[name]

        for column in required:
            empty = df[column].isna() | (
                df[column].astype(str).str.strip() == ""
            )

            if empty.any():
                rows = (df.index[empty] + 2).tolist()
                errors.append(
                    f"{name}: Empty '{column}' in Excel row(s): {rows}"
                )

    # Stop further checks if required fields are missing
    if errors:
        return errors

    # 6. Validate class data
    semester = pd.to_numeric(classes["semester"], errors="coerce")
    invalid = semester.isna() | (semester < 1) | (semester % 1 != 0)

    if invalid.any():
        errors.append(
            "Classes: Semester must be a positive whole number."
        )

    duplicate_classes = classes.duplicated(
        subset=["course_name", "branch_name", "semester", "section"],
        keep=False
    )

    if duplicate_classes.any():
        errors.append("Classes: Duplicate class/section entries found.")

    # 7. Validate subject periods
    for column in ["weekly_periods", "periods_per_session"]:
        values = pd.to_numeric(subjects[column], errors="coerce")
        invalid = values.isna() | (values < 1) | (values % 1 != 0)

        if invalid.any():
            errors.append(
                f"Subjects: '{column}' must contain positive whole numbers."
            )

    # Each weekly requirement must fit complete sessions
    weekly = pd.to_numeric(subjects["weekly_periods"], errors="coerce")
    session = pd.to_numeric(
        subjects["periods_per_session"], errors="coerce"
    )

    valid = weekly.notna() & session.notna() & (session > 0)
    if ((weekly[valid] % session[valid]) != 0).any():
        errors.append(
            "Subjects: weekly_periods must be divisible by "
            "periods_per_session."
        )

    # Check that each subject refers to an existing class
    known_classes = {norm_key(make_class_id(row))
                     for _, row in classes.iterrows()}

    unknown_classes = []
    for value in subjects["class_id"].astype(str).str.strip():
        if norm_key(value) not in known_classes and value not in unknown_classes:
            unknown_classes.append(value)

    if unknown_classes:
        errors.append(
            "Subjects: Unknown class_id value(s): "
            + ", ".join(unknown_classes)
            + ". Each class_id must match a row in the Classes sheet, "
            "written as '<course> <branch> <semester><section>' "
            "(for example 'B.Tech CSE 3A')."
        )

    # 8. Validate teacher assignments
    teacher_keys = teachers["teacher_name"].map(norm_text)
    duplicate_teachers = teacher_keys[teacher_keys.duplicated(keep=False)]

    if not duplicate_teachers.empty:
        names = sorted(set(
            teachers.loc[duplicate_teachers.index, "teacher_name"]
            .astype(str).str.strip()
        ))
        errors.append(
            "Teachers: Duplicate teacher name(s): " + ", ".join(names)
            + ". List each teacher once and put all their subjects "
            "in the 'subjects' cell, separated by commas."
        )

    assigned_subjects = set()
    for value in teachers["subjects"].dropna():
        for subject in str(value).split(","):
            subject = norm_text(subject)
            if subject:
                assigned_subjects.add(subject)

    missing_subjects = []
    for name in subjects["subject_name"].astype(str).str.strip():
        if norm_text(name) not in assigned_subjects and name not in missing_subjects:
            missing_subjects.append(name)

    if missing_subjects:
        errors.append(
            "Subjects without an assigned teacher: "
            + ", ".join(missing_subjects)
        )

    # 9. Validate settings (days, times, durations)
    _, setting_errors = parse_settings(settings)
    errors.extend(setting_errors)

    # 10. Optional Common_Classes sheet
    common = None
    if "Common_Classes" in workbook.sheet_names:
        common = pd.read_excel(workbook, sheet_name="Common_Classes")
        common.columns = common.columns.astype(str).str.strip()
        common = common.dropna(how="all")

        missing = set(COMMON_REQUIRED_COLUMNS) - set(common.columns)
        if missing:
            errors.append(
                "Common_Classes: Missing columns: "
                + ", ".join(sorted(missing))
            )
        elif not common.empty:
            for column in COMMON_REQUIRED_COLUMNS:
                empty = common[column].isna() | (
                    common[column].astype(str).str.strip() == ""
                )
                if empty.any():
                    rows = (common.index[empty] + 2).tolist()
                    errors.append(
                        f"Common_Classes: Empty '{column}' in Excel "
                        f"row(s): {rows}"
                    )
        if common.empty:
            common = None

    # 11. Cross-checks that need every sheet to be readable
    #     (unknown classes in common classes, duplicates, bad numbers ...)
    if not errors:
        _, _, plan_problems = build_plan(classes, subjects, teachers, common)
        errors.extend(plan_problems)

    return errors
