import io
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Styles definition
FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="1E293B")
FONT_SUBTITLE = Font(name="Calibri", size=10, italic=True, color="64748B")
FONT_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_PERIOD = Font(name="Calibri", size=10, bold=True, color="1E293B")
FONT_TIME = Font(name="Calibri", size=9, color="64748B")
FONT_SESSION_SUBJ = Font(name="Calibri", size=10, bold=True, color="1E1B4B")
FONT_SESSION_INFO = Font(name="Calibri", size=9, color="374151")
FONT_BREAK = Font(name="Calibri", size=10, bold=True, color="92400E")
FONT_FREE = Font(name="Calibri", size=9, italic=True, color="94A3B8")

FILL_HEADER = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
FILL_PERIOD_COL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
FILL_SESSION = PatternFill(start_color="EEF2FF", end_color="EEF2FF", fill_type="solid")
FILL_BREAK = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
FILL_FREE = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")
FILL_SUMMARY_HEADER = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")

THIN_BORDER_SIDE = Side(border_style="thin", color="CBD5E1")
BORDER_CELL = Border(
    left=THIN_BORDER_SIDE,
    right=THIN_BORDER_SIDE,
    top=THIN_BORDER_SIDE,
    bottom=THIN_BORDER_SIDE
)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def sanitize_sheet_name(name, existing_names):
    """Sanitize sheet name for Excel (max 31 chars, remove invalid chars, ensure unique)."""
    clean = re.sub(r'[\\/*?:\[\]]', '', str(name)).strip()
    if not clean:
        clean = "Sheet"
    clean = clean[:28]
    candidate = clean
    counter = 1
    while candidate.casefold() in [e.casefold() for e in existing_names]:
        candidate = f"{clean[:26]}_{counter}"
        counter += 1
    existing_names.append(candidate)
    return candidate


def build_class_grid_dict(sessions):
    """Returns grid[class_id][day][period] = session"""
    grids = {}
    for item in sessions:
        class_id = item["class_id"]
        days = grids.setdefault(class_id, {})
        periods = days.setdefault(item["day"], {})
        periods[item["period"]] = item
    return grids


def build_teacher_grid_dict(sessions):
    """Returns grid[teacher][day][period] = session (deduplicated for shared classes)"""
    grids = {}
    for item in sessions:
        teacher = item["teacher"]
        days = grids.setdefault(teacher, {})
        periods = days.setdefault(item["day"], {})

        if item["period"] in periods:
            existing = periods[item["period"]]
            classes_set = set(existing.get("class_ids", [existing["class_id"]]))
            classes_set.add(item["class_id"])
            existing["class_ids"] = list(classes_set)
            existing["display_class"] = ", ".join(sorted(classes_set))
            continue

        lesson_copy = dict(item)
        lesson_copy["class_ids"] = [item["class_id"]] + list(item.get("shared_with", []))
        lesson_copy["display_class"] = ", ".join(sorted(set(lesson_copy["class_ids"])))
        periods[item["period"]] = lesson_copy

    return grids


def populate_timetable_sheet(ws, title_text, subtitle_text, days, rows, grid_data, is_teacher_view=False):
    """Renders a standard timetable grid into an openpyxl worksheet."""
    # Title Row
    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = title_text
    title_cell.font = FONT_TITLE
    title_cell.alignment = ALIGN_LEFT
    ws.row_dimensions[1].height = 28

    # Subtitle Row
    ws.merge_cells("A2:G2")
    sub_cell = ws["A2"]
    sub_cell.value = subtitle_text
    sub_cell.font = FONT_SUBTITLE
    sub_cell.alignment = ALIGN_LEFT
    ws.row_dimensions[2].height = 18

    # Blank Row 3
    ws.row_dimensions[3].height = 10

    # Header Row (Row 4)
    start_row = 4
    ws.cell(row=start_row, column=1, value="Period / Time").font = FONT_HEADER
    ws.cell(row=start_row, column=1).fill = FILL_HEADER
    ws.cell(row=start_row, column=1).alignment = ALIGN_CENTER
    ws.cell(row=start_row, column=1).border = BORDER_CELL

    col_idx = 2
    for day in days:
        cell = ws.cell(row=start_row, column=col_idx, value=day)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_CELL
        col_idx += 1

    ws.row_dimensions[start_row].height = 26

    # Data Rows
    current_row = start_row + 1

    for row_info in rows:
        ws.row_dimensions[current_row].height = 42

        if row_info.get("type") == "break":
            # Lunch Break row
            period_label = f"{row_info['label']}\n({row_info['start']} - {row_info['end']})"
            p_cell = ws.cell(row=current_row, column=1, value=period_label)
            p_cell.font = FONT_BREAK
            p_cell.fill = FILL_BREAK
            p_cell.alignment = ALIGN_CENTER
            p_cell.border = BORDER_CELL

            # Merge across all day columns
            start_col = 2
            end_col = len(days) + 1
            ws.merge_cells(
                start_row=current_row, start_column=start_col,
                end_row=current_row, end_column=end_col
            )
            for c in range(start_col, end_col + 1):
                b_cell = ws.cell(row=current_row, column=c)
                b_cell.fill = FILL_BREAK
                b_cell.border = BORDER_CELL
                if c == start_col:
                    b_cell.value = f"★ {row_info['label'].upper()} BREAK ★"
                    b_cell.font = FONT_BREAK
                    b_cell.alignment = ALIGN_CENTER

            current_row += 1
            continue

        # Period Row
        period_num = row_info["period"]
        p_cell = ws.cell(row=current_row, column=1)
        p_cell.value = f"Period {period_num}\n{row_info['start']} - {row_info['end']}"
        p_cell.font = FONT_PERIOD
        p_cell.fill = FILL_PERIOD_COL
        p_cell.alignment = ALIGN_CENTER
        p_cell.border = BORDER_CELL

        col_idx = 2
        for day in days:
            cell = ws.cell(row=current_row, column=col_idx)
            cell.border = BORDER_CELL

            day_slots = grid_data.get(day, {})
            session = day_slots.get(period_num)

            if session:
                cell.fill = FILL_SESSION
                if is_teacher_view:
                    class_info = session.get("display_class", session.get("class_id", ""))
                    cell.value = f"{session['subject']}\nClass: {class_info}"
                else:
                    teacher_info = session.get("teacher", "")
                    shared = session.get("shared_with", [])
                    extra_text = f"\n[Common: {', '.join(shared)}]" if shared else ""
                    cell.value = f"{session['subject']}\nFaculty: {teacher_info}{extra_text}"
                cell.font = FONT_SESSION_SUBJ
                cell.alignment = ALIGN_CENTER
            else:
                cell.fill = FILL_FREE
                cell.value = "—" if not is_teacher_view else "Free"
                cell.font = FONT_FREE
                cell.alignment = ALIGN_CENTER

            col_idx += 1

        current_row += 1

    # Auto-adjust column widths
    ws.column_dimensions["A"].width = 18
    for col in range(2, len(days) + 2):
        col_letter = get_column_letter(col)
        ws.column_dimensions[col_letter].width = 24


def generate_classes_workbook(sessions, config):
    """Generate an Excel workbook with sheets for all classes."""
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet
    existing_sheets = []

    days = config["days"]
    rows = config["rows"]
    class_grids = build_class_grid_dict(sessions)
    classes = sorted(list(class_grids.keys()))

    if not classes:
        ws = wb.create_sheet(title="Timetable")
        ws.cell(row=1, column=1, value="No timetable data available.")
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    # Master Overview Sheet
    master_title = sanitize_sheet_name("Master Summary", existing_sheets)
    ws_master = wb.create_sheet(title=master_title)
    ws_master.cell(row=1, column=1, value="College Master Timetable - All Classes Overview").font = FONT_TITLE
    ws_master.merge_cells("A1:F1")
    ws_master.row_dimensions[1].height = 28

    headers = ["Class", "Day", "Period", "Time", "Subject", "Teacher", "Common With"]
    for col_idx, h in enumerate(headers, start=1):
        c = ws_master.cell(row=3, column=col_idx, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = ALIGN_CENTER
        c.border = BORDER_CELL
    ws_master.row_dimensions[3].height = 24

    m_row = 4
    for item in sessions:
        time_str = f"{item.get('start_time', '')} - {item.get('end_time', '')}"
        shared_str = ", ".join(item.get("shared_with", []))
        vals = [
            item["class_id"], item["day"], f"Period {item['period']}",
            time_str, item["subject"], item["teacher"], shared_str
        ]
        for col_idx, v in enumerate(vals, start=1):
            c = ws_master.cell(row=m_row, column=col_idx, value=v)
            c.font = FONT_SESSION_INFO
            c.alignment = ALIGN_CENTER
            c.border = BORDER_CELL
            if m_row % 2 == 0:
                c.fill = FILL_SESSION
        m_row += 1

    for col in range(1, len(headers) + 1):
        col_letter = get_column_letter(col)
        ws_master.column_dimensions[col_letter].width = 20

    # Individual Class Sheets
    for class_id in classes:
        sheet_name = sanitize_sheet_name(class_id, existing_sheets)
        ws = wb.create_sheet(title=sheet_name)
        populate_timetable_sheet(
            ws,
            title_text=f"Class Timetable: {class_id}",
            subtitle_text=f"Generated Schedule | Academic Session",
            days=days,
            rows=rows,
            grid_data=class_grids.get(class_id, {}),
            is_teacher_view=False
        )

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def generate_single_teacher_workbook(teacher_name, sessions, config):
    """Generate an Excel workbook for a single teacher's schedule."""
    wb = Workbook()
    ws = wb.active
    existing_sheets = []
    sheet_name = sanitize_sheet_name(teacher_name, existing_sheets)
    ws.title = sheet_name

    days = config["days"]
    rows = config["rows"]
    teacher_grids = build_teacher_grid_dict(sessions)
    grid_data = teacher_grids.get(teacher_name, {})

    # Calculate statistics
    teacher_sessions = [s for s in sessions if s.get("teacher") == teacher_name]
    total_periods = sum(s.get("periods_per_session", 1) for s in teacher_sessions)
    unique_subjects = sorted(list(set(s["subject"] for s in teacher_sessions)))
    unique_classes = sorted(list(set(s["class_id"] for s in teacher_sessions)))

    subtitle = (
        f"Weekly Teaching Load: {total_periods} Period(s) | "
        f"Subjects: {', '.join(unique_subjects)} | "
        f"Classes: {', '.join(unique_classes)}"
    )

    populate_timetable_sheet(
        ws,
        title_text=f"Faculty Timetable: {teacher_name}",
        subtitle_text=subtitle,
        days=days,
        rows=rows,
        grid_data=grid_data,
        is_teacher_view=True
    )

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def generate_all_teachers_workbook(sessions, config):
    """Generate an Excel workbook containing all individual teacher timetables."""
    wb = Workbook()
    wb.remove(wb.active)
    existing_sheets = []

    days = config["days"]
    rows = config["rows"]
    teacher_grids = build_teacher_grid_dict(sessions)
    teachers = sorted(list(teacher_grids.keys()))

    if not teachers:
        ws = wb.create_sheet(title="Teachers")
        ws.cell(row=1, column=1, value="No faculty sessions found.")
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    # Summary Sheet
    sum_title = sanitize_sheet_name("Faculty Workload Summary", existing_sheets)
    ws_sum = wb.create_sheet(title=sum_title)
    ws_sum.cell(row=1, column=1, value="Faculty Workload & Teaching Hours Summary").font = FONT_TITLE
    ws_sum.merge_cells("A1:E1")
    ws_sum.row_dimensions[1].height = 28

    headers = ["Faculty Name", "Total Weekly Periods", "Subjects Taught", "Classes Handled", "Active Days"]
    for col_idx, h in enumerate(headers, start=1):
        c = ws_sum.cell(row=3, column=col_idx, value=h)
        c.font = FONT_HEADER
        c.fill = FILL_SUMMARY_HEADER
        c.alignment = ALIGN_CENTER
        c.border = BORDER_CELL
    ws_sum.row_dimensions[3].height = 24

    s_row = 4
    for teacher in teachers:
        t_sessions = [s for s in sessions if s.get("teacher") == teacher]
        total_p = sum(s.get("periods_per_session", 1) for s in t_sessions)
        subjects = ", ".join(sorted(set(s["subject"] for s in t_sessions)))
        classes = ", ".join(sorted(set(s["class_id"] for s in t_sessions)))
        t_days = ", ".join(sorted(set(s["day"] for s in t_sessions), key=lambda d: days.index(d) if d in days else 99))

        vals = [teacher, total_p, subjects, classes, t_days]
        for col_idx, v in enumerate(vals, start=1):
            c = ws_sum.cell(row=s_row, column=col_idx, value=v)
            c.font = FONT_SESSION_INFO
            c.alignment = ALIGN_CENTER if col_idx in (2, 5) else ALIGN_LEFT
            c.border = BORDER_CELL
            if s_row % 2 == 0:
                c.fill = FILL_SESSION
        s_row += 1

    ws_sum.column_dimensions["A"].width = 22
    ws_sum.column_dimensions["B"].width = 22
    ws_sum.column_dimensions["C"].width = 30
    ws_sum.column_dimensions["D"].width = 28
    ws_sum.column_dimensions["E"].width = 25

    # Sheet for each teacher
    for teacher in teachers:
        sheet_title = sanitize_sheet_name(teacher, existing_sheets)
        ws = wb.create_sheet(title=sheet_title)

        t_sessions = [s for s in sessions if s.get("teacher") == teacher]
        total_periods = sum(s.get("periods_per_session", 1) for s in t_sessions)
        unique_subjects = sorted(list(set(s["subject"] for s in t_sessions)))
        unique_classes = sorted(list(set(s["class_id"] for s in t_sessions)))

        subtitle = (
            f"Weekly Load: {total_periods} Period(s) | "
            f"Subjects: {', '.join(unique_subjects)} | "
            f"Classes: {', '.join(unique_classes)}"
        )

        populate_timetable_sheet(
            ws,
            title_text=f"Faculty Timetable: {teacher}",
            subtitle_text=subtitle,
            days=days,
            rows=rows,
            grid_data=teacher_grids.get(teacher, {}),
            is_teacher_view=True
        )

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
