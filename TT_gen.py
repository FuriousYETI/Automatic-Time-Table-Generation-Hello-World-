# TT_gen.py -- Timetable generator with room allocation and global room conflict avoidance
# Run: python TT_gen.py
# Requires: pandas, openpyxl
import pandas as pd
import random
from datetime import datetime, time, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
from dataclasses import dataclass
import traceback
from datetime import date, timedelta
# ---------------------------
# Constants and durations (minutes)
# ---------------------------
DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

LECTURE_MIN = 90   # 1.5 hours
LAB_MIN = 120      # 2 hours
TUTORIAL_MIN = 60  # 1 hour
SELF_STUDY_MIN = 60

# Break windows
MORNING_BREAK_START = time(10, 30)
MORNING_BREAK_END = time(10, 45)
LUNCH_BREAK_START = time(13, 0)
LUNCH_BREAK_END = time(13, 45)

# ---------------------------
# Dataclasses
# ---------------------------
@dataclass
class UnscheduledComponent:
    department: str
    semester: int
    code: str
    name: str
    faculty: str
    component_type: str
    sessions: int
    section: int
    reason: str

# ---------------------------
# Load CSVs
# ---------------------------
try:
    df = pd.read_csv('combined.csv')
except FileNotFoundError:
    raise SystemExit("Error: 'combined.csv' not found in working directory.")

try:
    rooms_df = pd.read_csv('rooms.csv')
except FileNotFoundError:
    rooms_df = pd.DataFrame(columns=['roomNumber', 'type'])

# Normalize rooms lists (case-insensitive)
lecture_rooms = rooms_df[rooms_df.get('type', '').str.upper() == 'LECTURE_ROOM']['roomNumber'].tolist()
computer_lab_rooms = rooms_df[rooms_df.get('type', '').str.upper() == 'COMPUTER_LAB']['roomNumber'].tolist()
large_rooms = rooms_df[rooms_df.get('type', '').str.upper() == 'SEATER_120']['roomNumber'].tolist()

# ---------------------------
# Custom time slot definitions (irregular)
# ---------------------------
def generate_time_slots():
    slots = [
        (time(7, 30), time(9, 0)),     # Minor Slot (morning)
        (time(9, 0), time(10, 0)),
        (time(10, 0), time(10, 30)),
        (time(10, 30), time(10, 45)),  # Short break
        (time(10, 45), time(11, 0)),
        (time(11, 0), time(11, 30)),
        (time(11, 30), time(12, 0)),
        (time(12, 0), time(12, 15)),
        (time(12, 15), time(12, 30)),
        (time(12, 30), time(13, 15)),
        (time(13, 15), time(13, 30)),
        (time(13, 30), time(14, 0)),
        (time(14, 0), time(14, 30)),
        (time(14, 30), time(15, 30)),
        (time(15, 30), time(15, 40)),
        (time(16, 0), time(16, 30)),
        (time(16, 30), time(17, 10)),
        (time(17, 10), time(17, 30)),
        (time(17, 30), time(18, 30)),
        (time(18, 30), time(23, 59)),  # Minor Slot (evening)
    ]
    return slots

TIME_SLOTS = generate_time_slots()

# ---------------------------
# Helpers
# ---------------------------
def slot_minutes(slot):
    s, e = slot
    s_m = s.hour*60 + s.minute
    e_m = e.hour*60 + e.minute
    if e_m < s_m:
        e_m += 24*60
    return e_m - s_m

def overlaps(a_start, a_end, b_start, b_end):
    return (a_start < b_end) and (b_start < a_end)

def is_break_time_slot(slot, semester=None):
    start, end = slot
    if overlaps(start, end, MORNING_BREAK_START, MORNING_BREAK_END):
        return True
    if overlaps(start, end, LUNCH_BREAK_START, LUNCH_BREAK_END):
        return True
    return False

def is_minor_slot(slot):
    start, end = slot
    if start == time(7, 30) and end == time(9, 0):
        return True
    if start == time(18, 30):
        return True
    return False

def select_faculty(faculty_field):
    if pd.isna(faculty_field) or str(faculty_field).strip().lower() in ['nan', 'none', '']:
        return "TBD"
    s = str(faculty_field).strip()
    for sep in ['/', ',', '&', ';']:
        if sep in s:
            return s.split(sep)[0].strip()
    return s

def get_course_priority(row):
    try:
        l = int(row.get('L', 0)) if pd.notna(row.get('L', 0)) else 0
        t = int(row.get('T', 0)) if pd.notna(row.get('T', 0)) else 0
        p = int(row.get('P', 0)) if pd.notna(row.get('P', 0)) else 0
        return -(l + t + p)
    except Exception:
        return 0

def calculate_required_minutes(course_row):
    l = int(course_row['L']) if ('L' in course_row and pd.notna(course_row['L'])) else 0
    t = int(course_row['T']) if ('T' in course_row and pd.notna(course_row['T'])) else 0
    p = int(course_row['P']) if ('P' in course_row and pd.notna(course_row['P'])) else 0
    s = int(course_row['S']) if ('S' in course_row and pd.notna(course_row['S'])) else 0
    return (l, t, p, s)

def get_required_room_type(course_row):
    try:
        p = int(course_row['P']) if ('P' in course_row and pd.notna(course_row['P'])) else 0
        return 'COMPUTER_LAB' if p > 0 else 'LECTURE_ROOM'
    except Exception:
        return 'LECTURE_ROOM'

# ---------------------------
# Room allocation helpers (global room_schedule)
# ---------------------------
def find_suitable_room_for_slot(course_code, room_type, day, slot_indices, room_schedule, course_room_mapping):
    """
    Finds a suitable room for the given course and slots.
    Ensures:
    - Each course always gets the same room across all its sessions.
    - No two courses overlap in the same room at the same time.
    """
    # if this course already has a room, reuse it (if it's free for these slots)
    if course_code in course_room_mapping:
        fixed_room = course_room_mapping[course_code]
        for si in slot_indices:
            if si in room_schedule[fixed_room][day]:
                # room occupied -> cannot use, fail
                return None
        return fixed_room

    # else assign a new room (first time this course is scheduled)
    pool = computer_lab_rooms if room_type == 'COMPUTER_LAB' else lecture_rooms
    if not pool:
        return None
    random.shuffle(pool)
    for room in pool:
        if room not in room_schedule:
            room_schedule[room] = {d: set() for d in range(len(DAYS))}
        # check if room is free for all these slots
        if all(si not in room_schedule[room][day] for si in slot_indices):
            # assign this room permanently to this course
            course_room_mapping[course_code] = room
            return room
    return None

def find_consecutive_slots_for_minutes(timetable, day, start_idx, required_minutes,
                                       semester, professor_schedule, faculty,
                                       room_schedule, room_type, course_code, course_room_mapping):
    """
    Find consecutive TIME_SLOTS starting at start_idx whose total minutes >= required_minutes,
    respecting minor slots, breaks, existing timetable occupancy, professor schedule, and
    room availability. Returns (slot_indices, room) or (None, None).
    """
    n = len(TIME_SLOTS)
    slot_indices = []
    i = start_idx
    accumulated = 0

    # accumulate consecutive slots
    while i < n and accumulated < required_minutes:
        # can't schedule in minor slot or break
        if is_minor_slot(TIME_SLOTS[i]) or is_break_time_slot(TIME_SLOTS[i], semester):
            return None, None
        # slot already occupied in this timetable
        if timetable[day][i]['type'] is not None:
            return None, None
        # professor busy
        if faculty in professor_schedule and i in professor_schedule[faculty][day]:
            return None, None
        # if no rooms of required type exist at all, fail early
        if room_type == 'COMPUTER_LAB' and not computer_lab_rooms:
            return None, None
        if room_type == 'LECTURE_ROOM' and not lecture_rooms:
            return None, None

        slot_indices.append(i)
        accumulated += slot_minutes(TIME_SLOTS[i])
        i += 1

    # if we gathered enough minutes, find a room free across these slots
    if accumulated >= required_minutes:
        room = find_suitable_room_for_slot(course_code, room_type, day, slot_indices, room_schedule, course_room_mapping)
        if room is not None:
            return slot_indices, room

    return None, None

def get_all_possible_start_indices_for_duration():
    idxs = list(range(len(TIME_SLOTS)))
    random.shuffle(idxs)
    return idxs

def check_professor_availability(professor_schedule, faculty, day, start_idx, duration_slots):
    if faculty not in professor_schedule:
        return True
    if not professor_schedule[faculty][day]:
        return True
    new_start = TIME_SLOTS[start_idx][0]
    new_start_m = new_start.hour*60 + new_start.minute
    MIN_GAP = 180
    for s in professor_schedule[faculty][day]:
        exist_start = TIME_SLOTS[s][0]
        exist_m = exist_start.hour*60 + exist_start.minute
        if abs(exist_m - new_start_m) < MIN_GAP:
            return False
    return True

def load_rooms():
    return {'lecture_rooms': lecture_rooms, 'computer_lab_rooms': computer_lab_rooms, 'large_rooms': large_rooms}

def load_batch_data():
    batch_info = {}
    for _, r in df.iterrows():
        dept = str(r['Department'])
        sem = int(r['Semester'])
        batch_info[(dept, sem)] = {'num_sections': 1}
    return batch_info

# ---------------------------
# Main generation function
# ---------------------------
def generate_all_timetables():
    global TIME_SLOTS
    TIME_SLOTS = generate_time_slots()
    rooms = load_rooms()
    batch_info = load_batch_data()

    room_schedule = {}
    professor_schedule = {}
    course_room_mapping = {}

    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    overview = wb.create_sheet("Overview")
    overview.append(["Combined Timetable for All Departments and Semesters"])
    overview.append(["Generated on:", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    overview.append([])
    overview.append(["Department", "Semester", "Sheet Name"])
    row_index = 5

    unscheduled_components = []

    SUBJECT_COLORS = [
        "FF6B6B", "4ECDC4", "FF9F1C", "5D5FEF", "45B7D1",
        "F72585", "7209B7", "3A0CA3", "4361EE", "4CC9F0",
        "06D6A0", "FFD166", "EF476F", "118AB2", "073B4C"
    ]

    for department in df['Department'].unique():
        sems = sorted(df[df['Department'] == department]['Semester'].unique())
        for semester in sems:
# ---------------------------
# Section and Priority Rules
# ---------------------------

# Give 2 sections for CSE, ECE, and DSAI in semesters 2, 4, 6
            dept_upper = str(department).strip().upper()
            num_sections = 2 if (dept_upper in ["CSE", "ECE", "DSAI"] and int(semester) in [2, 4, 6]) else 1

            courses = df[(df['Department'] == department) & (df['Semester'] == semester)]
            if 'Schedule' in courses.columns:
                courses = courses[(courses['Schedule'].fillna('Yes').str.upper() == 'YES') | (courses['Schedule'].isna())]
            if courses.empty:
                continue

            # Split into lab and non-lab courses
            if 'P' in courses.columns:
                lab_courses = courses[courses['P'] > 0].copy()
                non_lab_courses = courses[courses['P'] == 0].copy()
            else:
                lab_courses = courses.head(0)
                non_lab_courses = courses.copy()

            # Priority by total workload (L + T + P)
            if not lab_courses.empty:
                lab_courses['priority'] = lab_courses.apply(get_course_priority, axis=1)
                lab_courses = lab_courses.sort_values('priority', ascending=False)
            non_lab_courses['priority'] = non_lab_courses.apply(get_course_priority, axis=1)
            non_lab_courses = non_lab_courses.sort_values('priority', ascending=False)

            # --- ELECTIVE PRIORITY LOGIC ---
            def is_elective(course_row):
                name = str(course_row.get('Course Name', '')).lower()
                code = str(course_row.get('Course Code', '')).lower()
                keywords = ["elective", "oe", "open elective", "pe", "program elective"]
                return any(k in name for k in keywords) or any(k in code for k in keywords)

            combined = pd.concat([lab_courses, non_lab_courses])
            combined['is_elective'] = combined.apply(is_elective, axis=1)

            # Electives first, then core — higher total workload within each group
            courses_combined = combined.sort_values(by=['is_elective', 'priority'], ascending=[False, False]).drop_duplicates()

            for section in range(num_sections):
                section_title = f"{department}_{semester}" if num_sections == 1 else f"{department}_{semester}_{chr(65 + section)}"
                ws = wb.create_sheet(title=section_title)

                overview.cell(row=row_index, column=1, value=department)
                overview.cell(row=row_index, column=2, value=str(semester))
                overview.cell(row=row_index, column=3, value=section_title)
                row_index += 1

                timetable = {d: {s: {'type': None, 'code': '', 'name': '', 'faculty': '', 'classroom': ''} for s in range(len(TIME_SLOTS))} for d in range(len(DAYS))}

                section_subject_color = {}
                color_iter = iter(SUBJECT_COLORS)
                course_faculty_map = {}

                for _, c in courses_combined.iterrows():
                    code = str(c.get('Course Code', '')).strip()
                    if code and code not in section_subject_color:
                        try:
                            section_subject_color[code] = next(color_iter)
                        except StopIteration:
                            section_subject_color[code] = random.choice(SUBJECT_COLORS)
                        course_faculty_map[code] = select_faculty(c.get('Faculty', 'TBD'))
                # --- PRIORITIZE ELECTIVES FIRST ---
                def is_elective(course_name):
                    if pd.isna(course_name):
                        return False
                    name = str(course_name).lower()
                    keywords = ["elective", "oe", "open elective", "pe", "program elective"]
                    return any(k in name for k in keywords)

                # Separate electives and non-electives
                elective_courses = courses_combined[courses_combined['Course Name'].apply(is_elective)]
                core_courses = courses_combined[~courses_combined['Course Name'].apply(is_elective)]

                # Recombine — electives first, then core
                courses_combined = pd.concat([elective_courses, core_courses])


                for _, course in courses_combined.iterrows():
                    code = str(course.get('Course Code', '')).strip()
                    name = str(course.get('Course Name', '')).strip()
                    faculty = select_faculty(course.get('Faculty', 'TBD'))

                    if faculty not in professor_schedule:
                        professor_schedule[faculty] = {d: set() for d in range(len(DAYS))}

                    lec_count, tut_count, lab_count, ss_count = calculate_required_minutes(course)
                    room_type = get_required_room_type(course)

                    def schedule_component(required_minutes, comp_type, attempts_limit=800):
                        for attempt in range(attempts_limit):
                            day = random.randint(0, len(DAYS)-1)
                            starts = get_all_possible_start_indices_for_duration()
                            for start_idx in starts:
                                slot_indices, candidate_room = find_consecutive_slots_for_minutes(
    timetable, day, start_idx, required_minutes, semester,
    professor_schedule, faculty, room_schedule, room_type,
    code, course_room_mapping
)

                                if slot_indices is None:
                                    continue
                                if not check_professor_availability(professor_schedule, faculty, day, slot_indices[0], len(slot_indices)):
                                    continue
                                if candidate_room is None:
                                    continue
                                for si_idx, si in enumerate(slot_indices):
                                    timetable[day][si]['type'] = 'LEC' if comp_type == 'LEC' else ('LAB' if comp_type == 'LAB' else ('TUT' if comp_type == 'TUT' else 'SS'))
                                    timetable[day][si]['code'] = code if si_idx == 0 else ''
                                    timetable[day][si]['name'] = name if si_idx == 0 else ''
                                    timetable[day][si]['faculty'] = faculty if si_idx == 0 else ''
                                    timetable[day][si]['classroom'] = candidate_room if si_idx == 0 else ''
                                    professor_schedule[faculty][day].add(si)
                                    if candidate_room not in room_schedule:
                                        room_schedule[candidate_room] = {d: set() for d in range(len(DAYS))}
                                    room_schedule[candidate_room][day].add(si)
                                return True
                        return False

                    for _ in range(lec_count):
                        ok = schedule_component(LECTURE_MIN, 'LEC', attempts_limit=800)
                        if not ok:
                            unscheduled_components.append(UnscheduledComponent(department, semester, code, name, faculty, 'LEC', 1, section, "Lecture not scheduled"))

                    for _ in range(tut_count):
                        ok = schedule_component(TUTORIAL_MIN, 'TUT', attempts_limit=600)
                        if not ok:
                            unscheduled_components.append(UnscheduledComponent(department, semester, code, name, faculty, 'TUT', 1, section, "Tutorial not scheduled"))

                    for _ in range(lab_count):
                        ok = schedule_component(LAB_MIN, 'LAB', attempts_limit=800)
                        if not ok:
                            unscheduled_components.append(UnscheduledComponent(department, semester, code, name, faculty, 'LAB', 1, section, "Lab not scheduled"))

                    for _ in range(ss_count):
                        ok = schedule_component(SELF_STUDY_MIN, 'SS', attempts_limit=400)
                        if not ok:
                            unscheduled_components.append(UnscheduledComponent(department, semester, code, name, faculty, 'SS', 1, section, "Self-study not scheduled"))

                # Write sheet
                header = ['Day'] + [f"{slot[0].strftime('%H:%M')}-{slot[1].strftime('%H:%M')}" for slot in TIME_SLOTS]
                ws.append(header)
                header_fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
                header_font = Font(bold=True)
                header_alignment = Alignment(horizontal='center', vertical='center')
                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = header_alignment

                border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                lec_fill_default = PatternFill(start_color="FA8072", end_color="FA8072", fill_type="solid")
                lab_fill_default = PatternFill(start_color="7CFC00", end_color="7CFC00", fill_type="solid")
                tut_fill_default = PatternFill(start_color="87CEEB", end_color="87CEEB", fill_type="solid")
                ss_fill_default = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
                break_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
                minor_fill = PatternFill(start_color="9ACD32", end_color="9ACD32", fill_type="solid")

                for day_idx, day_name in enumerate(DAYS):
                    ws.append([day_name] + [''] * len(TIME_SLOTS))
                    row_num = ws.max_row
                    merges = []
                    for slot_idx in range(len(TIME_SLOTS)):
                        cell_obj = ws.cell(row=row_num, column=slot_idx + 2)
                        if is_minor_slot(TIME_SLOTS[slot_idx]):
                            cell_obj.value = "Minor Slot"
                            cell_obj.fill = minor_fill
                            cell_obj.font = Font(bold=True)
                            cell_obj.alignment = Alignment(horizontal='center', vertical='center')
                            cell_obj.border = border
                            continue

                        if is_break_time_slot(TIME_SLOTS[slot_idx], semester):
                            cell_obj.value = "BREAK"
                            cell_obj.fill = break_fill
                            cell_obj.font = Font(bold=True)
                            cell_obj.alignment = Alignment(horizontal='center', vertical='center')
                            cell_obj.border = border
                            continue

                        if timetable[day_idx][slot_idx]['type'] is None:
                            cell_obj.border = border
                            continue

                        typ = timetable[day_idx][slot_idx]['type']
                        code = timetable[day_idx][slot_idx]['code']
                        cls = timetable[day_idx][slot_idx]['classroom']
                        fac = timetable[day_idx][slot_idx]['faculty']

                        if code:
                            span = [slot_idx]
                            j = slot_idx + 1
                            while j < len(TIME_SLOTS) and timetable[day_idx][j]['type'] is not None and timetable[day_idx][j]['code'] == '':
                                span.append(j)
                                j += 1
                            display = f"{code} {typ}\nroom no. :{cls}\n{fac}"

                            if code in section_subject_color:
                                subj_color = section_subject_color[code]
                                fill = PatternFill(start_color=subj_color, end_color=subj_color, fill_type="solid")
                            else:
                                fill = {'LEC': lec_fill_default, 'LAB': lab_fill_default, 'TUT': tut_fill_default, 'SS': ss_fill_default}.get(typ, lec_fill_default)

                            cell_obj.value = display
                            cell_obj.fill = fill
                            merges.append((slot_idx + 2, slot_idx + 2 + len(span) - 1, display, fill))
                        cell_obj.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                        cell_obj.border = border

                    for start_col, end_col, val, fill in merges:
                        if end_col > start_col:
                            rng = f"{get_column_letter(start_col)}{row_num}:{get_column_letter(end_col)}{row_num}"
                            try:
                                ws.merge_cells(rng)
                                mc = ws[f"{get_column_letter(start_col)}{row_num}"]
                                mc.value = val
                                mc.fill = fill
                                mc.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                                mc.border = border
                            except Exception:
                                pass

                for col_idx in range(1, len(TIME_SLOTS) + 2):
                    ws.column_dimensions[get_column_letter(col_idx)].width = 15
                for r in range(2, 2 + len(DAYS)):
                    try:
                        ws.row_dimensions[r].height = 40
                    except Exception:
                        pass

                # ---------- LEGEND TABLE ----------
                legend_start = len(DAYS) + 4
                ws.cell(row=legend_start, column=1, value="Legend").font = Font(bold=True, size=12)
                legend_start += 1

                # Add a fifth column for actual room numbers
                headers = ["Code", "Color", "Course Name", "Faculty", "Room Numbers"]
                for i, header in enumerate(headers, start=1):
                    cell = ws.cell(row=legend_start, column=i, value=header)
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                legend_start += 1

                border_style = Border(left=Side(style='thin'), right=Side(style='thin'),
                                    top=Side(style='thin'), bottom=Side(style='thin'))

                # Build legend rows
                for code, color in section_subject_color.items():
                    rn = courses_combined[courses_combined['Course Code'] == code]
                    name_val = str(rn['Course Name'].iloc[0]) if not rn.empty else ''
                    fac_val = str(rn['Faculty'].iloc[0]) if not rn.empty else ''

                    # Collect all unique rooms used for this course in this section
                    used_rooms = set()
                    for day_idx in range(len(DAYS)):
                        for slot_idx in range(len(TIME_SLOTS)):
                            if timetable[day_idx][slot_idx]['code'] == code or timetable[day_idx][slot_idx]['name'] == name_val:
                                room_val = timetable[day_idx][slot_idx]['classroom']
                                if room_val:
                                    used_rooms.add(str(room_val))
                    rooms_str = ", ".join(sorted(list(used_rooms))) if used_rooms else "-"

                    ws.cell(row=legend_start, column=1, value=code)
                    color_cell = ws.cell(row=legend_start, column=2, value="")
                    color_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                    ws.cell(row=legend_start, column=3, value=name_val)
                    ws.cell(row=legend_start, column=4, value=fac_val)
                    ws.cell(row=legend_start, column=5, value=rooms_str)

                    for col in range(1, 6):
                        c = ws.cell(row=legend_start, column=col)
                        c.border = border_style
                        c.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)

                    ws.row_dimensions[legend_start].height = 25
                    legend_start += 1

                # Adjust column widths
                ws.column_dimensions[get_column_letter(1)].width = 12  # Code
                ws.column_dimensions[get_column_letter(2)].width = 10  # Color box
                ws.column_dimensions[get_column_letter(3)].width = 45  # Course Name
                ws.column_dimensions[get_column_letter(4)].width = 40  # Faculty
                ws.column_dimensions[get_column_letter(5)].width = 25  # Room Numbers


    for col in range(1, 4):
        overview.column_dimensions[get_column_letter(col)].width = 22
    for cell in overview[4]:
        try:
            cell.fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        except Exception:
            pass

    out_name = "timetable_all_departments.xlsx"
    wb.save(out_name)
    print(f"Combined timetable saved as {out_name}")
def generate_exam_timetable():
    print("\n🧾 Generating formatted exam timetable (exam_timetable.xlsx)...")

    try:
        df = pd.read_csv("combined_faculty_fullnames_doctorized.csv")
    except FileNotFoundError:
        df = pd.read_csv("combined.csv")

    df = df[df["Schedule"].fillna("Yes").str.upper() == "YES"]

    # Create workbook and sheets
    wb = Workbook()
    ws = wb.active
    ws.title = "Exam Timetable"
    legend = wb.create_sheet("Legend")

    # ===== HEADER =====
    ws.merge_cells("A1:K1")
    ws["A1"] = "Indian Institute of Information Technology Dharwad"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:K2")
    ws["A2"] = "Time table for All Departments - End Semester Exam (Nov 2025)"
    ws["A2"].font = Font(size=12, bold=True)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A3:K3")
    ws["A3"] = "AN: 03:00 PM to 04:30 PM"
    ws["A3"].font = Font(size=11, bold=True)
    ws["A3"].alignment = Alignment(horizontal="center", vertical="center")

    # ===== DATE RANGE =====
    start_date = date(2025, 11, 20)
    num_days = 10
    exam_dates = []
    d = start_date
    while len(exam_dates) < num_days:
        if d.weekday() != 6:  # Skip Sundays
            exam_dates.append(d)
        d += timedelta(days=1)

    # ===== TABLE HEADERS =====
    ws["A5"] = "Department - Sem"
    ws["A5"].font = Font(bold=True)
    ws["A5"].alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions["A"].width = 25

    for i, dt in enumerate(exam_dates):
        col = get_column_letter(i + 2)
        ws[f"{col}5"] = dt.strftime("%d-%b-%Y")
        ws[f"{col}5"].font = Font(bold=True)
        ws[f"{col}5"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{col}6"] = dt.strftime("%A")
        ws[f"{col}6"].font = Font(italic=True)
        ws[f"{col}6"].alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[col].width = 22

    # ===== GROUP COURSES BY CLASS =====
    df["Class"] = df["Department"].astype(str) + "_Sem" + df["Semester"].astype(str)
    class_groups = df.groupby("Class")

    # ===== LEGEND SHEET =====
    legend.append(["Course Code", "Course Name", "Faculty Name"])
    legend["A1"].font = legend["B1"].font = legend["C1"].font = Font(bold=True)
    legend.column_dimensions["A"].width = 20   # Course Code
    legend.column_dimensions["B"].width = 45   # Course Name
    legend.column_dimensions["C"].width = 30   # Faculty Name

    unique_courses = (
        df[["Course Code", "Course Name", "Faculty"]]
            .drop_duplicates()
            .sort_values("Course Code")
            .reset_index(drop=True)
    )
    for _, row in unique_courses.iterrows():
        legend.append([row["Course Code"], row["Course Name"], row["Faculty"]])


    # ===== SCHEDULING =====
    row = 7
    for class_name, group in class_groups:
        ws[f"A{row}"] = class_name
        ws[f"A{row}"].font = Font(bold=True, color="0000FF")
        ws[f"A{row}"].alignment = Alignment(horizontal="center", vertical="center")

        courses = group[["Course Code", "Faculty"]].drop_duplicates().values.tolist()
        random.shuffle(courses)

        for i, (course, _) in enumerate(courses):
            date_index = i % len(exam_dates)
            col = get_column_letter(2 + date_index)
            ws[f"{col}{row}"] = course  # ✅ Only course code
            ws[f"{col}{row}"].alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )

        row += 1

    # ===== STYLING =====
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for r in ws.iter_rows(min_row=5, max_row=row, min_col=1, max_col=len(exam_dates) + 1):
        for cell in r:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ===== SAVE =====
    wb.save("exam_timetable.xlsx")
    print("✅ Exam timetable saved as exam_timetable.xlsx (with Legend sheet)")

# ---------------------------
# Run
# ---------------------------
if __name__ == "__main__":
    print("\n🚀 Generating all timetables...")
    try:
        generate_all_timetables()     # Class timetable
        generate_exam_timetable()     # Exam timetable
        print("\n✅ All timetables generated successfully!")
    except Exception:
        print("\n❌ Error while generating timetables:")
        traceback.print_exc()
