import pandas as pd
import json, random, colorsys
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

DATA_DIR = Path("data")
TIME_SLOTS_FILE = DATA_DIR / "time_slots.json"
COURSES_FILE = DATA_DIR / "courses.csv"
ROOMS_FILE = DATA_DIR / "rooms.csv"

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
EXCLUDED_SLOTS = ["07:30-09:00", "13:15-14:00", "17:30-18:30"]
MAX_ATTEMPTS = 10

def load_time_slots(path):
    with open(path) as f:
        slots = json.load(f)["time_slots"]
    return [f"{s['start'].strip()}-{s['end'].strip()}" for s in slots]

def slot_duration(slot):
    h1, m1 = map(int, slot.split("-")[0].split(":"))
    h2, m2 = map(int, slot.split("-")[1].split(":"))
    return (h2 + m2 / 60) - (h1 + m1 / 60)

def generate_colors(n):
    colors = []
    for i in range(n):
        hue = i / n
        light, sat = random.uniform(0.45, 0.65), random.uniform(0.7, 0.9)
        r, g, b = colorsys.hls_to_rgb(hue, light, sat)
        colors.append(f"{int(r*255):02X}{int(g*255):02X}{int(b*255):02X}")
    random.shuffle(colors)
    return colors

class Course:
    def __init__(self, record):
        self.code = record.get("Course_Code", "").strip()
        self.faculty = record.get("Faculty", "").strip()
        self.sem_half = str(record.get("Semester_Half", "0")).strip()
        self.is_elective = str(record.get("Elective", "0")) == "1"
        try:
            self.L, self.T, self.P, *_ = map(int, record["L-T-P-S-C"].split("-"))
        except:
            self.L, self.T, self.P = 0, 0, 0

class RoomPool:
    def __init__(self, df):
        self.classrooms = df[df["Type"].str.lower() == "classroom"]["Room_ID"].tolist()
        self.labs = df[df["Type"].str.lower() == "lab"]["Room_ID"].tolist()
    def pick(self, session_type):
        if session_type == "P" and self.labs:
            return random.choice(self.labs)
        if self.classrooms:
            return random.choice(self.classrooms)
        return None

class TimetableGenerator:
    def __init__(self, courses_df, rooms_df):
        self.courses = [Course(c) for c in courses_df]
        self.rooms = RoomPool(rooms_df)
        self.slot_keys = load_time_slots(TIME_SLOTS_FILE)
        self.slot_durations = {s: slot_duration(s) for s in self.slot_keys}
        self.palette = generate_colors(50)
    def new_table(self):
        return pd.DataFrame("", index=DAYS, columns=self.slot_keys)
    def allocate_course(self, df, course):
        needed = {"L": course.L, "T": course.T, "P": course.P}
        for session, hours in needed.items():
            while hours > 0:
                day = random.choice(DAYS)
                slot = random.choice(self.slot_keys)
                if slot in EXCLUDED_SLOTS or df.at[day, slot]:
                    continue
                room = self.rooms.pick(session)
                df.at[day, slot] = f"{course.code} ({room})"
                hours -= 1
    def generate(self, sem_half="1"):
        records = [c for c in self.courses if c.sem_half in [sem_half, "0"]]
        df = self.new_table()
        for c in records:
            self.allocate_course(df, c)
        return df

class ExcelStyler:
    def __init__(self, palette):
        self.palette = palette
        self.align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        self.border = Border(left=Side(style="thin"), right=Side(style="thin"),
                             top=Side(style="thin"), bottom=Side(style="thin"))
    def apply_style(self, path, df):
        df.to_excel(path)
        wb = load_workbook(path)
        ws = wb.active
        colors = generate_colors(len(df.columns))
        for r in ws.iter_rows(min_row=2, min_col=2):
            for cell in r:
                val = cell.value
                if not val:
                    continue
                color = random.choice(colors)
                cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                cell.alignment = self.align
                cell.border = self.border
        wb.save(path)

def main():
    courses = pd.read_csv(COURSES_FILE).to_dict(orient="records")
    rooms = pd.read_csv(ROOMS_FILE)
    gen = TimetableGenerator(courses, rooms)
    styler = ExcelStyler(gen.palette)
    t1 = gen.generate("1")
    t2 = gen.generate("2")
    styler.apply_style("TT_first_half.xlsx", t1)
    styler.apply_style("TT_second_half.xlsx", t2)
    print("Saved: TT_first_half.xlsx, TT_second_half.xlsx")

if __name__ == "__main__":
    main()
