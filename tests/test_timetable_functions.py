import pytest
import pandas as pd
from datetime import time
from TT_gen import (
    generate_time_slots,
    calculate_lunch_breaks,
    check_professor_availability,
    is_break_time,
    sanitize_sheet_name,
    sanitize_filename,
    extract_faculty_names,
    load_config,
    generate_course_color
)

# ----------------------------------------------
# 1️⃣ Basic time slot generation
# ----------------------------------------------
def test_generate_time_slots():
    slots = generate_time_slots()
    assert len(slots) > 0
    assert slots[0][0].strftime("%H:%M") == "09:00"
    assert slots[-1][1].strftime("%H:%M") == "18:30"

# ----------------------------------------------
# 2️⃣ Lunch break calculation
# ----------------------------------------------
def test_calculate_lunch_breaks():
    breaks = calculate_lunch_breaks([1, 2, 3])
    assert len(breaks) == 3
    for sem, (start, end) in breaks.items():
        assert isinstance(start, time)
        assert isinstance(end, time)

def test_calculate_lunch_breaks_empty():
    assert calculate_lunch_breaks([]) == {}

# ----------------------------------------------
# 3️⃣ Professor availability conflict
# ----------------------------------------------
def test_professor_conflict_overlap():
    professor_schedule = {'Dr. A': {0: {3, 4, 5}}}
    assert not check_professor_availability(professor_schedule, 'Dr. A', 0, 4, 3, 'LEC')

def test_professor_conflict_clear():
    professor_schedule = {'Dr. A': {0: {1}}}
    assert check_professor_availability(professor_schedule, 'Dr. A', 0, 10, 3, 'LAB')

# ----------------------------------------------
# 4️⃣ Break time detection
# ----------------------------------------------
def test_is_break_time():
    assert is_break_time((time(10, 30), time(11, 0))) == True
    assert is_break_time((time(14, 0), time(14, 30))) == False

# ----------------------------------------------
# 5️⃣ String sanitization for Excel and filenames
# ----------------------------------------------
def test_sanitize_sheet_name():
    name = "Dr./John:Smith*Dept"
    assert sanitize_sheet_name(name) == "Dr._John_Smith_Dept"

def test_sanitize_filename():
    name = "Prof*John|Doe?.xlsx"
    assert sanitize_filename(name).startswith("Prof_John_Doe")

# ----------------------------------------------
# 6️⃣ Faculty name extraction
# ----------------------------------------------
def test_extract_faculty_names_ampersand():
    assert extract_faculty_names("Dr. A & Dr. B") == ["Dr. A", "Dr. B"]

def test_extract_faculty_names_slash():
    assert extract_faculty_names("Prof. A / Prof. B") == ["Prof. A", "Prof. B"]

def test_extract_faculty_names_commas():
    assert extract_faculty_names("Dr. X, Dr. Y, Dr. Z") == ["Dr. X", "Dr. Y", "Dr. Z"]

def test_extract_faculty_names_and():
    assert extract_faculty_names("Dr. A and Dr. B") == ["Dr. A", "Dr. B"]

# ----------------------------------------------
# 7️⃣ Config loading
# ----------------------------------------------
def test_load_config_defaults(monkeypatch):
    # Simulate missing config.json
    monkeypatch.setattr("builtins.open", lambda f, m='r': (_ for _ in ()).throw(FileNotFoundError()))
    config = load_config()
    assert "lecture_duration" in config

# ----------------------------------------------
# 8️⃣ Color generation
# ----------------------------------------------
def test_generate_course_color():
    color_gen = generate_course_color()
    first_color = next(color_gen)
    for _ in range(15):
        color = next(color_gen)
        assert len(color) == 6

