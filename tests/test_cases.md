# 🧪 Unit Test Cases for Timetable Generation Functions

| **Test Case Input** | **Description** | **Expected Output** |
|----------------------|------------------|----------------------|
| `generate_time_slots()` | Generate all half-hour slots from 09:00 to 18:30 | List of 19 time slots (e.g., 09:00–09:30 … 18:00–18:30) |
| `calculate_lunch_breaks([1, 2, 3])` | Test staggered lunch assignment for 3 semesters | `{1: (12:30–13:30), 2: (13:00–14:00), 3: (13:30–14:30)}` (approx values) |
| `calculate_lunch_breaks([])` | No semesters input | `{}` (empty dictionary) |
| `check_professor_availability({'Dr.A': {0: {3,4,5}}}, 'Dr.A', 0, 4, 3, 'LEC')` | Professor has class from slots 3–5, new lecture starts at 4 | `False` (conflict) |
| `check_professor_availability({'Dr.A': {0: {1}}}, 'Dr.A', 0, 10, 3, 'LAB')` | Professor has early class, new lab starts later | `True` (no conflict) |
| `is_break_time((time(10,30), time(11,0)))` | Slot within morning break | `True` |
| `is_break_time((time(14,0), time(14,30)))` | Slot outside break | `False` |
| `sanitize_sheet_name("Dr./John:Smith*Dept")` | Contains invalid Excel characters | `"Dr._John_Smith_Dept"` |
| `sanitize_filename("Prof*John Doe?.xlsx")` | Invalid filename characters | `"Prof_John_Doe.xlsx"` |
| `extract_faculty_names("Dr. A & Dr. B")` | Two faculty separated by `&` | `["Dr. A", "Dr. B"]` |
| `extract_faculty_names("Prof. A / Prof. B")` | Faculty separated by `/` | `["Prof. A", "Prof. B"]` |
| `extract_faculty_names("Dr. X, Dr. Y, Dr. Z")` | Multiple names with commas | `["Dr. X", "Dr. Y", "Dr. Z"]` |
| `extract_faculty_names("Dr. A and Dr. B")` | Separated by “and” | `["Dr. A", "Dr. B"]` |
| `load_config()` (with missing `config.json`) | Default config loaded | Returns dict `{ 'hour_slots': 2, ... }` |
| `generate_course_color()` exhausted after 15 palette colors | Check next generated color | Random color code in hex (e.g., `"F0D3E4"`) |

---

### ✅ Notes
- These test cases validate core logic functions in **TT_gen.py**.
- File-based tests (e.g., `combined.csv`, `rooms.csv`) should be in `tests/test_inputs/`.
- Use `pytest -v` to execute all test cases automatically.
