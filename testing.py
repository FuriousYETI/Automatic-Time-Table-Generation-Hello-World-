import unittest
import pandas as pd

import TT_gen


class TestTTGen(unittest.TestCase):
    def _empty_tt(self):
        data = {slot: [""] * len(TT_gen.days) for slot in TT_gen.slot_keys}
        return pd.DataFrame(data, index=TT_gen.days)

    def test_normalize_time_and_t2m(self):
        self.assertEqual(TT_gen.normalize_time("7:5"), "07:05")
        self.assertEqual(TT_gen.t2m("01:30"), 90)

    def test_shorten_faculty_name(self):
        name = "Dr.ABC XYZ / Prof. Jane Doe / Mr. Bob A"
        shortened = TT_gen.shorten_faculty_name(name)
        self.assertEqual(shortened, "Dr. ABC X / Prof. Jane D / Mr. Bob A")

    def test_split_faculty_names(self):
        parts = TT_gen.split_faculty_names("A / B /C")
        self.assertEqual(parts, ["A", "B", "C"])

    def test_classify_slot_val(self):
        self.assertEqual(TT_gen._classify_slot_val("CS101", "CS101 LAB"), "P")
        self.assertEqual(TT_gen._classify_slot_val("CS101", "CS101T1"), "T")
        self.assertEqual(TT_gen._classify_slot_val("CS101", "CS101"), "L")

    def test_extract_course_code(self):
        self.assertEqual(TT_gen.extract_course_code("CS101T (C3)"), "CS101")
        self.assertEqual(TT_gen.extract_course_code("CS102 (C4)"), "CS102")

    def test_build_room_map_from_tt(self):
        tt = self._empty_tt()
        first_slot = TT_gen.slot_keys[0]
        second_slot = TT_gen.slot_keys[1]
        tt.at["Monday", first_slot] = "CS101 (C3)"
        tt.at["Tuesday", second_slot] = "CS102 (C4)"
        room_map = TT_gen.build_room_map_from_tt(tt)
        self.assertEqual(room_map.get("CS101"), {"C3"})
        self.assertEqual(room_map.get("CS102"), {"C4"})

    def test_collect_code_slot_blocks(self):
        tt = self._empty_tt()
        s0, s1, s2, s3 = TT_gen.slot_keys[:4]
        tt.at["Monday", s0] = "CS101"
        tt.at["Monday", s1] = "CS101"
        tt.at["Monday", s3] = "CS101"
        blocks = TT_gen.collect_code_slot_blocks(tt, "CS101")
        self.assertIn(("Monday", [s0, s1]), blocks)
        self.assertIn(("Monday", [s3]), blocks)

    def test_collect_unscheduled(self):
        courses = [
            {
                "Departments": "CSE",
                "Semester": "1",
                "Section": "A",
                "Course_Code": "CS101",
                "Course_Title": "Intro",
                "Faculty": "Dr. A",
                "L-T-P-S-C": "3-0-0-0-3",
                "Elective": "0",
                "ElectiveBasket": "0",
                "Semester_Half": "1",
            },
            {
                "Departments": "CSE",
                "Semester": "1",
                "Section": "A",
                "Course_Code": "CS102",
                "Course_Title": "Elective",
                "Faculty": "Dr. B",
                "L-T-P-S-C": "3-0-0-0-3",
                "Elective": "1",
                "ElectiveBasket": "2",
                "Semester_Half": "1",
            },
            {
                "Departments": "CSE",
                "Semester": "1",
                "Section": "A",
                "Course_Code": "CS103",
                "Course_Title": "Data",
                "Faculty": "Dr. C",
                "L-T-P-S-C": "3-0-0-0-3",
                "Elective": "0",
                "ElectiveBasket": "0",
                "Semester_Half": "1",
            },
        ]
        placed = [courses[0]]
        elective_sync = {"Y1_B2": True}
        uns = TT_gen.collect_unscheduled(courses, placed, "TestGroup", year_tag=1, elective_sync=elective_sync)
        self.assertEqual([u["Course_Code"] for u in uns], ["CS103"])


if __name__ == "__main__":
    unittest.main()
