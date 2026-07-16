from __future__ import annotations

import unittest

from openpyxl import Workbook

from import_itemized import due_time_from_source_range, labels_for_item, parse_time_range, resolve_assignee_user_id


class EndTimeParsingTests(unittest.TestCase):
    def test_single_row_uses_row_end_time(self) -> None:
        workbook = Workbook()
        ws = workbook.active
        ws["A6"] = "8:00 - 11:15"

        due_time, source_window = due_time_from_source_range(ws, "Version 2!B6:F6")

        self.assertEqual(due_time, "11:15")
        self.assertEqual(source_window, "08:00 - 11:15")

    def test_multi_row_source_uses_final_row_end_time(self) -> None:
        workbook = Workbook()
        ws = workbook.active
        ws["A6"] = "8:00 - 11:15"
        ws["A7"] = "11:15 - 11:45"
        ws["A8"] = "12:00 - 13:30"

        due_time, source_window = due_time_from_source_range(ws, "Version 2!G6:H8")

        self.assertEqual(due_time, "13:30")
        self.assertEqual(source_window, "08:00 - 13:30")

    def test_merged_final_time_cell_is_resolved(self) -> None:
        workbook = Workbook()
        ws = workbook.active
        ws["A9"] = "13:30 - 15:00"
        ws.merge_cells("A9:A10")

        due_time, source_window = due_time_from_source_range(ws, "Version 2!G9:H10")

        self.assertEqual(due_time, "15:00")
        self.assertEqual(source_window, "13:30 - 15:00")

    def test_time_range_normalizes_single_digit_hour(self) -> None:
        self.assertEqual(parse_time_range("6:00 - 7:00"), ("06:00", "07:00"))


class AssigneeResolutionTests(unittest.TestCase):
    def test_username_match_wins(self) -> None:
        members = [
            {"userId": 1, "username": "songholmes", "displayName": "Sir"},
            {"userId": 2, "username": "songholmes_leona", "displayName": "Leona"},
        ]

        self.assertEqual(resolve_assignee_user_id(members, "songholmes_leona"), 2)

    def test_display_name_match_is_supported(self) -> None:
        members = [
            {"userId": 1, "username": "songholmes", "displayName": "Sir"},
            {"userId": 2, "username": "songholmes_leona", "displayName": "Leona"},
        ]

        self.assertEqual(resolve_assignee_user_id(members, "Leona"), 2)

    def test_missing_assignee_fails_clearly(self) -> None:
        members = [{"userId": 1, "username": "songholmes", "displayName": "Sir"}]

        with self.assertRaisesRegex(RuntimeError, "was not found"):
            resolve_assignee_user_id(members, "songholmes_leona")


class ItemLabelTests(unittest.TestCase):
    def test_self_prep_does_not_inherit_task_categories(self) -> None:
        labels = labels_for_item(["Weekday", "Morning", "Baby", "Cooking", "Pets"], "Wake up and get yourself ready")

        self.assertEqual(labels, ["Weekday", "Morning"])

    def test_optional_breakfast_is_not_cooking(self) -> None:
        labels = labels_for_item(["Weekday", "Cleaning", "Optional"], "Have your breakfast (optional 2)")

        self.assertEqual(labels, ["Weekday", "Optional"])

    def test_cooking_and_kitchen_cleanup_are_specific(self) -> None:
        labels = labels_for_item(["Weekday", "Cooking", "Errands"], "Cook Dinner and tidy up the kitchen")

        self.assertEqual(labels, ["Weekday", "Cleaning", "Cooking"])

    def test_baby_play_does_not_become_cooking_from_dinner_word(self) -> None:
        labels = labels_for_item(
            ["Weekday", "Baby"],
            "Or play with baby in case Sir and Madam didn't finish dinner",
        )

        self.assertEqual(labels, ["Weekday", "Baby"])

    def test_dinner_dishes_do_not_inherit_baby_or_pets(self) -> None:
        labels = labels_for_item(
            ["Weekday", "Baby", "Cleaning", "Pets", "Evening"],
            "Tidy up the dished and bowls from dinner",
        )

        self.assertEqual(labels, ["Weekday", "Evening", "Cleaning"])

    def test_walk_miya_is_pets(self) -> None:
        labels = labels_for_item(["Weekend", "Optional"], "Walk Miya if you finish early.")

        self.assertEqual(labels, ["Weekend", "Optional", "Pets"])


if __name__ == "__main__":
    unittest.main()
