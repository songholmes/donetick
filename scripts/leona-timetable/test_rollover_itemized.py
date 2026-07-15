from __future__ import annotations

import unittest
from datetime import date, datetime

from rollover_itemized import (
    APIError,
    EXPECTED_COUNT,
    IMPORT_PREFIX,
    PROJECT_ID,
    PROJECT_NAME,
    RolloverChange,
    SINGAPORE,
    calculate_rollover_due,
    next_days_of_week_due,
    update_with_retry,
    validate_itemized_scope,
)


NOW = datetime(2026, 7, 14, 12, 0, tzinfo=SINGAPORE)  # Tuesday


def weekday_chore(due: str, days: list[str] | None = None) -> dict:
    return {
        "id": 10,
        "name": "Test task",
        "projectId": PROJECT_ID,
        "description": f"importKey={IMPORT_PREFIX}test",
        "nextDueDate": due,
        "frequencyType": "days_of_the_week",
        "frequency": 1,
        "frequencyMetadata": {
            "days": days or ["tuesday"],
            "time": "2026-07-14T08:00:00+08:00",
            "timezone": "Asia/Singapore",
        },
        "updatedAt": "2026-07-14T00:00:00Z",
    }


class SchedulingTests(unittest.TestCase):
    def test_stale_weekday_advances_to_today(self) -> None:
        chore = weekday_chore("2026-07-07T08:00:00+08:00")
        actual = calculate_rollover_due(chore, NOW)
        self.assertEqual(actual, datetime(2026, 7, 14, 8, 0, tzinfo=SINGAPORE))

    def test_stale_weekend_advances_across_week(self) -> None:
        actual = next_days_of_week_due(
            weekday_chore("2026-07-11T08:00:00+08:00", ["saturday"])[
                "frequencyMetadata"
            ],
            date(2026, 7, 14),
        )
        self.assertEqual(actual, datetime(2026, 7, 18, 8, 0, tzinfo=SINGAPORE))

    def test_two_day_interval_preserves_phase(self) -> None:
        chore = weekday_chore("2026-07-13T08:00:00+08:00")
        chore["frequencyType"] = "interval"
        chore["frequency"] = 2
        chore["frequencyMetadata"] = {
            "unit": "days",
            "time": "2026-07-13T08:00:00+08:00",
            "timezone": "Asia/Singapore",
        }
        actual = calculate_rollover_due(chore, NOW)
        self.assertEqual(actual, datetime(2026, 7, 15, 8, 0, tzinfo=SINGAPORE))

    def test_today_is_unchanged(self) -> None:
        chore = weekday_chore("2026-07-14T08:00:00+08:00")
        self.assertIsNone(calculate_rollover_due(chore, NOW))

    def test_future_is_unchanged(self) -> None:
        chore = weekday_chore("2026-07-21T08:00:00+08:00")
        self.assertIsNone(calculate_rollover_due(chore, NOW))

    def test_invalid_weekday_metadata_stops(self) -> None:
        chore = weekday_chore("2026-07-07T08:00:00+08:00")
        chore["frequencyMetadata"]["days"] = []
        with self.assertRaisesRegex(ValueError, "configured days"):
            calculate_rollover_due(chore, NOW)


class SafetyTests(unittest.TestCase):
    def test_exact_itemized_scope_is_accepted(self) -> None:
        projects = [{"id": PROJECT_ID, "name": PROJECT_NAME}]
        chores = []
        for index in range(EXPECTED_COUNT):
            chore = weekday_chore("2026-07-14T08:00:00+08:00")
            chore["id"] = index + 1
            chore["description"] = f"importKey={IMPORT_PREFIX}{index:03d}"
            chores.append(chore)
        self.assertEqual(validate_itemized_scope(projects, chores), chores)

    def test_duplicate_import_key_stops(self) -> None:
        projects = [{"id": PROJECT_ID, "name": PROJECT_NAME}]
        chores = []
        for index in range(EXPECTED_COUNT):
            chore = weekday_chore("2026-07-14T08:00:00+08:00")
            chore["id"] = index + 1
            chore["description"] = f"importKey={IMPORT_PREFIX}duplicate"
            chores.append(chore)
        with self.assertRaisesRegex(RuntimeError, "safety check failed"):
            validate_itemized_scope(projects, chores)


class RetryTests(unittest.TestCase):
    def test_concurrent_update_refetches_and_retries_once(self) -> None:
        class FakeClient:
            def __init__(self) -> None:
                self.updated_at_values: list[str] = []
                self.refetches = 0

            def update_due_date(
                self, chore_id: int, due_date: datetime, updated_at: str
            ) -> None:
                self.updated_at_values.append(updated_at)
                if len(self.updated_at_values) == 1:
                    raise APIError(403, "PUT", f"/chores/{chore_id}/dueDate", "")

            def chore(self, chore_id: int) -> dict:
                self.refetches += 1
                return {"id": chore_id, "updatedAt": "2026-07-14T01:00:00Z"}

        chore = weekday_chore("2026-07-07T08:00:00+08:00")
        change = RolloverChange(
            chore=chore,
            old_due=datetime(2026, 7, 7, 8, 0, tzinfo=SINGAPORE),
            new_due=datetime(2026, 7, 14, 8, 0, tzinfo=SINGAPORE),
        )
        client = FakeClient()
        update_with_retry(client, change)  # type: ignore[arg-type]
        self.assertEqual(client.refetches, 1)
        self.assertEqual(
            client.updated_at_values,
            ["2026-07-14T00:00:00Z", "2026-07-14T01:00:00Z"],
        )


if __name__ == "__main__":
    unittest.main()
