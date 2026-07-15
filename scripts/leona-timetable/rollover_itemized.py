#!/usr/bin/env python3
"""Advance stale itemized timetable chores to their next valid schedule.

The default mode is a dry run. ``--once`` applies one reconciliation and
``--daemon`` reconciles immediately, then every day at 00:05 Singapore time.
Authentication uses a Donetick API token read from a file.
"""

from __future__ import annotations

import argparse
import json
import re
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime, time as clock_time, timedelta, timezone
from pathlib import Path
from typing import Any


DEFAULT_BASE_URL = "http://localhost:2021"
DEFAULT_TOKEN_FILE = (
    Path(__file__).resolve().parent
    / ".."
    / ".."
    / "data"
    / "secrets"
    / "donetick_api_token"
).resolve()
PROJECT_ID = 2
PROJECT_NAME = "Leona Household Timetable - Itemized"
IMPORT_PREFIX = "leona-v2-itemized-"
EXPECTED_COUNT = 71
SINGAPORE = timezone(timedelta(hours=8), "Asia/Singapore")
DEFAULT_RUN_AT = "00:05"


class APIError(RuntimeError):
    def __init__(self, status: int, method: str, path: str, body: str):
        super().__init__(f"{method} {path} failed with HTTP {status}: {body}")
        self.status = status


class APIClient:
    def __init__(self, base_url: str, api_token: str):
        self.base_url = base_url.rstrip("/")
        self.api_token = api_token

    def request(self, method: str, path: str, payload: Any | None = None) -> Any:
        headers = {"Content-Type": "application/json", "secretkey": self.api_token}
        data = json.dumps(payload).encode("utf-8") if payload is not None else None
        request = urllib.request.Request(
            self.base_url + path,
            data=data,
            headers=headers,
            method=method,
        )
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                body = response.read().decode("utf-8")
                return json.loads(body) if body else None
        except urllib.error.HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            raise APIError(exc.code, method, path, body) from exc

    def projects(self) -> list[dict[str, Any]]:
        return unwrap_items(self.request("GET", "/api/v1/projects"))

    def chores(self) -> list[dict[str, Any]]:
        return unwrap_items(
            self.request("GET", "/api/v1/chores/?includeArchived=true")
        )

    def chore(self, chore_id: int) -> dict[str, Any]:
        response = self.request("GET", f"/api/v1/chores/{chore_id}")
        if isinstance(response, dict):
            value = response.get("res", response)
            if isinstance(value, dict):
                return value
        raise RuntimeError(f"Could not read chore #{chore_id}.")

    def update_due_date(self, chore_id: int, due_date: datetime, updated_at: str) -> None:
        self.request(
            "PUT",
            f"/api/v1/chores/{chore_id}/dueDate",
            payload={
                "dueDate": due_date.isoformat(timespec="seconds"),
                "updatedAt": updated_at,
            },
        )


@dataclass(frozen=True)
class RolloverChange:
    chore: dict[str, Any]
    old_due: datetime
    new_due: datetime


def unwrap_items(response: Any) -> list[dict[str, Any]]:
    if isinstance(response, list):
        return response
    if isinstance(response, dict):
        for key in ("res", "data", "items"):
            value = response.get(key)
            if isinstance(value, list):
                return value
    raise RuntimeError(f"Could not unwrap list response: {response!r}")


def parse_datetime(value: str) -> datetime:
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        raise ValueError(f"Datetime must include a timezone: {value!r}")
    return parsed


def import_key(description: str | None) -> str | None:
    match = re.search(r"importKey=([^\s|]+)", description or "")
    return match.group(1) if match else None


def scheduled_time(metadata: dict[str, Any]) -> clock_time:
    value = metadata.get("time")
    if not isinstance(value, str) or not value:
        raise ValueError("Frequency metadata does not contain a schedule time.")
    local = parse_datetime(value).astimezone(SINGAPORE)
    return clock_time(local.hour, local.minute, local.second)


def next_days_of_week_due(
    metadata: dict[str, Any], today: date
) -> datetime:
    raw_days = metadata.get("days")
    if not isinstance(raw_days, list) or not raw_days:
        raise ValueError("days_of_the_week chore has no configured days.")
    days = {str(day).lower() for day in raw_days}
    due_time = scheduled_time(metadata)
    for offset in range(0, 8):
        candidate = today + timedelta(days=offset)
        if candidate.strftime("%A").lower() in days:
            return datetime.combine(candidate, due_time, SINGAPORE)
    raise ValueError("Could not find a matching weekday in the next seven days.")


def next_interval_due(
    chore: dict[str, Any], old_due: datetime, today: date
) -> datetime:
    metadata = chore.get("frequencyMetadata") or {}
    frequency = int(chore.get("frequency") or 0)
    unit = metadata.get("unit")
    if frequency <= 0:
        raise ValueError("Interval frequency must be positive.")
    if unit == "days":
        step = timedelta(days=frequency)
    elif unit == "weeks":
        step = timedelta(weeks=frequency)
    else:
        raise ValueError(f"Unsupported interval unit for rollover: {unit!r}")
    candidate = old_due.astimezone(SINGAPORE)
    for _ in range(0, 10000):
        if candidate.date() >= today:
            return candidate
        candidate += step
    raise ValueError("Interval rollover exceeded its safety iteration limit.")


def calculate_rollover_due(
    chore: dict[str, Any], now: datetime
) -> datetime | None:
    raw_due = chore.get("nextDueDate")
    if not isinstance(raw_due, str) or not raw_due:
        raise ValueError(f"Chore #{chore.get('id')} does not have a next due date.")
    old_due = parse_datetime(raw_due)
    local_now = now.astimezone(SINGAPORE)
    if old_due.astimezone(SINGAPORE).date() >= local_now.date():
        return None

    frequency_type = chore.get("frequencyType")
    metadata = chore.get("frequencyMetadata") or {}
    if frequency_type == "days_of_the_week":
        return next_days_of_week_due(metadata, local_now.date())
    if frequency_type == "interval":
        return next_interval_due(chore, old_due, local_now.date())
    raise ValueError(
        f"Unsupported frequency type for itemized rollover: {frequency_type!r}"
    )


def validate_itemized_scope(
    projects: list[dict[str, Any]], chores: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    project = [
        project
        for project in projects
        if int(project.get("id", -1)) == PROJECT_ID
        and project.get("name") == PROJECT_NAME
    ]
    if len(project) != 1:
        raise RuntimeError(
            f"Expected project #{PROJECT_ID} named {PROJECT_NAME!r}; found {len(project)}."
        )
    itemized = [chore for chore in chores if int(chore.get("projectId", -1)) == PROJECT_ID]
    keys = [import_key(chore.get("description")) for chore in itemized]
    if (
        len(itemized) != EXPECTED_COUNT
        or len(set(keys)) != EXPECTED_COUNT
        or any(not key or not key.startswith(IMPORT_PREFIX) for key in keys)
    ):
        raise RuntimeError(
            "Itemized safety check failed: expected 71 chores with 71 unique "
            f"{IMPORT_PREFIX!r} keys; found chores={len(itemized)}, "
            f"uniqueKeys={len(set(keys))}."
        )
    return itemized


def build_rollover_plan(
    projects: list[dict[str, Any]],
    chores: list[dict[str, Any]],
    now: datetime,
) -> tuple[list[RolloverChange], int]:
    itemized = validate_itemized_scope(projects, chores)
    changes: list[RolloverChange] = []
    effective_dates: list[date] = []
    today = now.astimezone(SINGAPORE).date()
    for chore in itemized:
        old_due = parse_datetime(str(chore["nextDueDate"]))
        new_due = calculate_rollover_due(chore, now)
        if new_due is not None:
            changes.append(RolloverChange(chore, old_due, new_due))
            effective_dates.append(new_due.astimezone(SINGAPORE).date())
        else:
            effective_dates.append(old_due.astimezone(SINGAPORE).date())
    today_count = sum(due_date == today for due_date in effective_dates)
    return changes, today_count


def update_with_retry(
    client: APIClient, change: RolloverChange
) -> None:
    chore_id = int(change.chore["id"])
    updated_at = change.chore.get("updatedAt")
    if not isinstance(updated_at, str) or not updated_at:
        raise RuntimeError(f"Chore #{chore_id} does not have updatedAt.")
    try:
        client.update_due_date(chore_id, change.new_due, updated_at)
    except APIError as exc:
        if exc.status not in (403, 409):
            raise
        fresh = client.chore(chore_id)
        fresh_updated_at = fresh.get("updatedAt")
        if not isinstance(fresh_updated_at, str) or not fresh_updated_at:
            raise RuntimeError(f"Refetched chore #{chore_id} does not have updatedAt.")
        client.update_due_date(chore_id, change.new_due, fresh_updated_at)


def reconcile(client: APIClient, apply: bool, now: datetime | None = None) -> int:
    effective_now = now or datetime.now(SINGAPORE)
    changes, today_count = build_rollover_plan(
        client.projects(), client.chores(), effective_now
    )
    print(
        f"Safety check: {EXPECTED_COUNT} itemized chores; "
        f"planned changes={len(changes)}; due today after reconciliation={today_count}",
        flush=True,
    )
    for change in changes:
        print(
            f"  #{change.chore['id']} {change.chore['name']}: "
            f"{change.old_due.astimezone(SINGAPORE).isoformat(timespec='minutes')} -> "
            f"{change.new_due.astimezone(SINGAPORE).isoformat(timespec='minutes')}",
            flush=True,
        )
    if not apply:
        print("Dry run only; no due dates were changed.", flush=True)
        return len(changes)

    failures: list[str] = []
    for change in changes:
        try:
            update_with_retry(client, change)
            time.sleep(0.05)
        except Exception as exc:  # continue so a rerun can repair only the remainder
            failures.append(f"#{change.chore['id']}: {exc}")
    if failures:
        raise RuntimeError("Rollover failures: " + "; ".join(failures))
    print(f"Applied {len(changes)} due-date changes.", flush=True)
    return len(changes)


def read_api_token(path: Path) -> str:
    if not path.is_file():
        raise RuntimeError(f"API token file does not exist: {path}")
    token = path.read_text(encoding="utf-8").strip()
    if not token:
        raise RuntimeError(f"API token file is empty: {path}")
    return token


def parse_run_at(value: str) -> clock_time:
    try:
        return datetime.strptime(value, "%H:%M").time()
    except ValueError as exc:
        raise argparse.ArgumentTypeError("run-at must use HH:MM") from exc


def seconds_until_next_run(now: datetime, run_at: clock_time) -> float:
    local_now = now.astimezone(SINGAPORE)
    target = datetime.combine(local_now.date(), run_at, SINGAPORE)
    if target <= local_now:
        target += timedelta(days=1)
    return max(1.0, (target - local_now).total_seconds())


def write_health(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(datetime.now(SINGAPORE).isoformat(), encoding="utf-8")


def daemon_loop(
    client: APIClient, health_file: Path, run_at: clock_time
) -> None:
    while True:
        try:
            reconcile(client, apply=True)
            write_health(health_file)
        except Exception as exc:
            health_file.unlink(missing_ok=True)
            print(f"Rollover failed; retrying in 60 seconds: {exc}", flush=True)
            time.sleep(60)
            continue
        delay = seconds_until_next_run(datetime.now(SINGAPORE), run_at)
        print(f"Next rollover in {int(delay)} seconds.", flush=True)
        time.sleep(delay)


def main() -> int:
    parser = argparse.ArgumentParser(description="Reconcile Leona itemized due dates.")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--dry-run", action="store_true")
    mode.add_argument("--once", action="store_true")
    mode.add_argument("--daemon", action="store_true")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL)
    parser.add_argument("--api-token-file", type=Path, default=DEFAULT_TOKEN_FILE)
    parser.add_argument("--health-file", type=Path, default=Path("/tmp/rollover-ready"))
    parser.add_argument("--run-at", type=parse_run_at, default=parse_run_at(DEFAULT_RUN_AT))
    args = parser.parse_args()

    client = APIClient(args.base_url, read_api_token(args.api_token_file))
    if args.daemon:
        daemon_loop(client, args.health_file, args.run_at)
        return 0
    reconcile(client, apply=args.once)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
