#!/usr/bin/env python3
"""Import Leona's timetable into Donetick as one chore per bullet item.

Dry-run by default. Use --apply with DONETICK_USERNAME and DONETICK_PASSWORD
to create missing chores through the Donetick HTTP API.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Literal
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from import_grouped import (  # noqa: E402
    CHORE_SPECS,
    DEFAULT_BASE_URL,
    DEFAULT_EXCEL,
    LABELS,
    SHEET_NAME,
    TIMEZONE,
    existing_imported_chores,
    get_profile,
    json_request,
    login,
    merged_value,
    next_due_iso,
    parse_hhmm,
    split_subtasks,
    unwrap_items,
    set_due_date,
)


PROJECT_NAME = "Leona Household Timetable - Itemized"
PROJECT_DESCRIPTION = "Imported itemized schedule from Leona's Time Table.xlsx"
IMPORT_PREFIX = "leona-v2-itemized"
EXPECTED_ITEMIZED_COUNT = 71

Cadence = Literal["block", "interval"]


@dataclass(frozen=True)
class ItemizedChore:
    import_key: str
    name: str
    days: list[str]
    due_time: str
    next_due_date: str
    source_range: str
    source_bullet: str
    labels: list[str]
    frequency_type: str
    frequency: int
    frequency_metadata: dict[str, Any]


SPECIAL_SPLITS: dict[str, list[dict[str, Any]]] = {
    "Clean the master bedroom: Reset the bed suit every day, clean the toilet every other day": [
        {"name": "Reset the bed suit", "cadence": "block"},
        {"name": "Clean the toilet", "cadence": "interval", "frequency": 2, "unit": "days"},
    ],
}


def slugify(value: str) -> str:
    normalized = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    return normalized[:60].strip("-") or "item"


def unique_labels(labels: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for label in labels:
        if label not in seen:
            seen.add(label)
            result.append(label)
    return result


def labels_for_item(block_labels: list[str], name: str) -> list[str]:
    labels = [label for label in block_labels if label != "Optional"]
    lower = name.lower()
    if any(marker in lower for marker in ("optional", "option ", "if needed", "as needed", "if baby", "if we ", "unless ", "if you finish")):
        labels.append("Optional")
    if any(marker in lower for marker in ("clean", "wash", "vacuum", "vaccum", "mop", "tidy", "fold", "trash", "bedsheet", "bed suit", "toilet")):
        labels.append("Cleaning")
    if any(marker in lower for marker in ("breakfast", "lunch", "dinner", "cook", "kitchen", "dish", "dishes", "bowl", "groceries", "ingredients")):
        labels.append("Cooking")
    if any(marker in lower for marker in ("baby", "uniform", "school", "bottle", "sterilize", "stroller", "high chair")):
        labels.append("Baby")
    if any(marker in lower for marker in ("miya", "rabbit", "pet")):
        labels.append("Pets")
    if any(marker in lower for marker in ("buy", "market", "ntuc", "grocery list", "groceries")):
        labels.append("Errands")
    if any(marker in lower for marker in ("sleep", "next day", "back to your room")):
        labels.append("Evening")
    return unique_labels(labels)


def next_interval_due_iso(days: list[str], hhmm: str, tz_name: str) -> str:
    """Use the next due date from the source block as the first interval due date."""
    return next_due_iso(days, hhmm, tz_name)


def interval_frequency_metadata(next_due_date: str, unit: str) -> dict[str, Any]:
    return {
        "time": next_due_date,
        "timezone": TIMEZONE,
        "unit": unit,
    }


def block_frequency_metadata(days: list[str], next_due_date: str) -> dict[str, Any]:
    return {
        "days": days,
        "time": next_due_date,
        "timezone": TIMEZONE,
        "weekPattern": "every_week",
    }


def expand_bullet(block_key: str, bullet_index: int, bullet: str, block_labels: list[str]) -> list[dict[str, Any]]:
    if bullet in SPECIAL_SPLITS:
        expanded = []
        for split_index, split in enumerate(SPECIAL_SPLITS[bullet], start=1):
            expanded.append(
                {
                    "name": split["name"],
                    "source_bullet": bullet,
                    "key_suffix": f"{bullet_index:02d}-{split_index:02d}-{slugify(split['name'])}",
                    "cadence": split["cadence"],
                    "frequency": split.get("frequency", 1),
                    "unit": split.get("unit"),
                    "labels": labels_for_item(block_labels, split["name"]),
                }
            )
        return expanded
    return [
        {
            "name": bullet,
            "source_bullet": bullet,
            "key_suffix": f"{bullet_index:02d}-{slugify(bullet)}",
            "cadence": "block",
            "frequency": 1,
            "unit": None,
            "labels": labels_for_item(block_labels, bullet),
        }
    ]


def build_itemized_preview(excel_path: Path) -> list[dict[str, Any]]:
    workbook = load_workbook(excel_path, data_only=True)
    if workbook.sheetnames != [SHEET_NAME]:
        raise RuntimeError(f"Expected exactly one sheet named {SHEET_NAME!r}; found {workbook.sheetnames!r}")
    ws = workbook[SHEET_NAME]

    preview: list[dict[str, Any]] = []
    for spec in CHORE_SPECS:
        raw_text = merged_value(ws, spec.value_cell)
        bullets = split_subtasks(raw_text)
        if not bullets:
            raise RuntimeError(f"No bullet items extracted for {spec.key} from {spec.value_cell}")

        block_slug = spec.key.removeprefix("leona-v2-")
        for bullet_index, bullet in enumerate(bullets, start=1):
            expanded_items = expand_bullet(block_slug, bullet_index, bullet, spec.labels)
            for expanded in expanded_items:
                import_key = f"{IMPORT_PREFIX}-{block_slug}-{expanded['key_suffix']}"
                if expanded["cadence"] == "interval":
                    next_due = next_interval_due_iso(spec.days, spec.due_time, TIMEZONE)
                    frequency_type = "interval"
                    frequency = int(expanded["frequency"])
                    frequency_metadata = interval_frequency_metadata(next_due, str(expanded["unit"]))
                else:
                    next_due = next_due_iso(spec.days, spec.due_time, TIMEZONE)
                    frequency_type = "days_of_the_week"
                    frequency = 1
                    frequency_metadata = block_frequency_metadata(spec.days, next_due)
                description = (
                    f"Imported itemized from Leona's Time Table.xlsx | {spec.source_range} | "
                    f"sourceBullet={expanded['source_bullet']} | importKey={import_key}"
                )
                preview.append(
                    {
                        "importKey": import_key,
                        "name": expanded["name"],
                        "days": spec.days,
                        "dueTime": spec.due_time,
                        "nextDueDate": next_due,
                        "sourceRange": spec.source_range,
                        "sourceBullet": expanded["source_bullet"],
                        "labels": expanded["labels"],
                        "frequencyType": frequency_type,
                        "frequency": frequency,
                        "frequencyMetadata": frequency_metadata,
                        "description": description,
                    }
                )
    return preview


def print_preview(preview: list[dict[str, Any]]) -> None:
    print(f"Workbook: {DEFAULT_EXCEL}")
    print(f"Sheet: {SHEET_NAME}")
    print(f"Project: {PROJECT_NAME}")
    print(f"Chores planned: {len(preview)}")
    print()
    for idx, item in enumerate(preview, start=1):
        day_text = ",".join(day[:3].title() for day in item["days"])
        labels = ", ".join(item["labels"])
        cadence = item["frequencyType"]
        if cadence == "interval":
            cadence = f"interval/{item['frequency']} {item['frequencyMetadata'].get('unit')}"
        print(f"{idx:02d}. {item['name']} [{day_text} {item['dueTime']} {cadence}]")
        print(f"    key: {item['importKey']}")
        print(f"    source: {item['sourceRange']}")
        print(f"    labels: {labels}")
        print()


def ensure_project(base_url: str, token: str) -> int:
    projects = unwrap_items(json_request("GET", base_url, "/api/v1/projects", token=token))
    for project in projects:
        if project.get("name") == PROJECT_NAME:
            return int(project["id"])
    created = json_request(
        "POST",
        base_url,
        "/api/v1/projects",
        token=token,
        payload={"name": PROJECT_NAME, "description": PROJECT_DESCRIPTION, "color": "#0EA5E9", "icon": "checklist"},
    )
    project = created.get("res") if isinstance(created, dict) else None
    if not project or "id" not in project:
        raise RuntimeError(f"Project creation did not return an ID: {created!r}")
    return int(project["id"])


def ensure_labels(base_url: str, token: str) -> dict[str, int]:
    existing = unwrap_items(json_request("GET", base_url, "/api/v1/labels", token=token))
    by_name = {str(label.get("name")): int(label["id"]) for label in existing if label.get("id") is not None}
    for name, color in LABELS.items():
        if name in by_name:
            continue
        created = json_request("POST", base_url, "/api/v1/labels", token=token, payload={"name": name, "color": color})
        label = created.get("res") if isinstance(created, dict) else None
        if not label or "id" not in label:
            raise RuntimeError(f"Label creation for {name!r} did not return an ID: {created!r}")
        by_name[name] = int(label["id"])
    return by_name


def create_chore(base_url: str, token: str, item: dict[str, Any], user_id: int, project_id: int, label_ids: dict[str, int]) -> int:
    payload = {
        "name": item["name"],
        "frequencyType": item["frequencyType"],
        "frequency": item["frequency"],
        "frequencyMetadata": item["frequencyMetadata"],
        "nextDueDate": item["nextDueDate"],
        "isRolling": False,
        "assignedTo": user_id,
        "assignees": [{"userId": user_id}],
        "assignStrategy": "keep_last_assigned",
        "notification": False,
        "notificationMetadata": None,
        "labelsV2": [{"id": label_ids[name]} for name in item["labels"]],
        "priority": 0,
        "description": item["description"],
        "subTasks": [],
        "requireApproval": False,
        "isPrivate": False,
        "projectId": project_id,
    }
    created = json_request("POST", base_url, "/api/v1/chores/", token=token, payload=payload)
    if isinstance(created, dict) and "res" in created:
        chore_id = int(created["res"])
        set_due_date(base_url, token, chore_id, item["nextDueDate"])
        return chore_id
    raise RuntimeError(f"Chore creation for {item['name']!r} did not return an ID: {created!r}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Import Leona timetable as itemized Donetick chores.")
    parser.add_argument("--excel", default=DEFAULT_EXCEL)
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL)
    parser.add_argument("--apply", action="store_true", help="Actually create missing project/labels/chores in Donetick.")
    parser.add_argument("--username", default=os.environ.get("DONETICK_USERNAME"))
    parser.add_argument("--password", default=os.environ.get("DONETICK_PASSWORD"))
    parser.add_argument("--preview-json", default="")
    args = parser.parse_args()

    preview = build_itemized_preview(Path(args.excel))
    if len(preview) != EXPECTED_ITEMIZED_COUNT:
        raise RuntimeError(f"Expected {EXPECTED_ITEMIZED_COUNT} itemized chores; generated {len(preview)}")
    if args.preview_json:
        Path(args.preview_json).write_text(json.dumps(preview, indent=2, ensure_ascii=False), encoding="utf-8")
    print_preview(preview)

    if not args.apply:
        print("Dry run only. Re-run with --apply to create missing itemized chores.")
        return 0
    if not args.username or not args.password:
        raise RuntimeError("DONETICK_USERNAME and DONETICK_PASSWORD are required for --apply.")

    token = login(args.base_url, args.username, args.password)
    profile = get_profile(args.base_url, token)
    user_id = int(profile["id"])
    project_id = ensure_project(args.base_url, token)
    label_ids = ensure_labels(args.base_url, token)
    seen_chores = existing_imported_chores(args.base_url, token)

    created: list[tuple[str, int]] = []
    skipped: list[str] = []
    for item in preview:
        existing = seen_chores.get(item["importKey"])
        if existing:
            if not existing.get("nextDueDate"):
                set_due_date(args.base_url, token, int(existing["id"]), item["nextDueDate"])
            skipped.append(item["name"])
            continue
        chore_id = create_chore(args.base_url, token, item, user_id, project_id, label_ids)
        created.append((item["name"], chore_id))

    print()
    print(f"Created chores: {len(created)}")
    for name, chore_id in created:
        print(f"  + #{chore_id}: {name}")
    print(f"Skipped existing chores: {len(skipped)}")
    for name in skipped:
        print(f"  = {name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
