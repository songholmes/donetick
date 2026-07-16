#!/usr/bin/env python3
"""Mark the Leona itemized project as the only default project in its circle.

Dry-run by default. Use --apply while Donetick is stopped to update SQLite.
"""

from __future__ import annotations

import argparse
import sqlite3
from pathlib import Path
from typing import Any


DEFAULT_DB = Path("data/donetick.db")
DEFAULT_PROJECT_NAME = "Leona Household Timetable - Itemized"
EXPECTED_ITEMIZED_COUNT = 71
IMPORT_PREFIX = "leona-v2-itemized-"


def one_project_by_name(conn: sqlite3.Connection, project_name: str) -> dict[str, Any]:
    rows = conn.execute(
        """
        SELECT id, name, circle_id, created_by, is_default
        FROM projects
        WHERE name = ?
        """,
        (project_name,),
    ).fetchall()
    if len(rows) != 1:
        raise RuntimeError(f"Expected exactly one project named {project_name!r}; found {len(rows)}.")
    row = rows[0]
    return {
        "id": int(row["id"]),
        "name": str(row["name"]),
        "circle_id": int(row["circle_id"]),
        "created_by": int(row["created_by"]),
        "is_default": bool(row["is_default"]),
    }


def itemized_chore_count(conn: sqlite3.Connection, project_id: int) -> int:
    row = conn.execute(
        """
        SELECT COUNT(*) AS count
        FROM chores
        WHERE project_id = ?
          AND description LIKE ?
        """,
        (project_id, f"%importKey={IMPORT_PREFIX}%"),
    ).fetchone()
    return int(row["count"])


def circle_projects(conn: sqlite3.Connection, circle_id: int) -> list[dict[str, Any]]:
    rows = conn.execute(
        """
        SELECT id, name, is_default
        FROM projects
        WHERE circle_id = ?
        ORDER BY id
        """,
        (circle_id,),
    ).fetchall()
    return [
        {"id": int(row["id"]), "name": str(row["name"]), "is_default": bool(row["is_default"])}
        for row in rows
    ]


def set_default_project(conn: sqlite3.Connection, project_id: int, circle_id: int) -> None:
    with conn:
        conn.execute("UPDATE projects SET is_default = 0 WHERE circle_id = ?", (circle_id,))
        conn.execute("UPDATE projects SET is_default = 1 WHERE id = ? AND circle_id = ?", (project_id, circle_id))


def main() -> int:
    parser = argparse.ArgumentParser(description="Set Leona itemized project as Donetick's default project.")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB)
    parser.add_argument("--project-name", default=DEFAULT_PROJECT_NAME)
    parser.add_argument("--apply", action="store_true", help="Actually update SQLite. Default is dry-run.")
    parser.add_argument(
        "--skip-count-check",
        action="store_true",
        help=f"Do not require exactly {EXPECTED_ITEMIZED_COUNT} itemized chores in the target project.",
    )
    args = parser.parse_args()

    db_path = args.db.resolve()
    if not db_path.exists():
        raise RuntimeError(f"SQLite database not found: {db_path}")

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        project = one_project_by_name(conn, args.project_name)
        count = itemized_chore_count(conn, project["id"])
        projects_before = circle_projects(conn, project["circle_id"])

        if not args.skip_count_check and count != EXPECTED_ITEMIZED_COUNT:
            raise RuntimeError(
                f"Expected {EXPECTED_ITEMIZED_COUNT} itemized chores in project #{project['id']}; found {count}."
            )

        print(f"Database: {db_path}")
        print(f"Target project: #{project['id']} {project['name']!r}")
        print(f"Circle: #{project['circle_id']}")
        print(f"Current target isDefault: {project['is_default']}")
        print(f"Itemized chore count: {count}")
        print("Circle projects before:")
        for circle_project in projects_before:
            print(f"  - #{circle_project['id']} {circle_project['name']!r}: isDefault={circle_project['is_default']}")

        if not args.apply:
            print("Dry run only. Re-run with --apply to set this project as default.")
            return 0

        set_default_project(conn, project["id"], project["circle_id"])
        projects_after = circle_projects(conn, project["circle_id"])
        print("Circle projects after:")
        for circle_project in projects_after:
            print(f"  - #{circle_project['id']} {circle_project['name']!r}: isDefault={circle_project['is_default']}")
        return 0
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
