# Leona timetable importers

These one-off importers read `Leona's Time Table.xlsx` and create recurring chores through the Donetick HTTP API. They do not edit the workbook and do not write directly to SQLite.

## Requirements

- Donetick running locally, default `http://localhost:2021`
- Workbook at `C:\Users\Songholmes\Downloads\Leona's Time Table.xlsx`
- Sheet name: `Version 2`
- Python with `openpyxl`
- Donetick credentials supplied outside Git:

```powershell
$env:DONETICK_USERNAME = "your-user"
$env:DONETICK_PASSWORD = "your-password"
```

## Grouped importer

`import_grouped.py` creates or reuses:

- Project: `Leona Household Timetable`
- Labels: `Weekday`, `Weekend`, `Morning`, `Baby`, `Cleaning`, `Cooking`, `Pets`, `Errands`, `Evening`, `Optional`
- 24 grouped recurring chores

Each chore corresponds to a timetable block, and the bullet lines inside that block become Donetick subtasks. Import markers are stored in the chore description as `importKey=...`, so reruns skip existing imports.

Dry run:

```powershell
python .\scripts\leona-timetable\import_grouped.py
```

Apply:

```powershell
python .\scripts\leona-timetable\import_grouped.py --apply
```

## Donetick behavior notes

- The project URL, such as `/chores?project=1`, shows all chores in that project.
- Tasks whose due time has already passed today appear under `Overdue`.
- A useful pinned filter for the grouped project is `project is Leona Household Timetable` plus `dueDate isOverdue`.
