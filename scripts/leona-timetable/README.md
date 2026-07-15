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

> Historical only: do not run the grouped importer with `--apply` on the itemized-only installation. The module remains because the itemized importer reuses its workbook parsing, schedule definitions, and API helpers.

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

- Project URLs, such as `/chores?project=2`, show all chores in that project.
- Tasks whose due time has already passed today appear under `Overdue`.
- For the itemized setup, use `http://localhost:2021/chores?project=2` as the stable project view.
- Do not force `filterId=due-today`; Donetick's strict today filter can make the page look like it contains only overdue chores once scheduled times have passed.

## Itemized importer

`import_itemized.py` creates or reuses:

- Project: `Leona Household Timetable - Itemized`
- The same label set as the grouped importer
- One chore per timetable bullet item

Each itemized chore inherits the day pattern and due time from its timetable block. Import markers use the separate `leona-v2-itemized-...` namespace so itemized chores do not collide with grouped chores.

Dry run:

```powershell
python .\scripts\leona-timetable\import_itemized.py
```

Apply:

```powershell
python .\scripts\leona-timetable\import_itemized.py --apply
```

Special split rule:

- `Clean the master bedroom: Reset the bed suit every day, clean the toilet every other day`
  - `Reset the bed suit`: inherits the original weekday `08:00` schedule
  - `Clean the toilet`: uses Donetick interval scheduling every 2 days at `08:00`

## Stable itemized view

The Docker setup exposes Donetick directly on `0.0.0.0:2021`, without an Nginx gateway or injected default-view script. This keeps browser loading simple and avoids forcing the app into Donetick's strict `due-today` filter.

Recommended URLs:

```text
http://localhost:2021
http://localhost:2021/chores?project=2
http://<your-lan-ip>:2021/chores?project=2
```

## Optional daily itemized rollover

Donetick advances a recurring chore after it is completed or skipped, but it does not automatically move an unfinished occurrence out of the past. The optional itemized rollover reconciler moves only stale overdue itemized occurrences to today or the next valid schedule. It never changes a chore that is already due today or in the future.

If you want to use it, store a Donetick API token outside Git in `data/secrets/donetick_api_token`, then review the planned changes or apply one pass with:

```powershell
python .\scripts\leona-timetable\rollover_itemized.py --dry-run
python .\scripts\leona-timetable\rollover_itemized.py --once
```

The script validates the itemized project and all 71 unique import keys before it updates anything. It is not enabled in `docker-compose.yaml` by default so browser access remains independent of the rollover utility.
