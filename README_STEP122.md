# Step122 — System Master review and repair

Based on `PM_Web_Server_Step121_DETERMINISTIC_MONTHLY_RENDER.zip` supplied by the user.

## Upgrade an existing installation

1. Stop the old server and take a copy of its complete folder, including `data`.
2. Replace `server.js`, `public`, `package.json`, `package-lock.json`, and `start.bat` with this release. Keep the existing installation's `data` directory. The data bundled in this ZIP is the original Step121 sample data, not your live data.
3. Start with `start.bat`. For a fresh installation, Node.js must already be installed; the launcher installs locked dependencies on first use.
4. Open Task Management for each existing line/year to perform its one-time legacy import if no system task master has been initialized. Rebuild Annual and the desired Monthly report through the existing rebuild endpoints/workflow. Keep historical months unless a rebuild is intended.

## What changed

- Annual machine/task structure now comes from `pm_master.json` and `pm_tasks.json` for the reviewed CRUD routes. Machine edits, renames, task edits, and combined deletion no longer edit/splice Excel rows as a database operation.
- Monthly local tasks, overrides, exclusions, and execution live in `data/monthly_state.json`. A missing Monthly Excel can be regenerated with its saved Done/User/Date. Existing legacy Monthly files are imported once; ambiguous missing task links fail explicitly instead of guessing.
- Initialization is tracked even when a line/year has zero tasks. Deleting the last task does not trigger another import of stale Annual rows. Invalid task JSON raises an error instead of silently falling back to Excel.
- Blank recreation of a deleted line starts with no machines/tasks. Retained Monthly folders are moved into `data/monthly_archive` when a new blank incarnation is created; files remain available there as history.
- Stable task IDs survive rename, edit, regeneration, and row shifts. Copying an existing line assigns independent IDs. Browser edits/deletions pass the displayed ID and stale row/ID combinations are rejected.
- Annual machine/task totals are written from the rendered system catalogue, including machines with no tasks. A:E merges are built only from complete machine blocks; both workbooks are checked before saving.
- Fixed missing `loadTaskPlanningMeta` in legacy bootstrap, lost PM code during migration, machine ownership after rename, undefined `name` in line edit, and monthly-only edits/deletions disappearing after regeneration.
- Machine identity in the Monthly task form is read-only; select a registered machine and edit its task data. Machine Management remains the identity editor.
- JSON writes use atomic replacement; API requests are serialized so concurrent mutations cannot overwrite each other's task changes. Added isolated runtime settings `PM_DATA_ROOT` and `PM_TEST_MODE` for repeatable tests.
- Fixed first-run Windows launcher control flow with `call npm ci` and an installation error check.

## Scope and behaviour

Annual-only and Monthly-only actions retain their chosen scope. A deliberate Monthly rebuild derives its base tasks from Annual system data while retaining stored monthly overrides and execution. Historical reports are not globally rewritten after each edit. Archived files are available on disk; this release does not add a new archive browsing page.

The recurrence generator, PM colour palette, and Done display function were compared against Step121 and are unchanged. Existing Done/Undo role/date rules remain; only persistence and stale-task checks were added. No email was sent during testing.

## Validation

Run `npm test` after installing dependencies. Tests create isolated temporary data and launch a real HTTP server; they do not mutate the bundled `data` folder. Set `TEST_PORT` if the default test port 13050 is occupied.

Passed: fresh test line; C1/C2 with zero and multiple tasks; Annual/Monthly A:E, IDs, task pairs and counts after saved XLSX reload; deletion of one and last machine task; final system task deletion; Done/User/Date across row shifts and regeneration; Undo; report removal and JSON-only recovery; monthly-only add/edit/delete; machine/line rename; zero-task machine deletion/re-add; task activation with stable IDs; stale ID rejection; delete/recreate line with retained Excel; restart; concurrent additions; identical task descriptions; invalid JSON rejection; migration of the supplied Cookies file (133 task rows, including one with extra whitespace); copied-line ID independence.

All server and inline browser scripts passed syntax checks. Browser checks covered login, Machine Management, Task Management editing, and Monthly Task Management editing, with no browser console errors observed. XLSX verification used ExcelJS reopening and cell/merge assertions; Microsoft Excel desktop recalculation/printing and external email delivery were not tested.
