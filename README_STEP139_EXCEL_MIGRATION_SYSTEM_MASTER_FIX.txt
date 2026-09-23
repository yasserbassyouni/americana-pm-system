STEP139 - Excel Migration -> System Master Fix

Problem fixed:
- Preview correctly detected legacy Annual machines/tasks, but Convert only created the Professional Excel file.
- The current Smart PM architecture uses pm_master.json for machines and pm_tasks.json for tasks, so the imported line remained empty in Machine/Task Management.

Fix:
- Annual migration now imports machine identities into pm_master.json for the selected line.
- Annual migration imports PM tasks into pm_tasks.json for the selected line/year.
- Existing tasks for only that same line/year are archived/replaced safely; other lines/years are untouched.
- The migrated PM week plan and PM codes are preserved in scheduleWeeks/scheduleCodes.
- The line is marked as no longer fresh after migration so System Master remains authoritative.
- Stable Task IDs are written into the Professional Annual workbook.
- Annual Excel is rebuilt from System Master after import, proving the saved machine/task structure is authoritative.
- Conversion result now reports both machine and System Master task counts.

Verification performed on a copy using the included legacy upload that previews as 26 machines / 135 tasks:
- Preview: 26 machines / 135 tasks
- Convert: 26 machines / 135 System Master tasks
- Rebuilt Annual: 26 machines / 135 tasks

No header/branding changes are included in this step.
