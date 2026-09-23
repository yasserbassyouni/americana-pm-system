STEP140 - Safer Lines / Machines Management

Changes:
1. Lines / Machines list is now collapsed by default (accordion style).
2. Open one production line only when you need to work on it.
3. Machine Edit / Deactivate / Delete buttons are hidden behind a Manage menu.
4. Line controls are inside the opened line. Delete is separated at the far side as a danger action.
5. Line deletion now has three clear choices:
   - Remove from System Only: keep Annual, Monthly and folder.
   - Delete Annual + Monthly Excel: delete Excel files but keep line folder.
   - Delete Line Completely: delete system data, Annual, Monthly and line folder.
6. Complete deletion requires typing the exact line name as a second confirmation.
7. Monthly-file detection now scans line folders recursively, so usage/delete counts work with the current folder storage structure.

Based on STEP139 migration fix. Existing PM/Excel logic is otherwise unchanged.
