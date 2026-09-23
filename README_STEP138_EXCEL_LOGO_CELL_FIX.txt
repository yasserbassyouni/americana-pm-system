STEP138 - Excel Logo Cell Auto-Fit definitive fix

Root cause:
The Annual and Monthly formatting functions were called again AFTER branding and reset Row 1 and columns A:C to fixed template sizes. That is why the logo image changed but the Excel cell did not.

Fix:
- Re-apply Excel logo cell auto-fit as the final step of every Annual/Monthly formatting pass.
- Row 1 height follows Excel logo height + padding.
- Columns A:C grow when needed for Excel logo width + padding.
- Branding API now returns Excel logo settings consistently after Save.
- The fix survives Apply Branding to Existing Excel Files and later machine/task edits that reformat Excel.
