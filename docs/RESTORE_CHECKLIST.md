# Restore checklist

Taken from `OPS_RESTORE_CHECKLIST` in `Ops_Daily.gs`, which Operations → Backups also shows. If you change one, change the other.

1. Decide what you are restoring: ONE sheet (a bad edit, a deleted row) or EVERYTHING (the spreadsheet is lost or corrupted). One sheet is almost always the answer.
2. Stop the damage first: switch the web app to "Only myself" (Deploy → Manage deployments) so nobody writes while you restore. Tell the desks.
3. Pick the backup: Admin Dashboard → Operations → Backups lists them. Choose the newest one from BEFORE the problem. Check its "Verified" column says YES.
4. ONE SHEET: open the backup, right-click the sheet tab → Copy to → the live spreadsheet. In the live file, rename the damaged sheet to "<name>_damaged", rename the copy to the exact original name, and drag it to the same position. Formulas and other modules find sheets by name.
5. Rows added after the backup: compare the "_damaged" sheet with the restored one and copy across any rows added since the backup time. Appointments, invoices and results added today are the usual ones.
6. EVERYTHING: make a copy of the backup (File → Make a copy), then point the deployment at it — open Extensions → Apps Script in the COPY, Deploy → New deployment, and update the link staff use. Script properties (API keys, clinic name, backup folder) do NOT travel with a copy: re-enter them from Project Settings of the old project.
7. Reinstall the scheduled jobs in the restored project: run opsInstallTriggers() and dpdpInstallTriggers() from the editor. Triggers do not copy.
8. Check: sign in as a test user, open a patient, open today's appointments, the lab desk and the pharmacy. Run verifyDeployment() from the editor.
9. Sessions: everyone signs in again — the session cache does not carry over. Passwords and two-step sign-in do (they are in the Users sheet).
10. Record it: add a row to Backup_Log saying who restored what, from which backup, and why. Under the DPDP Act a lost or corrupted patient record may be a personal data breach — check with the grievance officer whether it must be reported.
11. Re-open the web app to its normal audience, and delete the "_damaged" sheet only once you are sure nothing in it is still needed.
