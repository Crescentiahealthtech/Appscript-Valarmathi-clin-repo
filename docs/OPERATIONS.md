# Operations: reminders, stock alerts, day summary, backups, caching

Everything here is in **Admin Dashboard → Operations**. Reception also reaches
*Reminders* from Appointments & IP Routing, and the pharmacy reaches *Stock
alerts* from the Pharmacy Command Center. Each tab appears only for a role
with its permission.

## One-time setup (Apps Script editor)

1. Run `opsInstallTriggers()`. It installs three daily jobs, which run as your account:

   | Job | Time | What it does |
   |---|---|---|
   | `opsNightlyBackup` | ~01:00 | Copies the whole spreadsheet to a private Drive folder, checks the copy, and removes old copies |
   | `remindersDaily` | ~18:00 | Sends tomorrow's reminders, if a provider is set up (see below) |
   | `ownerDailySummary` | ~21:00 | Emails the owner the day's figures |

2. In **Project Settings → Script properties**, set any of these you need:

   | Property | Meaning |
   |---|---|
   | `OWNER_EMAIL` | Who gets the day summary. If you leave it out, it goes to the email on every `Super_Admin` Users row. |
   | `BACKUP_FOLDER_ID` | Optional. If you leave it out, a private folder called "CresRx Backups (restricted)" is created on the first run. |
   | `BACKUP_KEEP_DAYS` | Daily copies to keep (default 30). The first copy of each month is kept for a year. |
   | `STOCK_MIN_UNITS` | The low-stock floor for a medicine with no recent sales (default 10). |
   | `REMINDER_PROVIDER` | Leave it blank to send from the desk. Set it to `WHATSAPP_CLOUD` or `SMS_HTTP` to send automatically. |
   | `WA_PHONE_NUMBER_ID`, `WA_ACCESS_TOKEN`, `WA_TEMPLATE_NAME`, `WA_TEMPLATE_LANG` | WhatsApp Business Cloud API. The template must be approved and take four body parameters: `{{1}}` name, `{{2}}` what, `{{3}}` when, `{{4}}` clinic. |
   | `SMS_URL_TEMPLATE` | An HTTP GET URL containing `{to}` and `{text}`. Indian SMS also needs the message text registered as a DLT template. |

3. The first run asks you to re-authorise, because it adds Drive and external-request (UrlFetch) permissions.

## Reminders

- **Appointments**: sent for appointments that are still *Booked*.
- **Follow-ups**: sent when `OP_Encounters.Next_Review_Date` is that day and the patient has no appointment booked for it.
- **Vaccinations**: sent when a child's next dose on the immunisation card (Maternal & Child) falls due that day.
- **Consent**: only patients whose DPDP `COMMUNICATION` consent is *Given* are messaged. The list shows everyone else with the reason, so the desk can ask them at their next visit.
- **Without a provider**: each row has WhatsApp and SMS buttons that open the message ready to send from the clinic phone. Once a row is opened, it is recorded and won't be offered again.
- **Record**: every reminder, sent or failed, is logged in `Reminder_Log`.

## Stock alerts

- **Low stock** is measured in days: current stock divided by the average daily sale over the last 60 days (from `Pharmacy_Invoice_Items`). An item is low when it has fewer than 14 days left.
- **Items with no recent sales** use the plain floor (`STOCK_MIN_UNITS`).
- **Per-medicine levels**: add a sheet `Pharmacy_Reorder_Levels` with columns `Brand, Generic, Min_Qty, Order_Up_To`.
- **Expiry**: batches expiring within 90 days are flagged, and within 30 days are shown in red. Expired stock still on the shelf is listed separately, with its value at cost.
- **Purchase order**: the suggestion brings each low item up to 30 days of stock plus 7 days for delivery. Lines are grouped by the supplier of the most recent batch and costed at the last buying price. It exports to CSV.

## Day summary

The summary contains figures only; no patient is named. It covers:

- Appointments and new registrations
- Admissions, discharges and bed use
- Money received today, by stream, and every open balance
- Stock that needs attention
- Tomorrow's reminders
- Sign-in trouble: failed attempts and locked accounts
- The last backup

Operations → Day summary shows the same content on screen, and **Email it now** sends it immediately.

## Backups

- The nightly copy is checked against the original: the number of rows on key sheets must match or be higher.
- Every backup is logged in `Backup_Log`, and Operations → Backups lists them.
- **Back up now** takes a copy immediately, for example before a big change.
- The folder and every copy are set to private. A backup is every patient record the clinic holds.
- A copy of the spreadsheet also carries the bound script code, but not the Script Properties or the triggers.
- Restoring is covered by the [restore checklist](RESTORE_CHECKLIST.md).

## Reference-list caching

The drug master, lab tests, clinical phrases, IP packages and active doctors are cached for 10 minutes and shared by all desks (`Master_Cache.gs`). The cached copy is dropped early in three cases:

- The app saves a change to one of these lists.
- Someone edits one of those sheets by hand in the spreadsheet (the `onEdit` trigger).
- An administrator clicks Operations → Backups → **Reload reference lists**.

Empty or failed reads are never cached.
