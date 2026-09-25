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
   | `REMINDER_AUTO_EMAIL` | Set it to `NO` to stop the 6 pm job from emailing reminders by itself. By default it emails every patient who has consented and has an email address. |
   | `REMINDER_PROVIDER` | Leave it blank to send WhatsApp from the desk. Set it to `WHATSAPP_CLOUD` to send WhatsApp automatically through the WhatsApp Business API. |
   | `WA_PHONE_NUMBER_ID`, `WA_ACCESS_TOKEN`, `WA_TEMPLATE_NAME`, `WA_TEMPLATE_LANG` | WhatsApp Business Cloud API. The template must be approved and take four body parameters: `{{1}}` name, `{{2}}` what, `{{3}}` when, `{{4}}` clinic. |

3. Run `RUN_00_authorizeServices()` (in `RUN_Setup.gs`) once and allow the permissions it asks for. The web app runs as you, so Google sign-in (which checks the token with Google), backups and emails only work after you have approved them. Until you do, Google sign-in shows "Google sign-in is not switched on for this clinic yet". Run it again after any update that adds a Google service.

## Reminders

- **Appointments**: sent for appointments that are still *Booked*.
- **Follow-ups**: sent when `OP_Encounters.Next_Review_Date` is that day and the patient has no appointment booked for it.
- **Vaccinations**: sent when a child's next dose on the immunisation card (Maternal & Child) falls due that day.
- **Consent**: only patients whose DPDP `COMMUNICATION` consent is *Given* are messaged. The list shows everyone else with the reason, so the desk can ask them at their next visit.
- **Two channels, WhatsApp and email**:
  - *WhatsApp*: the WhatsApp button opens the message on the clinic computer's WhatsApp. Opening it records nothing, because the desk may go back without sending. Once it has gone, press the tick (✓) to close the row.
  - *Email*: sent by the app from the clinic's Google account to the patient's `Email` (Patients, column Q). The Email button sends one immediately, **Send now** sends all of them, and the 6 pm job sends tomorrow's by itself. A personal Gmail account can send about 100 emails a day.
- **Sending again**: a closed row can always be reopened with **Send again**. The automatic job never sends the same reminder twice.
- **Record**: every reminder, sent or failed, is logged in `Reminder_Log`.

## Stock alerts

- **Low stock** is measured in days: current stock divided by the average daily sale over the last 60 days (from `Pharmacy_Invoice_Items`). An item is low when it has fewer than 14 days left.
- **Items with no recent sales** use the plain floor (`STOCK_MIN_UNITS`).
- **Per-medicine levels**: add a sheet `Pharmacy_Reorder_Levels` with columns `Brand, Generic, Min_Qty`.
- **Expiry**: batches expiring within 90 days are flagged, and within 30 days are shown in red. Expired stock still on the shelf is listed separately, with its value at cost.
- **Ordering**: no order quantity is suggested; how much to order is the pharmacist's decision. Each low item shows what is left, how fast it sells, and the supplier and buying price of its latest batch. The list exports to CSV.

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

## Schedule H, H1 and X medicines

- **Marking a medicine**: set **Drug Schedule** under Pharmacy → Add New Stock, or in the edit dialog under Live Inventory. The schedule is stored on the Pharmacy_Inventory row in a `Schedule` column, which is added on first use. It applies to **every batch** of that medicine. A new batch of a medicine already on the shelf inherits the schedule. In Live Inventory, search `sch h1` (or `sch h`, `sch x`) to list scheduled medicines.
- **At billing**: a scheduled medicine shows a red **Sch H / H1 / X** badge. The bill cannot be saved until the **Prescribing Doctor** names a doctor ("Self / OTC" is refused) and the patient's name is filled in. The server enforces this even if the page does not. The clinic's own doctors are suggested; an outside doctor can be typed in.
- **The register**: Pharmacy → **Schedule H Register**. Choose a period and the schedules, then **Print register** (A4 landscape): S.No, date of issue, patient name, doctor's name, drug issued, quantity, and a blank signature column. It also exports to CSV.
  - Each invoice line records its schedule at the moment of sale, so reclassifying a medicine later does not change past entries.
  - Sales from before the Schedule column existed are listed by the medicine's schedule today and marked `*`.
  - Entries sold before the doctor's name was compulsory are flagged so it can be written in by hand.
  - Cancelled bills are left out. Returned units are shown on the line.

## Billing ledgers (pharmacy and lab)

Pharmacy → **Billing Ledger** and Lab → **Billing Ledger**, matching the hospital billing ledger.

- Filters: a period (or **Everything still owed**, which covers all dates), status (settled, unsettled, unpaid, part-paid, on IP account, cancelled) and a search.
- Totals: billed, collected, still owed and refunded; money collected by mode; and what is owed by age (0–7, 8–30, 31–90, over 90 days).
- Each row can be **reprinted**, and anything owed can be **collected**. After a collection the desk offers to print the receipt.
- The list prints on the letterhead and exports to CSV.
- **Lab reprints** now list every payment received on the bill (at billing, and each later collection) and what is still owed. An unsettled lab bill can be reprinted from the Lab Billing desk's *Unsettled* tab.

## Sheet data repair

Operations → Backups → **Check sheet data** (or `repairSheetData()` in the editor) previews what it would fix and changes nothing until you press **Apply**. **Take a backup first.** It fixes:

- the missing `Fee` and `Timestamp` headers on Appointments (columns H and I);
- timestamps stored as ISO text (`2026-05-17T04:53:02.393Z`) in Appointments and Pharmacy_Inventory;
- any cell holding a stringified date (`Sat Dec 30 1899 11:45:00 GMT+0521 …`), rewritten as the date or time it meant. Signed discharge snapshots and the audit trails are never rewritten; they are cleaned where they are displayed instead;
- LAB_AUDIT_LOG rows written under the wrong header, moved into their named columns;
- the IP_Discharge_Drafts header, so saved ward charges load back.

It is safe to run again: a second run finds nothing to do.
