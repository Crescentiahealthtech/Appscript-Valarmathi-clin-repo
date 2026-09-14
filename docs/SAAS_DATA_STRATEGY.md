# Data at scale: what breaks, in what order, and what to do about it

*Crescentia HealthTech / CresRx — written against the code as it stands, September 2026.*

The question this answers: **"Though multiple sheets are there, the entries and data
would keep on piling. How should we handle it as SaaS and how should we manage data?"**

Short version: the sheet count is not the problem, and the cell limit is not the
problem you will hit first. **Read volume is.** The application will start timing out
long before it runs out of room, and it will do so without warning, on the busiest
day of the month, on the largest sheet.

---

## 1. Where the system actually stands today

Measured from the repository, not estimated:

| | |
|---|---|
| Distinct sheets | **52**, in one spreadsheet |
| `getDataRange().getValues()` calls | **125** |
| Tenants | **1**. `getTenantId_()` returns the literal string `"VALARMATHI"` |
| Deployment | One Apps Script project, bound to one Google Sheet |
| Retention policy | None. Nothing is ever archived or deleted |

Two of those lines are the whole problem.

`getDataRange().getValues()` pulls **an entire sheet** into memory to look at a
handful of rows. There are 125 of them. It costs nothing today because the sheets
are small, and it will keep costing nothing right up until it doesn't.

`getTenantId_()` returning a constant means this is not multi-tenant software yet.
That is a decision to make deliberately (§6), not a gap to patch.

---

## 2. The three ceilings, in the order you will hit them

### Ceiling 1 — the six-minute execution limit (you hit this first)

Apps Script kills a script at **6 minutes** (30 on Workspace Enterprise tiers). It
does not warn, it does not degrade: the call fails and the user sees a spinner that
never stops.

A `getDataRange().getValues()` on a 50,000-row × 20-column sheet moves a million
cells and takes seconds. A screen that does three of them takes tens of seconds. A
screen that does one *inside a loop* — and several do — is quadratic.

The order these break in follows row growth, so the fastest-growing sheets go first:

| Sheet | Grows by | First to break because |
|---|---|---|
| `Audit_Log` | every clinical and financial action | append-only, never pruned, read whole by several callers |
| `LAB_RESULTS` | one row **per parameter**, not per test | a CBC alone is ~20 rows |
| `IP_Timeline_DB` | every ward note, every shift | read whole on every timeline open |
| `Pharmacy_Invoice_Items` | one row per dispensed batch line | ~4× the invoice count |
| `Sessions` | one row per sign-in, forever | `revokeSession()` reads it whole on every sign-out |

**Signal that you are close:** any screen that used to open instantly starts taking
more than two seconds, consistently, and gets worse month over month. That is not a
slow network. That is a sheet that has outgrown a full read.

### Ceiling 2 — 10 million cells per spreadsheet (hard, and final)

Google Sheets refuses writes past **10,000,000 cells per spreadsheet**, counted
across every tab. At ~20 columns average that is about **500,000 rows in total, for
the whole clinic, for all time**.

Rough arithmetic for a single busy clinic: 80 OP encounters a day, each producing an
encounter row, a billing row, several prescription rows and several audit rows; plus
lab orders at 15–30 rows each once you count `LAB_ORDER_TESTS`, `LAB_SAMPLES`,
`LAB_RESULTS`, `LAB_TAT_LOG` and `LAB_BILLING`. Call it 250–400 rows a day, ~100,000
a year.

**That is roughly four to five years for one clinic.** For a SaaS with several
clinics in one workbook, divide by the number of clinics. There is no way to raise
this limit and no warning before it lands — writes simply begin to fail.

### Ceiling 3 — concurrency

`LockService.getScriptLock()` is a **single global lock for the entire project**.
When the discharge module holds it for a long assembly, the pharmacy counter waits.
One clinic at moderate load is fine. Several clinics sharing a deployment are not.

---

## 3. What to do, in the order that buys the most time per hour spent

### 3.1 Stop reading whole sheets (highest value, lowest risk)

The pattern is already in this codebase and already correct — the discharge module
uses it. `createTextFinder` on a single column, then re-read only the rows it names:

```js
// Instead of: sheet.getDataRange().getValues()  ← the whole sheet
var col  = map['OrderID'] + 1;
var hits = sheet.getRange(2, col, sheet.getLastRow() - 1, 1)
                .createTextFinder(orderId).matchEntireCell(true).findAll();
var row  = sheet.getRange(hits[0].getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
```

`labVerifyPage_()` in `LabIntegrationEngine.gs` is a worked example on a public route.

For append-only sheets read in reverse chronological order — the audit log, the
disposal register, anything "recent activity" — **read the tail, not the whole
thing**. `crescGetLoginAudit()` and `getDisposalLog()` both do this:

```js
var from = Math.max(2, sheet.getLastRow() - limit + 1);
var data = sheet.getRange(from, 1, sheet.getLastRow() - from + 1, cols).getValues();
```

Convert in this order — biggest sheet on the hottest path first:

1. `IP_Timeline_DB` — the ward notes timeline
2. `LAB_RESULTS` / `LAB_ORDERS` — the lab queue
3. `IP_Pharmacy_Queue` — the pharmacy queue
4. `Audit_Log` — every reader
5. `Sessions` — `revokeSession()` in particular

### 3.2 Add a period close, and mean it

`AccX.lockPeriod()` already exists. Extend the idea from accounting to data:

**A closed period is read-only and is never read in full again.** Once March is
closed, no query about March touches the live sheet; it reads the archive. This is
what makes archiving safe — without a close, archiving races with edits.

### 3.3 Hot / warm / cold

| Tier | Where | Holds | Read by |
|---|---|---|---|
| **Hot** | the live sheets | the current period — the last 90 days | everything |
| **Warm** | one archive spreadsheet per financial year | closed periods | reports, the patient's own history |
| **Cold** | a Drive folder of per-year CSV or JSON exports | everything past the retention floor | a human, on request |

A monthly trigger moves rows past the hot window into the year's archive workbook and
deletes them from the live sheet, leaving **one index row per patient per period** —
patient id, period, archive file id, row range. That index is what makes a cold
record findable without opening every archive.

This is the single change that keeps a Sheets-backed system alive for years rather
than for four of them, because it caps the *live* row count at a constant instead of
letting it grow without limit.

### 3.4 Set a retention floor before you build the archiver

Archiving without a retention policy just moves the pile. Every sheet needs a class,
and the class decides how long the data lives. These are the anchors to *start the
conversation with your compliance advisor* — confirm them against current law before
you encode any of them:

| Class | Example sheets | Anchor |
|---|---|---|
| Inpatient clinical record | `IP_CaseSheets_DB`, `IP_Timeline_DB`, `DS_Summaries` | IMC (Professional Conduct) Regulations 2002, Reg. 1.3.1 — indoor records **3 years** from commencement of treatment |
| Outpatient clinical record | `OP_Encounters`, `Encounters_DB` | Commonly aligned to the same 3 years; many clinics keep longer |
| Medico-legal cases | any of the above flagged MLC | Kept substantially longer — treat as a separate, never-purged class |
| Minors | any clinical record for a patient under 18 | Until majority **plus** the limitation period |
| Books of account | `Billing_Ledger`, `Audit_Event_Ledger`, `Pharmacy_Invoices` | Companies Act 2013, s.128 — **8 years** |
| Tax | the same, plus GST returns | GST: **72 months** from the annual return due date |
| Audit / access logs | `Audit_Log`, `Sessions` | No fixed statutory floor; long enough to investigate an incident. 2 years is a reasonable default |
| Disposal register | `Pharmacy_Disposal_Log` | Drug destruction records — keep with the pharmacy licence records |

Two things worth saying plainly:

- **Under the DPDP Act 2023, keeping data longer than you need it is now a
  liability, not just a storage cost.** Personal data has to be erased once the
  purpose is served and no law requires retention. "We never delete anything" stopped
  being the safe answer.
- **A retention policy is a product feature, not an internal chore.** Every clinic
  buying this will eventually ask where their data goes and how long you keep it.
  Having an answer is part of what you are selling.

### 3.5 Widen what does not need to be a row at all

Several sheets store derived data that could be recomputed: TAT logs, some queue
tables, some of the master ledger. Every row you never write is a row you never
archive. Worth one pass before building the archiver, not after.

---

## 4. The audit log deserves its own paragraph

There are currently **three** audit trails — `Audit_Log`, `Audit_Logs` and
`Audit_Event_Ledger` — with three different column sets, written by three modules
that do not know about each other. Clinical note edits and case sheet amendments are
in none of them.

That is a correctness problem before it is a scale problem: *"who touched this
patient's record, and when"* cannot be answered from any one of them.

It is also the fastest-growing sheet in the system, so it is where full reads will
hurt first. Merge into one append-only `Audit_Log` with one schema **before** the
volume makes the merge a project of its own — and give it the tail-read treatment
from §3.1 at the same time. Sign-in events already write there
(`Auth_Audit.gs`); the rest should follow.

---

## 5. What to measure, so you find out before your users do

Add a scheduled function that writes one row a week:

```js
// per sheet: name, rows, columns, cells
// per workbook: total cells, % of the 10,000,000 ceiling
// slowest observed response time per screen
```

Three thresholds worth alerting on:

- any sheet past **20,000 rows** → convert its readers to indexed reads
- workbook past **60% of the cell ceiling** → start the archiver now, not later
- any screen consistently past **3 seconds** → find the full read behind it

Without this you will discover the ceiling by hitting it.

---

## 6. The SaaS question: what "multi-tenant" has to mean here

`getTenantId_()` returns a hardcoded `"VALARMATHI"`, and a `Tenant_ID` column already
exists on several sheets. So the shape was anticipated and never built. There are
three honest options, and they are not equally good.

### Option A — one spreadsheet + one Apps Script deployment per clinic (what you have)

- **For:** perfect data isolation. One clinic's cell ceiling is theirs alone. One
  clinic's script lock is theirs alone. A breach is contained to one clinic. It is
  also the easiest thing to sell to a hospital's compliance officer.
- **Against:** every deployment has to be updated separately. With 5 clinics that is
  an afternoon; with 50 it is a full-time job, and versions drift until you cannot
  reproduce a bug.
- **Verdict:** correct for the first 5–10 customers. Invest in a **scripted
  deployment** (`clasp` + a release checklist) early, because the pain is
  update-management, and that is solvable.

### Option B — one shared workbook, `Tenant_ID` on every row

- **For:** one deployment, one update.
- **Against:** every clinic shares the 10M-cell ceiling, so it arrives N times
  sooner. Every clinic shares the single global script lock. And one missing
  `Tenant_ID` filter anywhere — in 125 read sites — shows one clinic another
  clinic's patients. That is the kind of bug that ends a healthcare SaaS.
- **Verdict:** **do not do this on Sheets.** The isolation guarantee is too weak for
  patient data and the ceiling maths is against you.

### Option C — move the data layer off Sheets, keep the Apps Script front end

Keep the UI exactly as it is; replace the sheet reads with a real database
(Cloud SQL, Firestore, or Supabase) behind the same function signatures. `UrlFetchApp`
or the JDBC service reaches either.

- **For:** removes every ceiling in §2 at once. Real indexes, real concurrency, real
  row-level security, real backups, real point-in-time recovery.
- **Against:** it is the largest piece of work on this page, and it is only worth
  starting when the ceilings are actually near.
- **Verdict:** this is the destination. The work in §3.1 is not wasted on the way
  there — an indexed read is already shaped like a query, which is what makes the
  eventual port mechanical rather than a rewrite.

### The recommendation

1. **Now:** §3.1 indexed reads, §3.4 retention classes, §5 measurement. Stay on
   Option A. Script the deployment.
2. **At 5–10 clinics, or the first sheet past 20,000 rows:** build the §3.3 archiver.
   Still Option A.
3. **At ~20 clinics, or the first workbook past 60% of the ceiling:** start Option C,
   one module at a time, behind the existing function signatures. Lab and audit first
   — they are the biggest and the least entangled with the UI.

Skipping straight to step 3 is a common and expensive mistake. Steps 1 and 2 are
weeks of work that buy years; step 3 is months.

---

## 7. The thing that is not about volume at all

Two of the items in `OPEN.md` are worth saying here because they get worse with
every row added, not better:

- **Staff passwords are stored in the spreadsheet as plain text.** Every person the
  workbook has ever been shared with can read every password. Hash them
  (`Utilities.computeDigest`, already used elsewhere in this project) with a
  per-user salt. This needs a one-time migration and a forced reset, which is why it
  is a decision rather than a silent change.
- **The web app is deployed `ANYONE_ANONYMOUS` + `USER_DEPLOYING`.** Anyone with the
  URL can call any server function from the browser console *without signing in*, and
  each call runs with the deploying account's full access to the spreadsheet.
  `RBAC.gs` is the fix; `crescRbacCoverage()` reports how far it has got.

Growing the data set makes both of those worse in exact proportion to how much data
there is.

---

## Appendix — the checks that already exist

Run from the repository:

```
./tools/check.sh
```

Run from inside the Apps Script editor:

| | |
|---|---|
| `crescRbacCoverage()` | which server functions still have no permission check |
| `crescRbacSelfTest()` | the permission matrix is internally consistent |
| `crescVerifyBundles()` | every lazy-loaded UI partial exists |
| `verifyDeployment()` | which `.gs` files are missing from the deployment |
| `runIPHealthCheck()` | IP schema drift |
