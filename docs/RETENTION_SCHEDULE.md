# Retention schedule

**DPDP Act section 8(7): erase personal data when the purpose it was collected
for is served, unless the law requires it to be kept.**

Two obligations pull in opposite directions here, and a clinic that only feels
one of them gets into trouble either way. Keeping everything for ever is a
section 8(7) breach. Deleting on a timer is how a clinic loses the record it is
about to be asked for by a consumer forum, an insurer or a court.

So this schedule says, for each kind of record: how long, on what basis, and
**who decides** — and the software reports rather than deletes.

This is an engineering document, not legal advice. The periods below are the
starting points in `DPDP_CFG.RETENTION_DAYS` (`DPDP_Compliance.gs`). A lawyer
signs them off, the clinic adopts them, and then they are the clinic's.

---

## The schedule

| Record | Period | From when | Basis | Who decides to delete |
|---|---|---|---|---|
| Outpatient clinical record | 3 years | last entry | NMC (Professional Conduct, Etiquette and Ethics) Regulations 2002, reg. 1.3.1 | the treating doctor |
| Inpatient record, case sheet, ward notes | **longer — take advice** | discharge | state Clinical Establishments rules; many require 5 years or more | the treating doctor |
| Medico-legal record (any injury, poisoning, assault, road traffic accident, custodial or unnatural death) | **do not delete without legal advice** | — | limitation periods run from discovery, not from the event | the clinic's lawyer |
| A minor's clinical record | until **3 years after the patient turns 18**, and never less than the ordinary period | last entry | limitation does not run against a minor | the treating doctor |
| Bills, receipts, ledgers, GST records | 8 years | end of the financial year | Income Tax Act s.44AA and s.44AB practice; GST records 72 months | the accountant |
| Insurance and TPA claim records | 8 years | claim settled | the same, plus the insurer's own dispute window | the accountant |
| Sign-in sessions (`Sessions`) | 30 days | issue | no reason to keep them; `purgeExpiredSessions()` clears them at 7 days past expiry | automatic |
| Audit log (`Audit_Log`) | 3 years | the entry | matches the clinical period it documents, so the record and the account of who touched it expire together | the clinic |
| Document links (`Document_Grants`) | 14 days live, register kept 3 years | issue | the link is a capability and expires; the ROW is evidence of a disclosure under s.11(1)(b) | automatic for the link, the clinic for the row |
| Breach register (`Breach_Register`) | **permanent** | — | it is the evidence that the clinic looked, decided and acted | never |
| Consent register (`Consent_Register`) | as long as the record it covers, plus 3 years | the decision | s.6(10) — the clinic has to be able to prove consent after the fact | never, while the clinical record stands |
| Nomination register (`Nomination_Register`) | as long as the clinical record | the nomination | s.14 | never, while the clinical record stands |
| Marketing contact data | 1 year, or immediately on withdrawal | last visit | no statutory basis to keep it at all | automatic on withdrawal |
| Data principal requests (`DPDP_Requests`) | 3 years | closure | evidence that statutory requests were answered in time | the clinic |

### Two rules that override every row above

1. **A legal hold beats the schedule.** The moment a record is the subject of a
   complaint, a claim, a police request, a consumer case or any inquiry, it is
   kept until that ends — whatever the table says. Write the hold down, with
   the date and the reason, before the period expires.

2. **An erasure request is assessed, not obeyed or refused.** Section 12(3)
   gives the right; section 8(7) preserves what the law requires the clinic to
   keep. So the answer to a patient asking for erasure is usually *partial*:
   the marketing number goes today, the clinical record goes when its period
   ends, and the patient is told which is which and why. `closeDPDPRequest()`
   will not accept "Closed" as an outcome for exactly this reason.

---

## How it runs

**Nothing deletes clinical or financial records automatically, and that is
deliberate.** `dpdpRetentionReport()` produces the list; a person acts on it.

| What | When | What it does |
|---|---|---|
| `dpdpDailyMaintenance()` | daily, 02:00 | expires document links, revokes any legacy public Drive shares, purges dead sessions |
| `dpdpMonthlyRetentionReport()` | 1st of the month, 06:00 | emails the grievance officer what is past its period |
| `dpdpEraseUnusedFields(true)` | once, by hand | clears the three columns nothing ever read |

Install them with `dpdpInstallTriggers()`; confirm with `dpdpTriggerStatus()`.

## The monthly routine, for whoever holds this

1. Open the retention email.
2. For each line, ask: is any of this under a legal hold? If yes, note it and
   move on.
3. For clinical records past their period: give the list to the treating
   doctor. Nobody else decides.
4. For financial records past their period: give the list to the accountant.
5. Delete what is agreed, **in the spreadsheet, by hand**, and write down what
   was deleted and who agreed. The deletion itself is a fact the clinic may
   need to prove later.
6. If nothing is deleted this month, write that down too. A schedule that is
   reviewed and consciously not acted on is compliance; a schedule nobody
   opened is not.

## What is NOT covered here

- **Google's own copies.** Deleting a row in Sheets does not delete Google's
  backups on Google's schedule. That is Google's processing under the Workspace
  DPA, and the clinic's answer to a patient should say so rather than promise
  an erasure it cannot perform.
- **Documents already sent.** A PDF a patient downloaded is theirs, and a
  WhatsApp message forwarded to a relative is beyond recall. Withdrawing the
  link (`dpdpRevokeDocumentLink`) stops further access; it does not unsend.
- **Printed paper.** Anything printed from this system leaves it. The clinic's
  paper retention and shredding policy is a separate document and needs to
  exist.
