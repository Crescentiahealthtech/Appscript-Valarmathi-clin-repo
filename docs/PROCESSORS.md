# Processors, and the record of processing

**DPDP Act section 8(2): a Data Fiduciary remains responsible for personal data
processed on its behalf by a Data Processor, and may engage one only under a
valid contract.**

"We use Google" is not a contract. This is the list of everyone who touches
patient data on the clinic's behalf, what reaches them, what the contract is,
and what is missing.

---

## 1. Who processes what

| Processor | What reaches them | Contract | Status |
|---|---|---|---|
| **Google** — Sheets, Drive, Apps Script, Gmail, the account the script runs as | **Everything.** The whole record lives in a Google spreadsheet; documents are PDFs in Drive; every server function runs on Apps Script. | Google Workspace Data Processing Addendum — **only on a Workspace account**. A personal `@gmail.com` account is covered by the consumer terms, which are **not** a processor agreement. | 🔴 **CHECKED: the spreadsheet is owned by `crescentiahealthtech@gmail.com`, a personal Google account. There is no DPA.** |
| **Meta / WhatsApp** | The patient's name and mobile number in the message, and a link to a document about them. Not the document itself. | WhatsApp's consumer terms, which are not a processor agreement. | 🟡 **The consent route is now built, and enforced.** Every WhatsApp dispatch checks the patient's `COMMUNICATION` consent server-side and refuses without it (`DPDP_Dispatch.gs`), and the notice the patient is read **names Meta**, says the message passes through their service outside India, and says the link can be opened by anyone who gets hold of the message. Remaining option if the clinic wants a contract rather than consent: move to the WhatsApp Business API. |
| **Google (Gmail)** — outbound email | The patient's document as an attachment, to the address on their record. Also the temporary password for a self-service reset. | Same as the row above: the DPA depends on the account being a Workspace one. | 🟡 Gated on the same `COMMUNICATION` consent as WhatsApp, and the notice says email is not encrypted end to end. Covered by whatever contract covers the Google row above — which today is none. |
| **Google / Microsoft speech services** | The **audio** of a clinician dictating a clinical note, if voice typing is enabled. | Nothing specific. It is the browser's own feature, used by the clinician's browser. | ⬜ Decide: `dpdpSetVoicePolicy("ALLOWED"\|"FORBIDDEN")`. If allowed, it goes in the notice. |
| **The insurer / TPA named on a claim** | The patient's identity, the admission, the diagnosis and the bill. | The insurer's own empanelment agreement. | ⬜ Confirm the agreement has a data-protection clause. Consent for this purpose (`INSURANCE`) is recorded per patient. |
| **[Your SMS gateway or WhatsApp Business API, if reminders are sent automatically]** | Name, mobile, and the date and kind of visit (appointment, follow-up, child's vaccination). Set up via `REMINDER_PROVIDER` in Script Properties (`Patient_Reminders.gs`). | The gateway's terms; Meta's WhatsApp Business terms include a data processing addendum. | ⬜ Only patients with `COMMUNICATION` consent are messaged, automatically or from the desk. Name the gateway here before switching it on. |
| **[Your accountant or auditor, if they get exports]** | Bills, patient names, sometimes diagnoses on insurance claims. | | ⬜ A person outside the clinic handling patient data is a processor, even when they are a family friend. |

### The first one is the one that matters, and it has been checked

**The clinic's spreadsheet is owned by `crescentiahealthtech@gmail.com`.** That
is a personal Google account, not Google Workspace. It was read from the file's
own metadata, so this is not an inference.

That means, today:

- there is **no Data Processing Addendum**, so section 8(2) is not met for
  100% of the clinic's personal data;
- there is no admin console, no audit of who the file was shared with, no
  ability to transfer ownership when someone leaves, and no Google support
  obligation;
- the account's owner personally, rather than the clinic, controls the records.

**Migrating to Google Workspace is the single highest-value non-engineering
action available to this clinic.** It is not a code change and nothing in this
repository can do it. Roughly what it involves:

1. Buy Workspace on the clinic's own domain (Business Starter is enough for
   this; the DPA comes with every paid tier).
2. Create a clinic account for each member of staff.
3. **Transfer ownership** of the spreadsheet and the Drive folders
   (`Crescentia_Lab_Invoices`, `Valarmathi_OP_Prescriptions`,
   `Crescentia_Pharmacy_Invoices` and the rest) to the clinic account —
   transfer, not copy, so there is one authoritative record and not two.
4. Move the Apps Script project with them, and **redeploy**. The `/exec` URL
   changes, so re-print anything that carries it.
5. Set the Workspace data region if the clinic wants storage in India.
6. Remove the personal account's access, and check it holds no leftover copy.

Until that is done, two things follow that the clinic should be honest about:
section 8(2) is not met for any of its patient data, and if anything happens to
that personal account — lost password, recovery failure, the person who owns it
leaving — the clinic's entire medical record goes with it. There is no admin
console to recover it from.

---

## 2. Record of processing

What is held, why, on what basis, and where it goes. This is the table an
inquiry asks for first.

| Category | Examples | Purpose | Basis | Retention | Leaves the clinic to |
|---|---|---|---|---|---|
| Identity | name, age, sex, date of birth, patient ID, photo ID if taken | care, billing | s.7 legitimate use | 3 years from last visit | insurer (on consent) |
| Contact | mobile, WhatsApp, email, address | care, billing, reminders | s.7 for care; consent for reminders | as above; marketing data 1 year | Meta (on consent) |
| Health | history, examination, diagnoses, prescriptions, lab results, ward notes, discharge summaries | care | s.7 legitimate use | 3 years outpatient, longer inpatient / medico-legal | insurer (on consent), the patient |
| Financial | bills, payments, credit, insurance claims | billing | s.7 legitimate use | 8 years | insurer, accountant, tax authorities |
| Relationships | guardian, emergency contact, nominee | care, s.9 guardian consent, s.14 nomination | s.7 / consent | with the clinical record | — |
| Staff credentials | username, salted password digest, role, MFA secret, email | running the system | s.7 (employment) | while employed + 1 year | — |
| Access records | `Audit_Log`, `Auth_Audit`, `Sessions`, `Document_Grants` | security, s.8(5) | s.7 | 3 years | — |
| Consent records | `Consent_Register`, `Nomination_Register`, `DPDP_Requests`, `Breach_Register` | proving compliance | s.6(10), s.8(6) | with the record, or permanently for breaches | the Board, on a breach |

### Special category, in practice

Indian law does not have the GDPR's "special categories", but health data
carries a real-world risk that the clinic should treat as higher: HIV status,
pregnancy and termination, mental health, substance use, genetic findings, and
anything about a minor. Where a report contains one of these, the consequences
section of a breach notification is not boilerplate.

---

## 3. Cross-border transfer

Section 16 lets the Central Government restrict transfer to notified countries.
As of this writing no restricting notification is in force, which means transfer
is currently permitted — but the clinic should know, and be able to say, that:

- **Google Workspace stores and processes data in data centres in several
  countries, not only India.** Workspace admins can set a data region policy;
  personal Gmail accounts cannot.
- **Meta processes WhatsApp message metadata outside India.**

If the clinic wants everything in India, that is a Workspace data-region
decision plus dropping WhatsApp dispatch — a business decision with a cost, not
a setting in this repository.

---

## 4. What to do with this file

1. Fill in the blank rows and the ⬜ boxes.
2. **Start the Workspace migration.** The question is answered — the records
   are on a personal account — so this is the action, not the enquiry. Nothing
   else on this page moves section 8(2) as far.
3. For each processor without a contract, either get one or stop using them for
   personal data.
4. Review it when anything changes — a new gateway, a new insurer, a new
   accountant — and at least once a year.
5. Keep it with the consent notice, the retention schedule and the breach
   procedure. Those four documents plus `docs/DPDP_READINESS.md` are the
   clinic's compliance file.
