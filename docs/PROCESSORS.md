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
| **Google** — Sheets, Drive, Apps Script, Gmail, the account the script runs as | **Everything.** The whole record lives in a Google spreadsheet; documents are PDFs in Drive; every server function runs on Apps Script. | Google Workspace Data Processing Addendum — **only on a Workspace account**. A personal `@gmail.com` account is covered by the consumer terms, which are **not** a processor agreement. | ⬜ **CHECK WHICH ONE THIS DEPLOYMENT IS** |
| **Meta / WhatsApp** | The patient's name and mobile number in the message, and a link to a document about them. Not the document itself — links are now private and expiring. | WhatsApp's consumer terms, which are not a processor agreement. | ⬜ Move to the WhatsApp Business API with a contract, **or** rely on the patient's specific consent (`COMMUNICATION`) and say in the notice that the message goes through Meta. |
| **Google / Microsoft speech services** | The **audio** of a clinician dictating a clinical note, if voice typing is enabled. | Nothing specific. It is the browser's own feature, used by the clinician's browser. | ⬜ Decide: `dpdpSetVoicePolicy("ALLOWED"\|"FORBIDDEN")`. If allowed, it goes in the notice. |
| **The insurer / TPA named on a claim** | The patient's identity, the admission, the diagnosis and the bill. | The insurer's own empanelment agreement. | ⬜ Confirm the agreement has a data-protection clause. Consent for this purpose (`INSURANCE`) is recorded per patient. |
| **[Your SMS gateway, if any]** | Name, mobile, appointment or result availability. | | ⬜ |
| **[Your accountant or auditor, if they get exports]** | Bills, patient names, sometimes diagnoses on insurance claims. | | ⬜ A person outside the clinic handling patient data is a processor, even when they are a family friend. |

### The first one is the one that matters

Everything in this repository runs inside one Google account. If that account is
a personal Gmail account:

- there is **no Data Processing Addendum**, so section 8(2) is not met for
  100% of the clinic's personal data;
- there is no admin console, no audit of who the file was shared with, no
  ability to transfer ownership when someone leaves, and no Google support
  obligation;
- the account's owner personally, rather than the clinic, controls the records.

**How to check:** sign in and open <https://admin.google.com>. A Workspace
account reaches an admin console; a personal account does not. Alternatively,
the email address is the answer — `something@gmail.com` is personal,
`something@yourclinic.in` is usually Workspace.

If it is personal: migrating to Workspace is the single highest-value
non-engineering action on this list. Until then, say so honestly in any answer
to a patient about where their data is held.

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
2. Get the Workspace question answered this week. It is one click and it
   determines whether section 8(2) is met at all.
3. For each processor without a contract, either get one or stop using them for
   personal data.
4. Review it when anything changes — a new gateway, a new insurer, a new
   accountant — and at least once a year.
5. Keep it with the consent notice, the retention schedule and the breach
   procedure. Those four documents plus `docs/DPDP_READINESS.md` are the
   clinic's compliance file.
