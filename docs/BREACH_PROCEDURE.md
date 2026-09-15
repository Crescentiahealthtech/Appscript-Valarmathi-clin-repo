# Breach procedure

**DPDP Act section 8(6): on a personal data breach, the Data Fiduciary shall
give intimation to the Data Protection Board AND to each affected Data
Principal.**

There is no materiality threshold in that section. The judgement the clinic
makes is about **what happened**, not about whether to say.

One page, because a procedure nobody can hold in their head is a procedure that
is read for the first time on the day it is needed.

---

## Who

| Role | Person | Fills in |
|---|---|---|
| **Decides it is a breach** | [name] — the grievance officer | |
| **Notifies the Board and the patients** | [name] — same person unless they are unavailable | |
| **Stands in when they are away** | [name] | |
| **Can be telephoned at any hour** | [number] | |

Write the names in. A procedure with a job title in it and no name is a
procedure where everybody assumes somebody else is doing it.

---

## What counts

A personal data breach is any unauthorised processing, or accidental
disclosure, acquisition, sharing, use, alteration, destruction or loss of
access to, personal data that compromises its confidentiality, integrity or
availability.

In this clinic, in plain terms:

- A report, prescription or bill sent to the wrong person.
- A document link forwarded beyond the patient it was for.
- A staff account used by somebody it does not belong to.
- The spreadsheet shared with an account that should not have it — including a
  staff member who has left.
- A laptop or phone with the clinic's records signed in on it, lost or stolen.
- Records deleted, overwritten or corrupted, where the clinic cannot get them
  back.
- The web app deployed so that it answers without a sign-in.

**"Nothing bad seems to have come of it" is not one of the tests.** Whether
anyone actually read the data goes to the consequences, not to whether it
happened.

---

## The four steps

### 1. Stop it — within the hour

Whoever notices does this, without waiting for anybody:

- Withdraw the document link: Privacy console → **A patient** → withdraw.
  (`dpdpRevokeDocumentLink`)
- Reset the credential: `crescAdminResetPassword("username")`.
- Remove the spreadsheet share: Sheets → Share → remove the account.
- Revoke the session: sign the account out; `revokeSession(token)` if you have
  the token.
- If the deployment is answering anonymously: **unpublish it**
  (Deploy → Manage deployments → archive) before anything else.

Do not delete anything, do not "tidy up" the sheet and do not edit the audit
log. What is there is the evidence.

### 2. Record it — the same day

Privacy console → **Breach register** → *Record an incident*, or
`dpdpRaiseBreach()`. Write what happened in sentences a stranger could follow
in two years.

Record it **even if you are sure it is nothing**. The register of incidents you
decided were not notifiable, with the reason, is the thing that shows an
inquiry that the clinic was looking. A register holding only confirmed breaches
proves nothing.

### 3. Decide — within 24 hours

Privacy console → **Notifiable** / **Not notifiable**, with a reason. The
software will not take the decision without one, and "not notifiable" is the
answer you will be asked to justify, so write that one most carefully.

What to weigh:

- **What data?** A name and a phone number is not the same as a diagnosis. HIV,
  pregnancy, mental health, substance use and genetic data carry a risk of
  discrimination that outlives any technical fix.
- **How many people?** One misdirected report is a breach. So is the register.
- **Who now holds it?** A relative who was forwarded a report is different from
  an unknown recipient, and both are different from nobody — but "we think
  nobody opened it" needs the open counter in `Document_Grants` to support it,
  not a feeling.
- **Can it be undone?** A revoked link that was never opened is a near miss.
  Write it down as one.

### 4. Tell — as soon as the facts are known

`dpdpBreachNotice("BRC-…")` drafts both notifications, filled in from the row.
Every square bracket in them is yours to complete.

- **The Data Protection Board of India.** Use the Board's current intimation
  route — check it before you need it, and write the route into this document.
- **Every affected patient.** By the route they would expect to hear from the
  clinic: telephone them if it is serious, then follow with the written notice.
  Do not lead with the section number. Lead with what happened to *them* and
  what they can do.

Then record that you did it: Privacy console, or
`dpdpRecordBreachNotification("BRC-…", "BOARD" | "PRINCIPALS", token)`.

### And afterwards

Close the incident with what changed so that it cannot happen the same way
again (`dpdpCloseBreach`). "Staff reminded to be careful" is not a change. A
setting, a check, a permission, a removed share — those are changes.

---

## How you will find out

Most breaches are noticed by a person. Two things watch as well:

- **The weekly review** (`dpdpWeeklyReview`, Mondays 07:00) reads the audit log
  for failed-login clusters against one account, failed logins spread across
  many identifiers — which is what walking the patient ID range looks like — an
  unusual number of records read by one person in a day, and one shared
  document opened six times or more. It emails the grievance officer. **Read
  that email.** It is the only part of this that runs on its own.
- **The readiness check** (`dpdpReadinessCheck`) reports plain-text credentials,
  an anonymous deployment and links past their window.

Neither replaces the person at the desk who says "that report went to the wrong
number".

---

## The clock

The Act does not put an hour count in section 8(6); the Rules and the Board's
practice will. Until that is settled, the clinic's own commitment is:

| Step | By |
|---|---|
| Contain | 1 hour from noticing |
| Record in the register | same day |
| Decide notifiable or not | 24 hours |
| Intimate the Board | 72 hours from deciding |
| Tell the affected patients | 72 hours, and sooner if they can act to protect themselves |

Being early is never the thing that gets a clinic into trouble.
