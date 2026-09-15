# Consent Notice

**Digital Personal Data Protection Act, 2023 — sections 5 and 6**

This is the notice the patient is given **at or before** their information is
collected. It is the words behind `getDPDPNotice()`: the purposes listed here
are the purposes in `DPDP_PURPOSES` (`DPDP_Compliance.gs`), and the rights
listed here are the ones the application can actually act on.

**Before you use it**, three things are yours to fill in, and one is your
lawyer's:

1. Replace every `[…]` with the clinic's own detail.
2. Have it translated into **every language the clinic serves** — in this
   clinic that means Tamil, and a notice in English only is not a notice given
   to a patient who does not read English. Section 5(3) gives the Data
   Principal the right to the notice in any language in the Eighth Schedule.
3. Decide the two open questions marked **DECIDE** below.
4. Have a lawyer read it and decide whether the clinic is a Significant Data
   Fiduciary under section 10.

When you change the words, **bump the notice version** so that consents
already recorded are not silently attributed to a notice nobody agreed to:

```
PropertiesService.getScriptProperties()
  .setProperty('DPDP_NOTICE_VERSION', 'v2-20260401');
```

`getConsentStatus()` then marks every older consent `stale`, which is the
prompt to ask again — not a bug.

---

## The notice

### Your information, and what we do with it

**[Clinic name]**
[Address]
[Phone] · [Email]

We are the **Data Fiduciary** for the information we hold about you. That means
we are responsible for it, and answerable to you for it.

### What we collect

When you register with us we collect your name, age, date of birth, sex, mobile
number, WhatsApp number if you give one, email address if you give one, home
address, your known conditions, your blood group, who to contact in an
emergency, and who referred you if anybody did.

When we treat you we record what you told us, what we found, what we tested for
and what the results were, what we prescribed, what we admitted you for and what
happened while you were with us, and what you were billed and paid.

We do not collect anything else. If a form here ever asks you for something and
you cannot see why, ask us — and if we cannot give you a reason, we should not
be asking.

### Why we hold it, purpose by purpose

**Care and treatment.** Recording your history, examination, investigations,
prescriptions and admissions so that you can be treated, and so that the next
doctor who sees you knows what happened.

**Billing and accounts.** Raising your bills and receipts, and keeping them for
as long as tax law requires.

These two are what you came to us for. We process them as a **legitimate use**
under section 7 of the Act rather than asking you to tick a box, because a
consent form should not be able to stop somebody being treated. You can still
ask us what we hold, ask us to correct it, and complain about how we handle it.

Everything below is **your choice**, and you may say no to any of it and be
treated exactly the same:

**Insurance and TPA claims.** Sharing your records with your insurer or their
third-party administrator so that a claim can be settled.

**Reports and reminders by WhatsApp or email.** Sending your prescription, lab
report or invoice to your mobile number or email address. If you say yes: those
messages travel through WhatsApp, which is run by Meta — a company outside this
clinic and outside India — and a link we send you opens a document about you.
Anyone who gets hold of that message can open it, so please do not forward it.
We send links that stop working after a short time for that reason. Email is
not encrypted end to end: your email provider, and anyone with access to your
inbox, can read what we send. **If you say no, or take it back later, nothing
is sent** — the system refuses it, and you collect your documents at the
clinic instead.

**Health camps and offers.** Telling you about camps, screenings and services.
Say no and nothing about your treatment changes.

**De-identified research and audit.** Using your data with your name and
contact details removed, to check and improve how we treat people.

### Dictated notes

**[DECIDE — delete the paragraph that is not true of your clinic]**

*If voice typing is allowed:* Some of our clinicians dictate notes instead of
typing them. The dictation is turned into text by the speech service built into
their web browser, which means the audio of what they say — which will include
things you told us — is sent to the browser's maker, usually Google or
Microsoft, to be transcribed. We do not keep the audio ourselves.

*If voice typing is forbidden:* Our staff type notes. We do not use dictation,
so nothing you tell us is recorded as audio or sent anywhere to be transcribed.

Record the decision in the application so the software agrees with this notice:
`dpdpSetVoicePolicy("ALLOWED")` or `dpdpSetVoicePolicy("FORBIDDEN")`, from the
Privacy console.

### Where your information goes

Our records are kept in Google Workspace — Sheets, Drive and Apps Script —
which means Google processes them on our behalf under a contract with us.
Google stores data in data centres in several countries, so some of it is held
outside India.

If you have agreed to reports by WhatsApp, the message and the link go through
WhatsApp, run by Meta.

If you have agreed to insurance claims, your records go to the insurer or
third-party administrator that you name.

We do not sell your information, and we do not give it to anybody else unless
the law requires us to.

### How long we keep it

Your clinical record is kept for **three years from your last visit**, which is
what the National Medical Commission's regulations require of us. Inpatient and
medico-legal records are kept longer where local rules require it. Bills and
receipts are kept for **eight years**, which is what tax law expects. A mobile
number kept only to send you camp reminders is deleted **a year after you last
came**, or as soon as you tell us to stop.

The full schedule, with what each period is based on, is in
`docs/RETENTION_SCHEDULE.md` and we will show it to you if you ask.

### What you can ask us for

- **A copy of what we hold about you**, and a list of who we have given it to.
- **Correction** of anything wrong, and completion of anything missing.
- **Erasure.** We will do it unless the law requires us to keep the record —
  and if it does, we will tell you which record and for how long.
- **Withdrawal** of any consent above that was your choice. It is as easy as
  giving it: there is a switch on your page in the patient portal, or tell us
  at the desk.
- **To nominate somebody** to exercise these rights for you if you die or
  become unable to exercise them yourself.
- **To complain** — to the person named below first, and to the Data Protection
  Board of India after that.

We will answer within **30 days**. Before we send you anything we will contact
you on the number or email we already hold for you, to check that it is really
you asking. That check is there to protect you.

You can ask at the clinic, on the portal, or on this page:
**[web app URL]?privacy**

### Children

If you are under 18, we need your parent's or guardian's consent, and we record
who gave it. We do not track children or show them advertising, ever.

### Who to contact

**[Grievance officer name]**
[Email] · [Phone]

If you are not satisfied with our answer, you may complain to the **Data
Protection Board of India**.

*Notice version [v1-YYYYMMDD]. If we change what we do with your information we
will give you a new notice and ask you again.*

---

## Notes for the clinic, not for the patient

**Where this notice must appear.** On the registration screen (it already does
— `Admin_DPDP_Console` renders it from `getDPDPNotice()`), on the wall at
reception in both languages, and at `?privacy` on the web app. Print that URL
on the registration slip and in the invoice footer; a notice nobody can find
after the day they registered is a notice given once and then withdrawn.

**Two open decisions.** The dictation paragraph above, and whether the clinic
uses WhatsApp for documents at all. Both change what this notice says and both
are business decisions, not engineering ones.

**This consent is now enforced, not just recorded.** Every route that sends a
patient a document — the pharmacy invoice, the lab invoice, the lab report, the
archived invoice and the OP prescription, each by WhatsApp and by email — checks
this purpose on the server before it sends, and refuses without it
(`DPDP_Dispatch.gs`). Until this change the purpose was asked at the desk,
recorded, withdrawable in the portal, and read by nothing. If the clinic edits
the paragraph above, the words at the counter change with it: the text the desk
reads out before asking is `DPDP_DISPATCH.CHANNELS[…].notice`, and it has to
say the same thing as this notice.

**What makes the consent provable.** `Consent_Register` is append-only and each
row carries the notice version it was given against. That is what answers
section 6(10) — the burden of proving consent is on the clinic — so never edit
that sheet by hand. A correction is another row, which is what
`withdrawConsent()` and `recordConsent()` already do.
