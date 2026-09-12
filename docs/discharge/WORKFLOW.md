# Discharge — the operational workflow

One patient leaves the hospital once, so there is one workflow, and it has an
order. This document is what the screens now enforce.

---

## The shape of it

```
  ADMITTED
     │
     │  ①  PREPARE       ward / doctor        Discharge Desk
     │        write the summary from the case sheet, notes and orders
     ▼
  GENERATED ──► IN_PREPARATION ──► PENDING_SIGNATURE
     │                                   │
     │                                   │ returned for correction ──┐
     │  ②  SIGN         consultant       ▼                           │
     │        verify and sign                                        │
     ▼                                                               │
   SIGNED ◄───────────────────────────────────────────────────────────┘
     │
     │  ③  SETTLE       accounts          Discharge Settlement
     │        post ward charges, apply package / insurance, take payment
     ▼
  DISCHARGED   ← the admission closes and the bed is released HERE
```

Stage ③ is the only stage that ends an admission. Stages ① and ② are clinical
and produce a document; stage ③ is financial and produces a receipt, and it is
the receipt that frees the bed.

---

## Where each stage happens

| Stage | Who | Screen | What it writes |
|---|---|---|---|
| ① Prepare | doctor, nurse, ward clerk | **EMR → Discharge Desk**, or the Discharge button on any ward list | `DS_Summaries`, `DS_Working`, `DS_Snapshots` |
| ② Sign | consultant | **Discharge Desk → the summary editor → Verify & sign** | a frozen `SIGNED` snapshot + hash chain |
| ③ Settle | accounts, admin | **Finance Hub → Discharge Settlement** | `IP_Settlements`, ledger receipts, insurance claims, then `IP_Admissions.Status = DISCHARGED` and the bed to `Cleaning` |

### One button, not four

Every ward screen — IP Admissions and the IP Ledger — has a single **Discharge**
button. It opens the discharge panel (`DS_Discharge_Flow.html`), which shows all
three stages, where this patient has got to, and the one action that is next.

Before this, four buttons on three screens each ran a complete discharge path of
their own, two of them ending the admission outright, and nothing on screen said
which one raised a bill. The IP Ledger's copy did not even handle the summary
gate, so with the gate on, every discharge from that screen ended on "Discharge
failed." with no way to explain why the patient was going without a summary.

---

## The exception: closing without settling

Step 3 of the panel is **Close the admission** — the ward's direct path, which
frees the bed with no bill raised. It is marked as the exception because that is
what it is. Use it only when there is genuinely nothing to settle:

* a death,
* an absconded patient,
* a stay that was already settled.

It writes the same `DISCHARGED` status and releases the same bed; it simply
posts no charges. It is audited as itself.

---

## The summary gate

`DS_BILLING_GATE` (a Script Property) decides what happens when someone tries to
discharge a patient whose summary is not signed. Both discharge paths —
`processPatientDischarge()` and `settleDischarge()` — go through the same check,
`dsx_gateCheck_()`, so gating one and not the other would be decorative.

| Mode | Behaviour |
|---|---|
| `OFF` | no check at all — the rollback switch, one property, no data change |
| `WARN` *(default)* | refused once with `DS_NOT_SIGNED`; proceeds when the caller supplies a reason, which is written to **both** the clinical and the accounts audit trails |
| `BLOCK` | refused outright — except for `DEATH` and `ABSCONDED`, which always behave as `WARN`. A death certificate cannot wait on a typed summary, and blocking an absconded patient only teaches staff to route around the system |

The panel shows the current mode, so nobody discovers it by being refused.

---

## Discharge types

`NORMAL`, `LAMA`, `DAMA`, `REFERRED`, `DEATH`, `ABSCONDED`. The type is chosen
when the summary is initiated; it changes the document title, the sections the
readiness check insists on (LAMA and DAMA need the risks-explained block, a
death needs the certificate details), and how the gate treats the discharge.

---

## What the states mean

| Status | Meaning |
|---|---|
| `GENERATED` | assembled from the record, nobody has edited it yet |
| `IN_PREPARATION` | someone is editing it |
| `PENDING_SIGNATURE` | submitted; waiting on a consultant |
| `RETURNED` | a consultant sent it back with section comments |
| `SIGNED` | frozen, hashed, printable, verifiable by QR |
| `AMENDMENT_IN_PROGRESS` | a signed summary is being amended; the signed version still stands until the amendment is signed |
| `CANCELLED` | abandoned; a new one can be initiated |

A signed summary is never rebuilt from current data. If its stored snapshot
cannot be read, that is reported as the integrity failure it is rather than
papered over with a fresh assembly — the record has moved since it was signed,
and a rebuilt document carrying someone's signature would be worse than none.

---

## If the summary will not load

`ds_getSummary` can fail in a way that looks like a dead connection, because
`google.script.run` calls the *success* handler with `null` whenever the reply
cannot be carried back. The editor no longer treats that as an answer:

1. it retries the same call once,
2. then falls back to `ds_getSummaryLite` — the document without the workflow
   timeline or the readiness panel, which is always small enough to arrive,
3. and only then reports anything, by which point an expired sign-in is a real
   possibility rather than a guess.

On the server, `dsx_fitForWire_()` measures the reply and sheds the advisory
parts rather than returning a document too big to arrive, and says on screen
what it left out. `dsx_wire_()` removes the values the transport cannot carry
(an Invalid Date, a `NaN`, a cycle) before anything is returned.
