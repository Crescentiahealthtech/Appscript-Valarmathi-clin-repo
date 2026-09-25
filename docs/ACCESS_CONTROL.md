# Access control: roles, per-person access, and what changed

Every screen and every server call in CresRx is gated by a **permission** (for example `lab.bill` or `pharmacy.register`). A person's permissions are their **role's** permissions, plus anything they were **given individually**, minus anything **taken away** from them individually. The server checks this on every call (`crescRequire_` in `RBAC.gs`). What the browser hides is only a convenience.

## Managing one person's access

Go to **Admin Dashboard → Staff Accounts**. The **Access** column shows *role default* or *+N / −M custom*.

- **change** opens that person's access. Each module (Command Center, Patients, Appointments, Clinical, Ward, Laboratory, Pharmacy, Hospital billing, Accounts, Administration) can be switched on or off as a whole, or function by function. **added** and **removed** mark where this person differs from their role. **Reset to role defaults** clears their custom access.
- **Role** can be changed between Nurse, Receptionist, Pharmacist, Lab technician and Accountant. The system owner can also choose Administrator. A role change signs the person out, so their screens are redrawn for the new role. A doctor's role is not changed here: switch the account off and add the person again.
- **Who can open what** shows every account against every module. A ringed cell differs from the role default. Click a name to change that person's access.
- Every save is recorded in the Audit Log as `STAFF_ACCESS_CHANGED`, with the before and after and the reason typed in.
- Changes take effect on the person's **next click**. Their menu redraws within five minutes, or when their window regains focus.

The same lists are stored on the **Users** sheet in `Access_Grant` and `Access_Revoke` (comma-separated keys). You can edit them by hand; `crescRbacSelfTest()` reports any key typed wrongly.

### The rules the server enforces

| Rule | Why |
|---|---|
| Only `admin.users` holders (administrators) can open the Access screen | It creates and removes power |
| Doctors' and administrators' access is changed by the **system owner** only (the `Super_Admin` account) | Same boundary as creating or resetting them |
| Nobody can change their **own** access | An administrator who could widen their own account would need no owner |
| `admin.users`, `admin.config`, `admin.audit`, `dpdp.manage` and `accounts.lock_period` can be given by the owner only | Each reaches beyond one person's own work |
| The owner's key (`admin.users.elevated`) is never grantable | Ownership is the `Super_Admin` flag on the Users sheet |

## Account state is checked on every call

- **Switch off** (Staff Accounts) takes effect immediately. The person's next click is refused and all their open sessions end. Before this change, a switched-off account kept working until its session expired.
- A **password reset** by an administrator also ends the account's open sessions.
- A **role** typed differently on the Users sheet applies on the person's next call.
- Sessions have a sliding 8-hour lifetime and an absolute 16-hour cap.

## What changed for each role in this release

| Role | Change | If a person needs the old behaviour |
|---|---|---|
| Nurse, Receptionist, Pharmacist, Lab technician, Accountant | **Command Center** (`dashboard.read`) is no longer part of the role. They land on their own desk: EMR, Appointments, Pharmacy, Lab or Accounts. | Give them *Command Center* in Access |
| Pharmacist | `billing.read` / `billing.write` replaced by **`pharmacy.bill`** (settle, search, reprint pharmacy bills) and **`pharmacy.register`** (Schedule H register). A pharmacist can no longer take payments on lab or hospital bills. | Give *Raise or take payment on a hospital bill* if they also run the hospital counter |
| Lab technician | `billing.read` / `billing.write` replaced by **`lab.bill`**. A lab technician can no longer settle pharmacy bills. Voiding a lab bill **that has money on it** now needs `billing.cancel`, which administrators and accounts hold. An unpaid bill can still be voided by the lab. | Give *Cancel or void a bill* only to a lab lead you trust with refunds |
| Receptionist | Unchanged, except the Command Center | Give *Raise, collect on and reprint lab bills* if reception bills lab tests |
| Accountant | Gains **`billing.cancel`** (cancel hospital invoices, void paid lab bills) | — |
| Administrator | Gains every new key automatically | — |

The cash drawer works at every counter for anyone who holds that counter's billing key: `billing.write`, `pharmacy.bill` or `lab.bill`.

## Permission list

Run `crescRbacSelfTest()` from the Apps Script editor after any hand edit to roles or to the Access columns. `node tools/access.js` (part of `tools/check.sh`) exercises the rules above against a pretend Users sheet.
