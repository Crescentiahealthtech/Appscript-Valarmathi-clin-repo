# tools/ — checks that run on your computer, not in Apps Script

**Do not paste anything in this folder into the Apps Script editor.**

These files are Node.js scripts. They read the `.gs` and `.html` files in this
repository as *text* and report problems in them. They are not part of the
application and they never run on Google's servers.

Two things would go wrong if they were pasted into the Apps Script project:

- They use `require()`, `__dirname`, `process.env` and `process.exit`, none of
  which exist in Apps Script. The file would fail the moment the project loaded.
- `dep.js` declares a global called `DEP_MAP`. So does `Deployment_Check.gs`,
  where it is the real list of which function lives in which file. Apps Script
  has one global scope shared by every file, so pasting `dep.js` would replace
  that list with a copy of itself and `verifyDeployment()` would stop being able
  to tell you which file is missing.

## Running them

From the repository root, with Node installed:

```
./tools/check.sh
```

That runs all seven and prints a report. Only the syntax check sets the exit
code, so it is safe to use in a pre-commit hook or a CI job.

Individually, if you want just one:

```
node tools/validate.js     # syntax of every .gs file and every inline <script>
node tools/audit.js        # onclick= and getElementById that resolve to nothing
node tools/rpc.js          # google.script.run calls with no server function
node tools/nofail.js       # server calls with no failure handler
node tools/dep.js          # DEP_MAP vs the functions that actually exist
node tools/sym.js          # CRESC_SYMBOL_HOME vs the same
node tools/css.js          # classes used in markup that no stylesheet defines
```

`css.js` needs a human eye — template literals produce false positives — so it
prints its findings and never fails the run.

## What each one is for

Every check exists because the thing it looks for was actually found in this
codebase, not because it seemed like a good idea:

| Script | Found |
|---|---|
| `validate.js` | the baseline — nothing ships that does not parse |
| `audit.js` | `#avail-date` and `#hd-doctor-badge`, two element ids nothing declared |
| `rpc.js` | nothing, and that is the point: all 154 endpoints resolve |
| `nofail.js` | eleven calls with no failure handler, including two patient-portal buttons that stuck on "Saving…" for ever |
| `dep.js` | `Drug_Interactions.gs`, a filename `verifyDeployment()` told people to look for that has never existed |
| `sym.js` | the same typo in the client-side copy of that map |
| `css.js` | nine classes in use with no rule anywhere, including `premium-card` on every Accounts panel |

## The three checks that *do* run inside Apps Script

These are `.gs` functions in the project proper. Run them from the editor's
function dropdown after pasting files across:

- `cresc_testDates()` — `Shared_Dates.gs`. Twenty checks on the date parser.
  Run it if you ever touch that file.
- `verifyDeployment()` — `Deployment_Check.gs`. Names the `.gs` files that are
  missing or out of date in the project. Run it first whenever the app behaves
  strangely after an update.
- `normaliseSheetDates()` — `IP_Schema_Repair.gs`. Converts text date cells in
  the workbook to real dates. **Dry run by default** — call it with no argument,
  read the report it returns, then call `normaliseSheetDates(true)` to apply.
