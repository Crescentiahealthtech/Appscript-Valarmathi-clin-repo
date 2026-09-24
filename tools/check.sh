#!/usr/bin/env bash
# ============================================================================
# Static checks for this Apps Script project.
#
# There is no build step here and no test runner - the code is pasted into an
# Apps Script editor, where a typo is discovered by a user. These checks
# are what can be verified without a Google account, and they each exist
# because the thing they look for was actually found in this codebase.
#
#   ./tools/check.sh
#
# Requires node. Exits non-zero if the syntax check fails.
# ============================================================================
set -u
cd "$(dirname "$0")/.."

hr() { printf '\n\033[1m%s\033[0m\n' "$1"; }

hr "1. Syntax — every .gs file and every inline <script>"
node tools/validate.js || FAILED=1

hr "2. Script text the HTML parser would cut short"
node tools/hazard.js || FAILED=1

hr "3. Inline handlers and element ids that resolve to nothing"
node tools/audit.js

hr "4. google.script.run calls with no matching .gs function"
node tools/rpc.js

hr "5. Server calls with no failure handler (they fail silently)"
node tools/nofail.js && echo "(nothing listed above = every call has a failure path)"

hr "5b. The same top-level name defined in two .gs files (last one wins)"
node tools/dupes.js || FAILED=1

hr "6. Deployment_Check.gs DEP_MAP vs the functions that actually exist"
node tools/dep.js

hr "7. Auth.html CRESC_SYMBOL_HOME vs the same"
node tools/sym.js

hr "8. Endpoints the browser can call with no permission check"
node tools/rbac.js

hr "9. Calls that do not pass the session token their endpoint asks for"
node tools/token.js

hr "10. The password code, exercised against Node's own HMAC"
node tools/credtest.js || FAILED=1

hr "10b. Sign-in flows: reset, MFA, Google, against a pretend spreadsheet"
node tools/authflow.js || FAILED=1

hr "11. CSS classes used but never defined"
echo "   (review by hand — template literals produce false positives)"
node tools/css.js | tail -n 20

hr "Also run from inside the Apps Script editor"
cat <<'NOTE'
  cresc_testDates()        Shared_Dates.gs      the date parser's own checks
  verifyDeployment()       Deployment_Check.gs  which .gs files are missing
  normaliseSheetDates()    IP_Schema_Repair.gs  dry run: text dates in sheets
  dpdpReadinessCheck()     DPDP_Compliance.gs   DPDP posture of this deployment
  crescRbacCoverage()      RBAC.gs              the same count as check 7
  crescRbacSelfTest()      RBAC.gs              the role matrix against itself
  crescCredentialStatus()  Auth_Credentials.gs  hashed vs plain-text passwords
  crescPwdBenchmark()      Auth_Credentials.gs  what hashing costs on this runtime
  dpdpTriggerStatus()      DPDP_Triggers.gs     whether the scheduled jobs exist
NOTE

hr "Set up once, on a new deployment"
cat <<'NOTE'
  dpdpSetup()              create every DPDP register
  dpdpSetGrievanceOfficer(name, email, phone)      section 13
  dpdpInstallTriggers()    the daily, weekly and monthly jobs
  crescMigrateCredentials()  hash every password and force a reset
  See docs/DPDP_READINESS.md, "Do these five things this week".
NOTE

exit "${FAILED:-0}"
