#!/usr/bin/env bash
# ============================================================================
# Static checks for this Apps Script project.
#
# There is no build step here and no test runner - the code is pasted into an
# Apps Script editor, where a typo is discovered by a user. These eight checks
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

hr "2. Inline handlers and element ids that resolve to nothing"
node tools/audit.js

hr "3. google.script.run calls with no matching .gs function"
node tools/rpc.js

hr "4. Server calls with no failure handler (they fail silently)"
node tools/nofail.js && echo "(nothing listed above = every call has a failure path)"

hr "5. Deployment_Check.gs DEP_MAP vs the functions that actually exist"
node tools/dep.js

hr "6. Auth.html CRESC_SYMBOL_HOME vs the same"
node tools/sym.js

hr "7. Endpoints the browser can call with no permission check"
node tools/rbac.js | head -n 4

hr "8. CSS classes used but never defined"
echo "   (review by hand — template literals produce false positives)"
node tools/css.js | tail -n 20

hr "Also run from inside the Apps Script editor"
cat <<'NOTE'
  cresc_testDates()        Shared_Dates.gs      the date parser's own checks
  verifyDeployment()       Deployment_Check.gs  which .gs files are missing
  normaliseSheetDates()    IP_Schema_Repair.gs  dry run: text dates in sheets
  dpdpReadinessCheck()     DPDP_Compliance.gs   DPDP posture of this deployment
  crescRbacCoverage()      RBAC.gs              the same count as check 7
NOTE

exit "${FAILED:-0}"
