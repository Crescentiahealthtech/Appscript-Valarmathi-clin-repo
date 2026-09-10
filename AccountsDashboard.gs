// AccountsDashboard.gs — RETIRED (superseded by Accounts_Dashboard.gs)
//
// Every function that used to live here was a byte-for-byte older copy of one in
// Accounts_Dashboard.gs, which additionally counts saved discharge drafts as open
// tabs. Apps Script has one global scope: with both files present, whichever loaded
// last silently won, so the finance dashboard reported different open-tab totals
// depending on file order. The bodies were removed rather than the file, so the
// history stays in git. DELETE THIS FILE from the Apps Script project.
