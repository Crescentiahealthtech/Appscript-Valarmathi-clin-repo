# The debrief, and how to regenerate it

`DPDP_Debrief.html` is the source of the PDF handed to the clinic owner and
the lawyer: what was wrong, what changed, what is still theirs to do, and
where every function lives.

It is kept as HTML rather than as a PDF because a PDF in a repository cannot
be diffed, and this document goes stale the moment the code moves. The three
things in it that drift are the **line numbers in §5**, the **commit id on the
cover**, and the **counts** in §2 and §10. Everything else is prose.

## Regenerate

Any headless Chromium will do:

```bash
chromium --headless --disable-gpu --no-pdf-header-footer \
  --print-to-pdf=DPDP_Remediation_Debrief.pdf \
  file://"$PWD"/docs/debrief/DPDP_Debrief.html
```

The page setup (A4, margins, the full-bleed cover) is in the file's own
`@page` rules, so nothing needs to be passed on the command line.

## Before regenerating, refresh the numbers

```bash
node tools/rbac.js        # the counts in §2 and §10
node tools/token.js       # the call-site figure in §2
grep -n "^function NAME(" *.gs   # the line numbers in §5
git rev-parse --short HEAD       # the commit on the cover
```

If a line number in §5 is wrong, the file name beside it is still right —
that is why both are printed, and why the section says to search rather than
to scroll.
