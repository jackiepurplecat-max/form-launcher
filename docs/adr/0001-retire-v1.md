# 1. Retire v1

Date: 15 Nov 2026 · Status: accepted

## Context

v1 (Google Forms + `onFormSubmit` + a GitHub Pages page reading the sheet via
the Sheets API, on the original Google account) was frozen at cutover and kept
only to work down its backlog (`REBUILD-PLAN.md` → *Cutover*, step 5). The owner
confirmed on 15 Nov 2026 that everything needed has been transferred to v2.

## Decision

- **Repo:** v1's files are deleted, not archived — `Code.js`, `index.html`,
  `build.js`, `appsscript.json`, the root `.claspignore`, `README.md` and
  `health-rename-shortcut.md`, plus the `build`/`dev`/`start`/`clasp:*` npm
  scripts. Git history keeps them; the last commit holding them is the parent
  of the commit that adds this ADR.
- **GitHub Pages stays on**, serving only `favicon.png`, because v2's
  `HtmlOutput.setFaviconUrl()` needs a public HTTPS address for its icon
  (`.scratch/v1-branding/`). The v1 page itself goes with `index.html`.
- **Google side, done by hand by the owner:** stop the four v1 forms accepting
  responses, delete the v1 web app deployment and its trigger, delete the v1
  `ICLOUD_EMAIL` Script Property and the iOS rename automation. The v1 sheet and
  Drive folder are **kept as the archive**, not deleted.

## Consequences

- The `clasp:*` / default-credential trap no longer exists in this repo; every
  clasp command is `--user v2`.
- `index.html` no longer publishes two NIFs on the live Pages site. They remain
  in git history; the 13 Aug decision to leave that is unchanged, and making the
  repo private would now only cost the favicon's host.
- Local untracked v1 leftovers (`index.template.html`, `form-launcher/`, root
  `.clasp.json`) are not touched by this change; they stay gitignored.
