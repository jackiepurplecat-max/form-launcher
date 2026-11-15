# HelpfulForms

A personal expense, IVA, health-claim and income tracker, built as a Google
Apps Script web app over one Google Sheet.

- The code is in `v2/`. Start with `CLAUDE.md`, then `NEXT-SESSION.md` and
  `REBUILD-PLAN.md`.
- `favicon.png` is served by GitHub Pages so the web app can use it as its icon.
- The original v1 system was retired on 15 Nov 2026 (`docs/adr/0001-retire-v1.md`);
  it is in git history.

## Local setup

```bash
npm install
npm run v2:login     # once — the named `v2` clasp credential
npm run v2:test      # local harness
```

Secrets live in `.env` at the repo root (gitignored), never in a tracked file:

| Key | Used for |
|---|---|
| `V2_CLASP_ACCOUNT` | Which Google account v2 uses |
| `V2_DEPLOYMENT_ID` | The `/exec` deployment `npm run v2:deploy` moves |
| `V2_IVA_CLAIM_RECIPIENT` | Where IVA claims are emailed |
| `V2_COMPLETION_EMAIL_RECIPIENT` | Where completion requests are emailed |
| `V2_SIRI_API_KEY` | Shared key for the Siri endpoint |

Apps Script secrets live in Script Properties, set once with
`setupScriptProperties()` — see `REBUILD-PLAN.md`.
