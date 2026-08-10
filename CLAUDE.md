# This repository holds TWO systems

Read this before anything else. Most of this file describes **v1**, which is
frozen. New work happens in **v2**, which is a different design on a different
Google account, and several things that are true of v1 are deliberately false
of v2.

| | v1 (frozen) | v2 (active) |
|---|---|---|
| Code | `Code.js` at the repo root, `index.html` | `v2/` |
| Account | the original one | a separate one — `V2_CLASP_ACCOUNT` in `.env` |
| Push | `npm run clasp:push` | `npm run v2:push` (`--user v2` credential) |
| Intake | four Google Forms + `onFormSubmit` | **no Forms, no triggers** — `createEntry()` only |
| UI | GitHub Pages + Sheets API + two client-side keys | Apps Script web app — `v2/Web.js` + `v2/Index.html` |

**Do not change v1.** It still serves the old account and is being worked down
to zero in parallel. Its known bugs are accepted, not fixed.

## v2 — the active system

`REBUILD-PLAN.md` is the source of truth: architecture, decisions and the
reasons behind them, the build order, and a state-of-play section saying exactly
what is done. Read it before starting work.

**Run the tests.** `npm run v2:test` executes the real `v2/` source against
stand-ins for Apps Script's services. Run it before you change anything, to
establish the baseline, and after, because it catches things review does not —
several of the bugs it now guards against were found that way rather than by
reading the code.

```
npm run v2:test    # the harness — do this first
npm run v2:check   # syntax only
npm run v2:push    # add --force if appsscript.json changed
```

**Things that are true of v2 and not of v1:**

- **Columns are resolved by header name, never by index.** No magic numbers.
- **Headers are generated from `SECTIONS`**, so adding a field is a config
  change plus a re-run of `bootstrap()`. Never type a header into a sheet.
- **Everything section-specific lives in `v2/Config.js`.** An
  `if (section === 'health')` anywhere else is a bug in the wrong place.
- **`createEntry()` is the only way a row is born.** There is no trigger.
- **Never report success for work that failed.** Operations return what
  actually happened; a status change that renames nothing says so.
- **Every value written to a sheet goes through `safeCellValue`.**
- **Claim mail sends once**, when the document is there, and at no other time —
  guarded by a `Claim Emailed` column, not by convention.
- **Every function the web page can call checks the caller itself.**
  `google.script.run` reaches any global in the project, so the deployment
  setting is never the only gate. The check reads `Session.getActiveUser()` — the
  visitor — not `getEffectiveUser()`, which under "execute as me" is always you
  and therefore proves nothing.

`v2/Smoke.js` is the live counterpart of the harness, run from the Apps Script
editor. `v2/test/` is local only and is not pushed.

---

## v1 — FROZEN. Do not change.

Everything below describes the old system. It is kept for reference while its
backlog is worked down. Google Apps Script-based expense tracking that
integrates Google Forms, Google Sheets, and automated file management with email
notifications.

### Project Structure

```
HelpfulForms/
├── form-launcher/
│   └── index.html          # Main UI - displays expenses, handles filtering/deletion
├── Code.js                 # Google Apps Script - handles form submissions
├── appsscript.json         # Apps Script project config
└── package.json            # Node.js config with clasp scripts
```

### Workflow

1. **User loads `index.html`**
   - Displays expense data from Google Sheet using Sheets API
   - Shows "Travel Expenses" section with interactive table

2. **User interactions:**

   **a. New Expense**
   - Clicks "New Expense" button → Opens Google Form (https://forms.gle/Efmbz5brKNyohqQe7)
   - User fills form with: Trip name, Expense Date, Amount, Currency, Description, File upload
   - Form submission → Updates Google Sheet → **Triggers `handleTravel()` in Code.js**

   **b. Refresh**
   - Reloads the page to fetch latest data from Google Sheet

   **c. Filter by Trip**
   - Dropdown filters table to show expenses for selected trip only
   - Client-side filtering (no server calls)

   **d. Delete Trip**
   - User enters trip name and clicks "Delete"
   - Sends POST request to Apps Script Web App
   - Deletes all rows matching the trip name from Google Sheet
   - Auto-refreshes table after deletion

3. **Script Automation (`Code.js` - `handleTravel()`)**
   - **Trigger:** Google Sheet edit (form submission)
   - **Process:**
     1. Checks if email already sent (column I)
     2. Extracts expense details from new row
     3. Renames uploaded file to: `YYYYMMDD_Description_Amount_Currency.ext`
     4. Sends email with:
        - Subject: "travel claim receipt [trip] [description]"
        - Attachment: Renamed file
        - Body: Expense details + Google Drive link
     5. Marks row as "Email sent" (column I = "Yes")
     6. Logs activity via `logRun()` function

### Configuration

This project uses environment variables for security. All secrets are stored in `.env` file (gitignored).

**`.env` file:**
```env
DELETE_API_KEY=your-secret-key-here
DELETE_WEBAPP_URL=your-web-app-url-here
SHEETS_API_KEY=your-sheets-api-key-here
SPREADSHEET_ID=your-spreadsheet-id-here
RECIPIENT_EMAIL=your-email@example.com
```

**Build process:**
- Run `npm run build` to inject .env values into index.html
- Template file: `index.template.html` (with {{placeholders}})
- Built file: `index.html` (with actual values, gitignored)

**Apps Script configuration:**
- Secrets stored in Script Properties (not in code)
- Run `setupScriptProperties()` once in Apps Script editor to configure
- See README.md for detailed setup instructions

### Development Setup

**Install clasp (Google Apps Script CLI):**
```bash
npm install
```

**Authenticate:**
```bash
npm run clasp:login
```

**Pull latest from Google:**
```bash
npm run clasp:pull
```

**Push changes to Google:**
```bash
npm run clasp:push
```

**Open in browser:**
```bash
npm run clasp:open
```

### Google Sheet Structure

**"Travel" sheet columns:**
- A: Trip name
- B: Expense Date
- C: (unused)
- D: Amount
- E: Currency
- F: Description
- G: File link/ID (from Google Form upload)
- H: (unused)
- I: Email sent? (Yes/blank)

### Features

- Real-time expense tracking with Google Sheets backend
- Automated file renaming with structured naming convention
- Email notifications with file attachments
- Trip-based filtering for expense organization
- Bulk deletion by trip name
- Mobile-responsive design with iOS PWA support
- Prevents duplicate email sends via status tracking

### Requirements

- Google account with access to:
  - Google Forms
  - Google Sheets
  - Google Drive
  - Google Apps Script
- Clasp CLI (installed via npm)
- Modern web browser
