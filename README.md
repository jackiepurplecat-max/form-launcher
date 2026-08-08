# HelpfulForms - Secure Configuration Setup

This project now uses environment variables to keep your API keys and secrets secure!

## Quick Start

### 1. Build the HTML file
```bash
npm run build
```

This reads from `form-launcher/.env` and injects the variables into `form-launcher/index.html`.

### 2. Push to Google Apps Script
```bash
npm run clasp:push
```

### 3. Set up Script Properties (ONE TIME ONLY)

1. Open your Apps Script project:
   ```bash
   npm run clasp:open
   ```

2. In the editor, run the `setupScriptProperties()` function:
   - Select **setupScriptProperties** from the function dropdown
   - Click **Run** (▶️)
   - This stores your API key securely in Script Properties

3. (Optional) Delete or comment out the `setupScriptProperties()` function after running once

---

## Configuration Files

### `.env` file (form-launcher/.env)
Contains all your secrets - **NEVER commit to git!**

```env
DELETE_API_KEY=your-secret-key-here
DELETE_WEBAPP_URL=your-web-app-url-here
SHEETS_API_KEY=your-sheets-api-key-here
SPREADSHEET_ID=your-spreadsheet-id-here
RECIPIENT_EMAIL=your-email@example.com
ICLOUD_EMAIL=your-address@icloud.com
FORM_ID=your-work-expenses-form-id
```

**To update:** Edit `.env`, then run `npm run build`

**Note:** only the first four are injected into the frontend. `RECIPIENT_EMAIL`,
`ICLOUD_EMAIL` and `FORM_ID` are used by Code.js and live in Script Properties.

| Property | Required | Used for |
|---|---|---|
| `DELETE_API_KEY` | yes | Key the web app checks on every POST |
| `RECIPIENT_EMAIL` | yes | Where Work expense receipts are emailed |
| `FORM_ID` | yes | Work expenses form, for dropdown add/remove |
| `ICLOUD_EMAIL` | no | Health iCloud rename Shortcut (Health section only) |

### `index.template.html`
Template with placeholders like `{{DELETE_API_KEY}}` - safe to commit to git.

### `index.html` (GENERATED)
Built automatically from template - **DO NOT edit directly!** This file is gitignored.

---

## Deployment Workflow

### First Time Setup

1. **Enable Apps Script API**: https://script.google.com/home/usersettings

2. **Clone your script** (already done):
   ```bash
   npm run clasp:clone <your-script-id>
   ```

3. **Update .env** with your values (especially `RECIPIENT_EMAIL`!)

4. **Build and push**:
   ```bash
   npm run build
   npm run clasp:push
   ```

5. **Set up Script Properties**:
   - Open Apps Script: `npm run clasp:open`
   - Fill in `SCRIPT_PROPERTY_VALUES` near the top of `Code.js` with the values
     from your `.env`. Leave any you don't want to change blank — blanks are
     skipped, so re-running never overwrites a real value with a placeholder.
   - Select `setupScriptProperties` from the function dropdown, click Run ▶️
   - The log lists what was written and then reports the state of every property
   - **Clear `SCRIPT_PROPERTY_VALUES` again afterwards** so secrets aren't left
     sitting in the script body
   - You can run `checkScriptProperties` at any time to see what's configured
     without changing anything — useful after moving to a different account

6. **Deploy as Web App**:
   - `npm run clasp:open`
   - **Deploy** → **New deployment**
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Copy the Web App URL

7. **Update .env** with `DELETE_WEBAPP_URL`

8. **Rebuild**:
   ```bash
   npm run build
   ```

---

## After Making Changes

### Frontend changes (index.template.html):
```bash
npm run build
```

### Backend changes (Code.js):
```bash
npm run clasp:push
```

### Config changes (.env):
```bash
npm run build
```

---

## Security Features

✅ **Secrets stored in .env** - Not in source code
✅ **.env is gitignored** - Won't be committed
✅ **index.html is gitignored** - Generated from template
✅ **Script Properties** - API key stored securely in Apps Script
✅ **Build script** - Injects secrets only when needed

---

## Useful Commands

| Command | Description |
|---------|-------------|
| `npm run build` | Build index.html from template + .env |
| `npm run clasp:push` | Push Code.js to Google Apps Script |
| `npm run clasp:pull` | Pull latest from Google Apps Script |
| `npm run clasp:open` | Open script in browser |
| `npm run clasp:status` | Check sync status |

---

## Troubleshooting

**Build warnings about missing variables?**
- Check that all required variables are set in `form-launcher/.env`
- Required: `DELETE_API_KEY`, `DELETE_WEBAPP_URL`, `SHEETS_API_KEY`, `SPREADSHEET_ID`, `RECIPIENT_EMAIL`

**Not sure which properties are configured?**
- Run `checkScriptProperties` in the Apps Script editor. It reports every
  property as set / missing / optional-and-unset, and never logs the API key
  back out in full.

**Health "Done" not renaming files in iCloud?**
- `ICLOUD_EMAIL` is likely unset — the send fails and is only logged, so the
  status still flips and the Drive files are still renamed
- Run `checkScriptProperties` to confirm, and see `health-rename-shortcut.md`

**Archive function not working?**
- Make sure you ran `setupScriptProperties()` in Apps Script
- Verify `DELETE_WEBAPP_URL` is set to your deployed Web App URL
- Check that both frontend and backend have the same `DELETE_API_KEY`

**"Invalid API key" error?**
- Frontend and backend keys don't match
- Run `setupScriptProperties()` again in Apps Script
- Rebuild with `npm run build`

---

## What's Protected

These files contain secrets and are **gitignored**:
- `form-launcher/.env` - Your configuration
- `form-launcher/index.html` - Generated file with secrets
- `.clasprc.json` - Your Google authentication

These files are **safe to commit**:
- `form-launcher/index.template.html` - Template with placeholders
- `Code.js` - Uses Script Properties (secrets not in code)
- `build.js` - Build script
- `.gitignore` - Git ignore rules
