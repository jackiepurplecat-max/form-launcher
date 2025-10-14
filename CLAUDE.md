## HelpfulForms

Google Apps Script-based expense tracking system that integrates Google Forms, Google Sheets, and automated file management with email notifications.

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

**index.html (lines 208-215):**
```javascript
SPREADSHEET_ID = '1J4dldGjb3SktoVxQD3BKm-iX6RDAF7YjoLImm7-MIsg'
SHEET_NAME = 'Travel!A:G'
SHEETS_API_KEY = 'AIzaSyCkQIbYsNr6ooub1lJysGoSflG9bka0fbs'
DELETE_WEBAPP_URL = 'YOUR_WEB_APP_URL_HERE'   // Apps Script web app URL
DELETE_API_KEY = 'YOUR_SUPER_SECRET_KEY_HERE' // Must match Code.js
```

**Code.js (line 9):**
```javascript
recipient = "you@example.com"  // Email recipient for expense notifications
```

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
