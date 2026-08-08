# Health Claim iCloud Rename Shortcut

When you press **✓ Done** on a Health claim in the web app, Apps Script emails a
machine-readable rename instruction to your iCloud address. This Shortcut reads
that email and renames the matching files in iCloud Drive.

> The formats below are taken from `toggleHealthClaimStatus()` in `Code.js`.
> If you change the filename pattern there, update this file too.

---

## Prerequisites

**1. Script property `ICLOUD_EMAIL`**

The email is sent to whatever `getIcloudEmail()` returns. Note that
`setupScriptProperties()` does **not** set this one — you must add it yourself:

Apps Script editor → **Project Settings** → **Script Properties** → Add:

| Property | Value |
|---|---|
| `ICLOUD_EMAIL` | your @icloud.com address |

Without it, pressing Done logs `ICLOUD_EMAIL not configured` and sends nothing.
The status still flips and the Drive files are still renamed, so the failure is
silent from the UI.

**2. Health sheet columns M and N**

| Column | Header |
|---|---|
| M | `Original Receipt Filename` |
| N | `Original Details Filename` |

These are filled in on form submit with the *original* upload names (with the
` - Username` suffix Google Forms appends stripped off). The Shortcut needs them
because the iCloud copies still carry those original names.

---

## The email

**Sent only when marking a claim Done** — not on Undo — and only if column M is
non-empty.

**Subject:**

```
Health Claim Rename: <patient initial> <yyMMdd> <provider first word> <amount>
```

Example: `Health Claim Rename: J 250115 Dentist 50`

**Body** — four `KEY=value` lines, nothing else:

```
ORIGINAL_RECEIPT=IMG_1234.pdf
NEW_RECEIPT=250115_J_Dentist_50_receipt.pdf
ORIGINAL_DETAILS=Invoice_567.pdf
NEW_DETAILS=250115_J_Dentist_50_details.pdf
```

If there's no details file, the last two lines are still present but empty:

```
ORIGINAL_DETAILS=
NEW_DETAILS=
```

### How the new name is built

`<yyMMdd>_<patient initial>_<provider first word>_<amount>_receipt<.ext>`

| Part | Source |
|---|---|
| `yyMMdd` | Column F (Invoice Date) |
| patient initial | First character of the first word of column B (Patient) |
| provider first word | First word of column D (Provider) |
| amount | Column I |
| `.ext` | Extension of the original filename in column M / N |

The details file is identical but ends `_details` instead of `_receipt`.

> **Note:** the copies in *Google Drive* additionally get a
> `Claimed (DD-MM-YYYY) ` prefix. The iCloud names deliberately do **not** —
> `NEW_RECEIPT` is the unprefixed name.

---

## Shortcut: "Health Claim Rename"

### Actions

**1. Find Mail Messages**
- Subject **contains** `Health Claim Rename`
- Sort by **Date Received**, Latest First
- Limit **1**

**2. Get Details of Messages**
- Get **Body** — from the Mail Messages above
- Set Variable: `Body`

**3. Extract the four values**

For each of the four keys, add a **Match Text** action against `Body`, then a
**Get Group from Matched Text** (Group 1), then **Set Variable**:

| Match Text pattern | Set Variable |
|---|---|
| `ORIGINAL_RECEIPT=(.+)` | `OriginalReceipt` |
| `NEW_RECEIPT=(.+)` | `NewReceipt` |
| `ORIGINAL_DETAILS=(.+)` | `OriginalDetails` |
| `NEW_DETAILS=(.+)` | `NewDetails` |

`(.+)` requires at least one character, so when there's no details file the
details variables come back empty — which is exactly what step 5 checks for.

**4. Rename the receipt**
- **Get File** from iCloud Drive (point it at the folder your receipts land in)
- **Filter Files** where Name **is** `OriginalReceipt`
- **If** Files has any value
  - **Rename File** → Filtered Files → New Name: `NewReceipt`
- **End If**

**5. Rename the details file, if there is one**
- **If** `OriginalDetails` **has any value**
  - **Get File** from iCloud Drive
  - **Filter Files** where Name **is** `OriginalDetails`
  - **If** Files has any value
    - **Rename File** → New Name: `NewDetails`
  - **End If**
- **End If**

**6. Show Notification**
- Title: `Health Claim Files Renamed`

### Running it automatically

Shortcuts → **Automation** → **+** → **Create Personal Automation** → **Email**

- **Sender contains** your Gmail address
- **Subject contains** `Health Claim Rename`
- Action: **Run Shortcut** → "Health Claim Rename"
- Turn **off** "Ask Before Running"

---

## Known issue: Done → Undo → Done

Undoing a claim overwrites columns M and N with `Claimed/<new name>` rather than
restoring the original upload names. So on a *second* Done, the email sends:

```
ORIGINAL_RECEIPT=Claimed/250115_J_Dentist_50_receipt.pdf
```

— a path, not a filename, which won't match anything in iCloud. Until that's
fixed, if you undo a claim and re-do it, correct columns M and N by hand first.

---

## Troubleshooting

| Symptom | Cause |
|---|---|
| No email at all | `ICLOUD_EMAIL` not set, or column M empty for that row |
| Email arrives, nothing renamed | Folder path in **Get File** doesn't match where the files actually are |
| Receipt renamed, details skipped | Expected when there's no details file — `ORIGINAL_DETAILS` is empty |
| `ORIGINAL_RECEIPT` looks like `Claimed/...` | See "Known issue" above |
| Automation never fires | "Ask Before Running" is still on, or the sender filter doesn't match |
