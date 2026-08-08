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

These track the **current** filename of the iCloud copies — the name the
Shortcut has to search for.

On form submit they're set to the original upload names (with the ` - Username`
suffix Google Forms appends stripped off). After a successful Done they're
updated to the new names, because the Shortcut has just renamed those files.
Undo leaves them alone, since no email is sent and iCloud is untouched.

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

## Done → Undo → Done

This cycle is safe. Worked example:

| Step | Drive | iCloud | Column M |
|---|---|---|---|
| After submit | `250115_J_Dentist_50_receipt.pdf` | `IMG_1234.pdf` | `IMG_1234.pdf` |
| Done | `Claimed (15-01-2025) 250115_J_…_receipt.pdf` | `250115_J_…_receipt.pdf` | `250115_J_…_receipt.pdf` |
| Undo | `250115_J_…_receipt.pdf` | unchanged | unchanged |
| Done again | `Claimed (…) 250115_J_…_receipt.pdf` | unchanged | unchanged |

The second Done emails `ORIGINAL_RECEIPT` and `NEW_RECEIPT` as the same value,
so the Shortcut performs a no-op rename. That's expected.

If the Shortcut email can't be sent (e.g. `ICLOUD_EMAIL` unset), columns M and N
are deliberately **not** updated — they keep pointing at the real iCloud names
so a later retry still finds the files.

---

## Troubleshooting

| Symptom | Cause |
|---|---|
| No email at all | `ICLOUD_EMAIL` not set, or column M empty for that row |
| Email arrives, nothing renamed | Folder path in **Get File** doesn't match where the files actually are |
| Receipt renamed, details skipped | Expected when there's no details file — `ORIGINAL_DETAILS` is empty |
| `ORIGINAL_RECEIPT` = `NEW_RECEIPT` | Expected on a repeat Done — the Shortcut does a no-op rename |
| Columns M/N drifted from the real iCloud names | Someone renamed the files outside the Shortcut; correct M/N by hand |
| Automation never fires | "Ask Before Running" is still on, or the sender filter doesn't match |
