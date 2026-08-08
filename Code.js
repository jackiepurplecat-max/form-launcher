// Configuration stored in Script Properties for security
// To set: Run setupScriptProperties() once, or set manually in Project Settings > Script Properties

function getDeleteApiKey() {
  const props = PropertiesService.getScriptProperties();
  let apiKey = props.getProperty('DELETE_API_KEY');

  // Fallback for initial setup - remove after setting Script Properties
  if (!apiKey) {
    Logger.log('⚠️ DELETE_API_KEY not found in Script Properties. Run setupScriptProperties() to configure.');
    throw new Error('DELETE_API_KEY not configured in Script Properties');
  }

  return apiKey;
}

function getRecipientEmail() {
  const props = PropertiesService.getScriptProperties();
  let email = props.getProperty('RECIPIENT_EMAIL');

  // Fallback for initial setup
  if (!email) {
    Logger.log('⚠️ RECIPIENT_EMAIL not found in Script Properties. Run setupScriptProperties() to configure.');
    throw new Error('RECIPIENT_EMAIL not configured in Script Properties');
  }

  return email;
}

function getIcloudEmail() {
  const props = PropertiesService.getScriptProperties();
  let email = props.getProperty('ICLOUD_EMAIL');

  // Fallback for initial setup
  if (!email) {
    Logger.log('⚠️ ICLOUD_EMAIL not found in Script Properties. Run setupScriptProperties() to configure.');
    throw new Error('ICLOUD_EMAIL not configured in Script Properties');
  }

  return email;
}

/**
 * One-time setup function to store secrets in Script Properties
 * Run this once from the Apps Script editor, then delete or comment out
 *
 * IMPORTANT: Update these values from your .env file before running!
 */
function setupScriptProperties() {
  const props = PropertiesService.getScriptProperties();

  // Update these values from your .env file!
  props.setProperty('DELETE_API_KEY', 'YOUR_DELETE_API_KEY_HERE');
  props.setProperty('RECIPIENT_EMAIL', 'your-email@example.com');
  props.setProperty('FORM_ID', 'YOUR_FORM_ID_HERE');

  Logger.log('✅ Configuration stored in Script Properties');
  Logger.log('   - DELETE_API_KEY: Set');
  Logger.log('   - RECIPIENT_EMAIL: ' + props.getProperty('RECIPIENT_EMAIL'));
  Logger.log('   - FORM_ID: ' + props.getProperty('FORM_ID'));
}

/**
 * Manually process existing rows that haven't been emailed
 * Run this once to process old entries retrospectively
 */
function processExistingWorkExpenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Work");

  if (!sheet) {
    Logger.log("❌ Work sheet not found");
    return;
  }

  const data = sheet.getDataRange().getValues();
  let processedCount = 0;

  // Start from row 2 (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = i + 1; // Sheet rows are 1-indexed
    const rowValues = data[i];
    const emailSent = rowValues[8]; // Column I (0-indexed = 8)

    // Only process if email hasn't been sent
    if (!emailSent || emailSent === "") {
      Logger.log(`Processing row ${row}...`);

      // Config for Work expenses
      const config = {
        statusCol: null,
        dateCol: 2,       // Column B
        descriptionCol: 6, // Column F
        fileCol: 7,       // Column G
        emailSentCol: 9,  // Column I
        sendEmail: true
      };

      // Rename file
      renameFile(sheet, row, "Work", rowValues, config);

      // Send email
      sendWorkExpenseEmail(sheet, row, "Work", rowValues, config);

      processedCount++;
    }
  }

  Logger.log(`✅ Processed ${processedCount} existing Work expense(s)`);
}

/**
 * Install unified trigger for all form submissions
 * Run this once to set up automatic status setting for all forms
 */
function installFormTrigger() {
  // Remove ALL existing form submit triggers to avoid duplicates and conflicts
  const triggers = ScriptApp.getProjectTriggers();
  let removedCount = 0;
  triggers.forEach(trigger => {
    if (trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
      const funcName = trigger.getHandlerFunction();
      Logger.log(`Removing trigger: ${funcName}`);
      ScriptApp.deleteTrigger(trigger);
      removedCount++;
    }
  });

  Logger.log(`✅ Removed ${removedCount} old form submit trigger(s)`);

  // Create new unified trigger
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('handleFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log('✅ Unified form submit trigger installed for handleFormSubmit()');
}

/**
 * Unified handler for all form submissions
 * - Renames files for all forms
 * - Sets status to "to do" for IVA, Health, Income
 * - Sends email for Work expenses only
 */
function handleFormSubmit(e) {
  const row = e.range.getRow();
  if (row === 1) return; // Skip header

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  Logger.log(`Form submitted to sheet: ${sheetName}, row: ${row}`);

  // Configuration for each sheet type
  let config = null;

  if (sheetName === "Work" || sheetName === "Work (Responses)") {
    config = {
      statusCol: 10,    // Column J (Status)
      dateCol: 2,       // Column B
      descriptionCol: 6, // Column F
      fileCol: 7,       // Column G
      emailSentCol: 9,  // Column I
      sendEmail: true
    };
  } else if (sheetName === "IVA" || sheetName === "IVA (Responses)") {
    config = {
      statusCol: 10,    // Column J
      dateCol: 3,       // Column C (Data)
      descriptionCol: 2, // Column B (Número)
      fileCol: 9,       // Column I (Ficheiro)
      emailSentCol: null,
      sendEmail: false
    };
  } else if (sheetName === "Health" || sheetName === "Health (Responses)") {
    config = {
      statusCol: 12,    // Column L
      dateCol: 6,       // Column F (Invoice Date) - used for file naming
      patientCol: 2,    // Column B (Patient)
      providerCol: 4,   // Column D (Provider)
      amountCol: 9,     // Column I (Amount)
      fileCol: 10,      // Column J (Receipt)
      detailsFileCol: 11, // Column K (Details - was Invoice)
      originalReceiptNameCol: 13, // Column M (Original Receipt Filename)
      originalDetailsNameCol: 14, // Column N (Original Details Filename)
      emailSentCol: null,
      sendEmail: false
    };
  } else if (sheetName === "Income" || sheetName === "Income (Responses)") {
    config = {
      statusCol: 8,     // Column H
      dateCol: null,    // No date needed (no file to rename)
      descriptionCol: 9, // Column I (calculated: G + C)
      calculateDescription: true,
      calcCol1: 7,      // Column G (prefix)
      calcCol2: 3,      // Column C (suffix)
      fileCol: null,    // No file upload for Income
      emailSentCol: null,
      sendEmail: false
    };
  } else {
    Logger.log(`No configuration for sheet: ${sheetName}`);
    return;
  }

  // Get row data
  let rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Calculate and set description if needed (Income and Health)
  if (config.calculateDescription) {
    let calculatedDesc;

    if (config.calcMethod === 'healthNames') {
      // Health: Patient first name + Provider first word
      const patient = (rowValues[config.patientCol - 1] || "").toString().trim();
      const provider = (rowValues[config.providerCol - 1] || "").toString().trim();
      const patientFirst = patient.split(/\s+/)[0] || '';
      const providerFirst = provider.split(/\s+/)[0] || '';
      calculatedDesc = `${patientFirst} ${providerFirst}`.trim();
      Logger.log(`${sheetName} Row ${row}: Calculated description "${calculatedDesc}" from patient and provider names`);
    } else if (config.calcMethod === 'firstLetters') {
      // Legacy: First letter of specified columns
      const letters = config.calcCols.map(colNum => {
        const value = (rowValues[colNum - 1] || "").toString().trim();
        return value.charAt(0).toUpperCase();
      }).join('');
      calculatedDesc = letters;
      Logger.log(`${sheetName} Row ${row}: Calculated description "${calculatedDesc}" from first letters`);
    } else {
      // Income: Column G + "-" + Column C
      const col1Value = (rowValues[config.calcCol1 - 1] || "").toString().trim();
      const col2Value = (rowValues[config.calcCol2 - 1] || "").toString().trim();
      calculatedDesc = `${col1Value}-${col2Value}`;
      Logger.log(`${sheetName} Row ${row}: Calculated description "${calculatedDesc}" and wrote to column I`);
    }

    // Write calculated description to the appropriate column
    sheet.getRange(row, config.descriptionCol).setValue(calculatedDesc);

    // Refresh row values to include the calculated description
    rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  // Set status to "To do" if applicable
  if (config.statusCol) {
    const statusCell = sheet.getRange(row, config.statusCol);
    const currentStatus = statusCell.getValue();
    if (!currentStatus || currentStatus === "") {
      statusCell.setValue("To do");
      Logger.log(`${sheetName} Row ${row}: Set status to "To do"`);
    }
  }

  // Rename file (if applicable)
  if (config.fileCol) {
    renameFile(sheet, row, sheetName, rowValues, config);
  }

  // Rename Details file for Health (column K)
  if (config.detailsFileCol) {
    renameDetailsFile(sheet, row, sheetName, rowValues, config);
  }

  // Send email if applicable (Work expenses only)
  if (config.sendEmail && config.emailSentCol) {
    const emailSent = rowValues[config.emailSentCol - 1];
    if (!emailSent) {
      sendWorkExpenseEmail(sheet, row, sheetName, rowValues, config);
    }
  }
}

/**
 * Rename uploaded file with appropriate format based on sheet type
 * - IVA: "Número Data.ext" (e.g., "INV-123 15-01-2025.pdf")
 * - Others: "yyyymmdd_description.ext"
 */
function renameFile(sheet, row, sheetName, rowValues, config) {
  try {
    const fileRef = (rowValues[config.fileCol - 1] || "").toString().trim();
    if (!fileRef) {
      Logger.log(`${sheetName} Row ${row}: No file to rename`);
      return;
    }

    const date = new Date(rowValues[config.dateCol - 1]);
    const description = (rowValues[config.descriptionCol - 1] || "").toString().trim();

    // Extract file ID from URL
    let fileId;
    const idMatch = fileRef.match(/[-\w]{25,}/);
    if (idMatch) fileId = idMatch[0];

    if (!fileId) {
      Logger.log(`${sheetName} Row ${row}: Could not extract file ID`);
      return;
    }

    const file = DriveApp.getFileById(fileId);
    const originalName = file.getName();
    const extMatch = originalName.match(/(\.[^.\s]+)$/);
    const extension = extMatch ? extMatch[0] : "";

    // Store original filename for Health claims (for later Shortcut use)
    // Strip " - Username" suffix added by Google Forms uploads
    if ((sheetName === "Health" || sheetName === "Health (Responses)") && config.originalReceiptNameCol) {
      const cleanedName = originalName.replace(/ - [^.]+(\.[^.]+)$/, '$1');
      sheet.getRange(row, config.originalReceiptNameCol).setValue(cleanedName);
      Logger.log(`${sheetName} Row ${row}: Stored original receipt filename "${cleanedName}"`);
    }

    let newFileName;
    if (sheetName === "IVA" || sheetName === "IVA (Responses)") {
      // IVA format: "Número Data.ext" (e.g., "INV-123 2025-01-15.pdf")
      const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
      newFileName = `${description} ${formattedDate}${extension}`;
    } else if (sheetName === "Health" || sheetName === "Health (Responses)") {
      // Health format: "yymmdd_initial_provider_amount_receipt.ext"
      const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyMMdd");
      const patient = (rowValues[config.patientCol - 1] || "").toString().trim();
      const provider = (rowValues[config.providerCol - 1] || "").toString().trim();
      const amount = (rowValues[config.amountCol - 1] || "").toString().trim();
      const patientInitial = (patient.split(/\s+/)[0] || '').charAt(0) || '';
      const providerFirst = provider.split(/\s+/)[0] || '';
      newFileName = `${formattedDate}_${patientInitial}_${providerFirst}_${amount}_receipt${extension}`;
    } else {
      // Default format: "yyyymmdd_description.ext"
      const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyyMMdd");
      const safeDescription = description.replace(/\s+/g, "_");
      newFileName = `${formattedDate}_${safeDescription}${extension}`;
    }

    file.setName(newFileName);
    Logger.log(`${sheetName} Row ${row}: ✅ Renamed file to "${newFileName}"`);

  } catch (error) {
    Logger.log(`${sheetName} Row ${row}: ❌ File rename error - ${error.toString()}`);
  }
}

/**
 * Rename Details file for Health claims
 * Format: "yyyymmdd_patient_provider_amount_details.ext"
 */
function renameDetailsFile(sheet, row, sheetName, rowValues, config) {
  try {
    const fileRef = (rowValues[config.detailsFileCol - 1] || "").toString().trim();
    if (!fileRef) {
      Logger.log(`${sheetName} Row ${row}: No details file to rename`);
      return;
    }

    // Extract file ID from URL
    let fileId;
    const idMatch = fileRef.match(/[-\w]{25,}/);
    if (idMatch) fileId = idMatch[0];

    if (!fileId) {
      Logger.log(`${sheetName} Row ${row}: Could not extract details file ID`);
      return;
    }

    const file = DriveApp.getFileById(fileId);
    const originalName = file.getName();
    const extMatch = originalName.match(/(\.[^.\s]+)$/);
    const extension = extMatch ? extMatch[0] : "";

    // Store original filename for Health claims (for later Shortcut use)
    // Strip " - Username" suffix added by Google Forms uploads
    if (config.originalDetailsNameCol) {
      const cleanedName = originalName.replace(/ - [^.]+(\.[^.]+)$/, '$1');
      sheet.getRange(row, config.originalDetailsNameCol).setValue(cleanedName);
      Logger.log(`${sheetName} Row ${row}: Stored original details filename "${cleanedName}"`);
    }

    const date = new Date(rowValues[config.dateCol - 1]);
    const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyMMdd");
    const patient = (rowValues[config.patientCol - 1] || "").toString().trim();
    const provider = (rowValues[config.providerCol - 1] || "").toString().trim();
    const amount = (rowValues[config.amountCol - 1] || "").toString().trim();
    const patientInitial = (patient.split(/\s+/)[0] || '').charAt(0) || '';
    const providerFirst = provider.split(/\s+/)[0] || '';

    const newFileName = `${formattedDate}_${patientInitial}_${providerFirst}_${amount}_details${extension}`;

    file.setName(newFileName);
    Logger.log(`${sheetName} Row ${row}: ✅ Renamed details file to "${newFileName}"`);

  } catch (error) {
    Logger.log(`${sheetName} Row ${row}: ❌ Details file rename error - ${error.toString()}`);
  }
}

/**
 * Send email for Work expense (with file attachment)
 */
function sendWorkExpenseEmail(sheet, row, sheetName, rowValues, config) {
  try {
    const recipient = getRecipientEmail();
    const trip = rowValues[0]; // Expense Reason
    const amount = rowValues[3]; // Column D
    const currency = rowValues[4]; // Column E
    const description = (rowValues[config.descriptionCol - 1] || "").toString().trim();
    const fileRef = (rowValues[config.fileCol - 1] || "").toString().trim();

    // Extract file ID
    let fileId;
    const idMatch = fileRef.match(/[-\w]{25,}/);
    if (idMatch) fileId = idMatch[0];

    if (!fileId) {
      Logger.log(`${sheetName} Row ${row}: No file to attach to email`);
      return;
    }

    const file = DriveApp.getFileById(fileId);
    const fileLink = `https://drive.google.com/file/d/${fileId}/view`;
    const subject = `expense ${trip || ""} ${description || ""}`.trim();

    const body = [
      `Hi,`,
      ``,
      `Here is the expense receipt for ${trip} (${description}).`,
      ``,
      `Amount: ${amount} ${currency}`,
      ``,
      `You can also access the file here:`,
      `${fileLink}`,
      ``,
      `Regards,`,
      `Automated System`
    ].join("\n");

    GmailApp.sendEmail(recipient, subject, body, { attachments: [file.getBlob()] });
    sheet.getRange(row, config.emailSentCol).setValue("Yes");
    Logger.log(`${sheetName} Row ${row}: ✅ Email sent to ${recipient}`);

  } catch (error) {
    Logger.log(`${sheetName} Row ${row}: ❌ Email error - ${error.toString()}`);
  }
}

function handleTravel(e) {
  const sheetName = "Work";
  const emailSentCol = 9;               // "Email sent?" column (I)
  const fileLinkCol = 7;                // File link or ID column (G)
  const dateCol = 2;                    // Expense Date column (B)
  const descriptionCol = 6;             // Description column (F)
  const amountCol = 4;                  // Amount column (D)
  const currencyCol = 5;                // Currency column (E)
  const recipient = getRecipientEmail(); // Get from Script Properties

  const row = e.range.getRow();
  const sheet = e.source.getSheetByName(sheetName);
  if (row === 1) return; // Skip header

  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const emailSent = rowValues[emailSentCol - 1];
  if (emailSent) {
    logRun(sheetName, row, "Skipped (already sent)");
    return;
  }

  const trip = rowValues[0];
  const expenseDate = new Date(rowValues[dateCol - 1]);
  const description = (rowValues[descriptionCol - 1] || "").toString().trim().replace(/\s+/g, "_");
  const amount = rowValues[amountCol - 1];
  const currency = rowValues[currencyCol - 1];
  const fileRef = (rowValues[fileLinkCol - 1] || "").toString().trim();

  const formattedDate = Utilities.formatDate(expenseDate, Session.getScriptTimeZone(), "yyyyMMdd");
  const newBaseName = `${formattedDate}_${description}_${amount}_${currency}`;

  let fileId;
  const idMatch = fileRef.match(/[-\w]{25,}/);
  if (idMatch) fileId = idMatch[0];

  try {
    const file = DriveApp.getFileById(fileId);
    const originalName = file.getName();
    const extMatch = originalName.match(/(\.[^.\s]+)$/); // keep extension
    const extension = extMatch ? extMatch[0] : "";
    const newFileName = `${newBaseName}${extension}`;
    file.setName(newFileName);

    // Build Drive file link
    const fileLink = `https://drive.google.com/file/d/${fileId}/view`;

    // Subject: travel claim <trip> <description>
    const subject = `travel claim ${trip || ""} ${description || ""}`.trim();

    // Body with file link
    const body = [
      `Hi,`,
      ``,
      `Here is the travel claim receipt for ${trip} (${description}).`,
      ``,
      `Amount: ${amount} ${currency}`,
      ``,
      `You can also access the file here:`,
      `${fileLink}`,
      ``,
      `Regards,`,
      `Automated System`
    ].join("\n");

    GmailApp.sendEmail(recipient, subject, body, { attachments: [file.getBlob()] });

    sheet.getRange(row, emailSentCol).setValue("Yes");
    logRun(sheetName, row, "Rename + Email", newFileName, recipient, "✅ Success", `Link: ${fileLink}`);
  } catch (error) {
    logRun(sheetName, row, "Rename + Email", newBaseName, recipient, "❌ Error", error.toString());
  }
}

/**
 * Validate a caller-supplied sheet row number.
 * Returns the row as a number, or null if it is not a real data row.
 * Guards against writes to the header row or past the end of the sheet.
 */
function resolveSheetRow(sheet, sheetRow) {
  const row = Number(sheetRow);
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) return null;
  return row;
}

/**
 * Read a full sheet row as a 0-indexed array of values.
 */
function readRowValues(sheet, row) {
  return sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * Web App endpoint for handling GET requests (for testing)
 */
function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({
      status: "Web app is running",
      message: "Use POST requests to add or delete expense reasons",
      version: "7"
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Web App endpoint for handling POST requests (delete expense reason, add expense reason)
 * Supports CORS for localhost and GitHub Pages
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const { action, tripName, expenseReason, apiKey } = data;

    // Support both old 'tripName' and new 'expenseReason' parameter names for backwards compatibility
    const reasonName = expenseReason || tripName;

    // Verify API key
    const expectedKey = getDeleteApiKey();
    if (apiKey !== expectedKey) {
      return createCORSResponse({
        success: false,
        error: "Invalid API key"
      });
    }

    // Status toggles take only the sheet row. Current status, file URLs and
    // invoice numbers are read from the sheet server-side, so a caller cannot
    // point these at an arbitrary Drive file or an unrelated row.
    if (action === "toggleIvaStatus") {
      return createCORSResponse(toggleIvaClaimStatus(data.sheetRow));
    }

    if (action === "toggleWorkStatus") {
      return createCORSResponse(toggleWorkClaimStatus(data.sheetRow));
    }

    if (action === "toggleHealthStatus") {
      return createCORSResponse(toggleHealthClaimStatus(data.sheetRow));
    }

    if (action === "toggleIncomeStatus") {
      return createCORSResponse(toggleIncomeStatus(data.sheetRow));
    }

    if (action === "addTrip" || action === "addExpenseReason") {
      if (!reasonName || reasonName.trim() === "") {
        return createCORSResponse({
          success: false,
          error: "Expense reason is required"
        });
      }
      return createCORSResponse(addExpenseReasonToForm(reasonName.trim()));
    }

    // Archiving replaced deleting. The legacy delete actions are mapped here
    // deliberately, so an older cached page archives rather than destroys.
    if (action === "archiveExpenseReason" || action === "archiveTrip" ||
        action === "deleteTrip" || action === "deleteExpenseReason") {
      if (!reasonName || reasonName.trim() === "") {
        return createCORSResponse({
          success: false,
          error: "Expense reason is required"
        });
      }
      return createCORSResponse(archiveExpenseReasonRows(reasonName.trim()));
    }

    // No implicit default: a missing or unrecognised action must never fall
    // through to deleting rows.
    return createCORSResponse({
      success: false,
      error: `Unknown action: ${action || "(none supplied)"}`
    });

  } catch (error) {
    return createCORSResponse({
      success: false,
      error: error.toString()
    });
  }
}

/**
 * Create response with CORS headers
 */
function createCORSResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);

  // Note: Apps Script Web Apps handle CORS automatically when deployed as "Anyone"
  // This function exists for consistency and future extensibility
  return output;
}

/** Name of the sheet archived Work rows are moved to (created on demand) */
const WORK_ARCHIVE_SHEET_NAME = "Work Archive";

/** Name of the Drive folder archived receipts are moved into (created on demand) */
const ARCHIVE_FOLDER_NAME = "Archived";

/**
 * Count the distinct non-empty expense reasons still present in a sheet
 */
function countRemainingReasons(sheet) {
  const data = sheet.getDataRange().getValues();
  const reasons = new Set();
  for (let i = 1; i < data.length; i++) {
    const reason = data[i][0];
    if (reason != null && String(reason).trim() !== "") {
      reasons.add(String(reason));
    }
  }
  return reasons.size;
}

/**
 * Get (or create) the archive sheet, seeding it with the Work headers
 * plus an "Archived At" column.
 */
function getOrCreateWorkArchiveSheet(ss, headerRow) {
  let archiveSheet = ss.getSheetByName(WORK_ARCHIVE_SHEET_NAME);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(WORK_ARCHIVE_SHEET_NAME);
    const headers = headerRow.concat(["Archived At"]);
    archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    archiveSheet.setFrozenRows(1);
    Logger.log(`Created "${WORK_ARCHIVE_SHEET_NAME}" sheet`);
  }
  return archiveSheet;
}

/**
 * Get (or create) the "Archived" folder alongside a file's current location,
 * so archived receipts stay near the originals rather than at Drive root.
 */
function getOrCreateArchiveFolder(file) {
  const parents = file.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  // Already archived - don't nest Archived inside Archived
  if (parent.getName() === ARCHIVE_FOLDER_NAME) return parent;

  const existing = parent.getFoldersByName(ARCHIVE_FOLDER_NAME);
  return existing.hasNext() ? existing.next() : parent.createFolder(ARCHIVE_FOLDER_NAME);
}

/**
 * Archive all Work rows matching the given expense reason:
 *  - copies the rows to the "Work Archive" sheet (with a timestamp)
 *  - moves each receipt in Drive into an "Archived" folder
 *  - removes the rows from Work and the reason from the form dropdown
 *
 * Rows are only removed from Work after they are safely written to the
 * archive sheet, so a failure part-way through never loses data.
 */
function archiveExpenseReasonRows(expenseReason) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Work");

    if (!sheet) {
      return { success: false, error: "Work sheet not found" };
    }

    const data = sheet.getDataRange().getValues();
    const reasonStr = expenseReason == null ? "" : String(expenseReason);

    // Collect matches top-down so archive order matches sheet order.
    // Compare as strings so numeric reasons like 202511 still match.
    const matches = [];
    for (let i = 1; i < data.length; i++) {
      const rowReason = data[i][0] == null ? "" : String(data[i][0]);
      if (rowReason === reasonStr) matches.push({ index: i, values: data[i] });
    }

    if (matches.length === 0) {
      // Nothing to archive, but still tidy the form dropdown
      removeExpenseReasonFromForm(expenseReason);
      return {
        success: true,
        expenseReason: expenseReason,
        archivedRows: 0,
        movedFiles: 0,
        fileErrors: [],
        remainingReasons: countRemainingReasons(sheet)
      };
    }

    // 1. Write to the archive sheet FIRST - nothing is removed until this works
    const width = data[0].length;
    const archiveSheet = getOrCreateWorkArchiveSheet(ss, data[0]);
    const archivedAt = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"
    );
    const rowsOut = matches.map(m => {
      const values = m.values.slice(0, width);
      while (values.length < width) values.push("");
      return values.concat([archivedAt]);
    });
    archiveSheet
      .getRange(archiveSheet.getLastRow() + 1, 1, rowsOut.length, width + 1)
      .setValues(rowsOut);
    Logger.log(`Archived ${rowsOut.length} row(s) for "${expenseReason}"`);

    // 2. Move receipts into the Archived folder. A file that cannot be moved
    //    is reported but does not abort the archive.
    let movedFiles = 0;
    const fileErrors = [];
    matches.forEach(m => {
      const fileRef = (m.values[6] || "").toString().trim(); // Column G
      if (!fileRef) return;

      const fileId = extractFileId(fileRef);
      if (!fileId) return; // iCloud path or non-Drive reference - nothing to move

      try {
        const file = DriveApp.getFileById(fileId);
        file.moveTo(getOrCreateArchiveFolder(file));
        movedFiles++;
      } catch (fileError) {
        Logger.log(`⚠️ Could not archive file ${fileId} - ${fileError.toString()}`);
        fileErrors.push(fileId);
      }
    });

    // 3. Remove the archived rows from Work, bottom-up to keep indexes valid
    for (let i = matches.length - 1; i >= 0; i--) {
      sheet.deleteRow(matches[i].index + 1); // +1 because sheet rows are 1-indexed
    }

    // 4. Remove the reason from the form dropdown
    const removeResult = removeExpenseReasonFromForm(expenseReason);
    Logger.log(`Remove from form result: ${JSON.stringify(removeResult)}`);

    return {
      success: true,
      expenseReason: expenseReason,
      archivedRows: matches.length,
      movedFiles: movedFiles,
      fileErrors: fileErrors,
      remainingReasons: countRemainingReasons(sheet)
    };
  } catch (error) {
    Logger.log(`Error in archiveExpenseReasonRows: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Remove an expense reason from the Google Form dropdown
 */
function removeExpenseReasonFromForm(expenseReason) {
  try {
    const formId = getFormId();
    const form = FormApp.openById(formId);

    // Find the Expense Reason dropdown question (searches for "expense" or "reason" in title)
    const items = form.getItems();
    let expenseReasonQuestion = null;

    for (let item of items) {
      if (item.getType() === FormApp.ItemType.LIST) {
        const listItem = item.asListItem();
        const title = listItem.getTitle().toLowerCase();
        if (title.includes('expense') || title.includes('reason')) {
          expenseReasonQuestion = listItem;
          Logger.log(`Found form question with title: "${listItem.getTitle()}"`);
          break;
        }
      }
    }

    if (!expenseReasonQuestion) {
      Logger.log("Warning: Expense reason dropdown not found in form");
      // Log all LIST items for debugging
      Logger.log("Available LIST items in form:");
      items.forEach(item => {
        if (item.getType() === FormApp.ItemType.LIST) {
          Logger.log(`  - "${item.getTitle()}" (type: LIST)`);
        }
      });
      return { success: false, error: "Expense reason dropdown not found in form" };
    }

    // Get existing choices and remove the specified expense reason
    const existingChoices = expenseReasonQuestion.getChoices().map(c => c.getValue());
    // Convert both to strings for comparison (handles numbers like 202511)
    const reasonStr = expenseReason ? expenseReason.toString() : "";
    const updatedChoices = existingChoices.filter(choice => {
      const choiceStr = choice ? choice.toString() : "";
      return choiceStr !== reasonStr;
    });

    // Update form with filtered list
    expenseReasonQuestion.setChoices(updatedChoices.map(c => expenseReasonQuestion.createChoice(c)));

    Logger.log(`Removed "${expenseReason}" from form dropdown. ${updatedChoices.length} expense reasons remain.`);

    return {
      success: true,
      expenseReason: expenseReason,
      totalReasons: updatedChoices.length
    };

  } catch (error) {
    Logger.log("Error removing expense reason from form: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Add a new expense reason to the Google Form dropdown
 */
function addExpenseReasonToForm(expenseReason) {
  try {
    const formId = getFormId();
    const form = FormApp.openById(formId);

    // Find the Expense Reason dropdown question (searches for "expense" or "reason" in title)
    const items = form.getItems();
    let expenseReasonQuestion = null;

    for (let item of items) {
      if (item.getType() === FormApp.ItemType.LIST) {
        const listItem = item.asListItem();
        const title = listItem.getTitle().toLowerCase();
        if (title.includes('expense') || title.includes('reason')) {
          expenseReasonQuestion = listItem;
          Logger.log(`Found form question with title: "${listItem.getTitle()}"`);
          break;
        }
      }
    }

    if (!expenseReasonQuestion) {
      Logger.log("Warning: Expense reason dropdown not found in form");
      // Log all LIST items for debugging
      Logger.log("Available LIST items in form:");
      items.forEach(item => {
        if (item.getType() === FormApp.ItemType.LIST) {
          Logger.log(`  - "${item.getTitle()}" (type: LIST)`);
        }
      });
      return { success: false, error: "Expense reason dropdown not found in form" };
    }

    // Get existing choices
    const existingChoices = expenseReasonQuestion.getChoices().map(c => c.getValue());

    // Check if expense reason already exists
    if (existingChoices.includes(expenseReason)) {
      return { success: false, error: `Expense reason "${expenseReason}" already exists in form` };
    }

    // Add new expense reason and sort
    const newChoices = [...existingChoices, expenseReason].sort();
    expenseReasonQuestion.setChoices(newChoices.map(c => expenseReasonQuestion.createChoice(c)));

    Logger.log(`Added "${expenseReason}" to form dropdown. Total: ${newChoices.length} expense reasons.`);

    return {
      success: true,
      expenseReason: expenseReason,
      totalReasons: newChoices.length
    };

  } catch (error) {
    Logger.log("Error adding expense reason to form: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Update the Google Form dropdown with current expense reasons from sheet
 */
function updateFormDropdown(expenseReasons) {
  try {
    const formId = getFormId();
    const form = FormApp.openById(formId);

    // Find the Expense Reason dropdown question (searches for "expense" or "reason" in title)
    const items = form.getItems();
    let expenseReasonQuestion = null;

    for (let item of items) {
      if (item.getType() === FormApp.ItemType.LIST) {
        const listItem = item.asListItem();
        const title = listItem.getTitle().toLowerCase();
        if (title.includes('expense') || title.includes('reason')) {
          expenseReasonQuestion = listItem;
          break;
        }
      }
    }

    if (!expenseReasonQuestion) {
      Logger.log("Warning: Expense reason dropdown not found in form");
      return;
    }

    // Update choices with sorted expense reason list
    const sortedReasons = expenseReasons.sort();
    expenseReasonQuestion.setChoices(sortedReasons.map(r => expenseReasonQuestion.createChoice(r)));

    Logger.log(`Form dropdown updated with ${sortedReasons.length} expense reasons`);

  } catch (error) {
    Logger.log("Error updating form dropdown: " + error.toString());
  }
}

/**
 * Get the Form ID from Script Properties
 */
function getFormId() {
  const props = PropertiesService.getScriptProperties();
  let formId = props.getProperty('FORM_ID');

  // Fallback for initial setup
  if (!formId) {
    Logger.log('⚠️ FORM_ID not found in Script Properties. Run setupScriptProperties() to configure.');
    throw new Error('FORM_ID not configured in Script Properties');
  }

  return formId;
}

/**
 * Toggle IVA claim status between "to do" and "done - DD-MM-YYYY"
 * - When marking done: Rename file to "IVA Claim (DD-MM-YYYY) Número.ext" and send email
 * - When undoing: Revert file name to "Número DD-MM-YYYY.ext"
 */
function toggleIvaClaimStatus(sheetRow) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("IVA");

    if (!sheet) {
      return { success: false, error: "IVA sheet not found" };
    }

    const row = resolveSheetRow(sheet, sheetRow);
    if (!row) {
      return { success: false, error: `Invalid IVA row: ${sheetRow}` };
    }

    // Read everything from the sheet rather than trusting the request body
    const rowValues = readRowValues(sheet, row);
    const currentStatus = (rowValues[9] || '').toString().trim();  // J: Status
    const numero = (rowValues[1] || '').toString().trim();         // B: Número
    const invoiceDate = rowValues[2];                              // C: Data
    const fileUrl = (rowValues[8] || '').toString().trim();        // I: Ficheiro

    const isClaimed = currentStatus.toLowerCase().startsWith('claimed');
    const today = new Date();
    const formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MM-yyyy");

    // Determine new status
    let newStatus;
    if (isClaimed) {
      // Undo: set back to "To do"
      newStatus = "To do";
    } else {
      // Mark as claimed with today's date
      newStatus = `Claimed ${formattedToday}`;
    }

    // Update status in sheet (column J = column 10)
    sheet.getRange(row, 10).setValue(newStatus);
    Logger.log(`IVA Row ${row}: Status changed from "${currentStatus}" to "${newStatus}"`);

    // Handle file rename
    if (fileUrl) {
      const fileId = extractFileId(fileUrl);
      if (fileId) {
        try {
          const file = DriveApp.getFileById(fileId);
          const originalName = file.getName();
          const extMatch = originalName.match(/(\.[^.\s]+)$/);
          const extension = extMatch ? extMatch[0] : "";

          let newFileName;
          if (isClaimed) {
            // Undo: Revert to "Número YYYY-MM-DD.ext"
            // Parse the invoice date (could be in various formats)
            const parsedDate = new Date(invoiceDate);
            const formattedInvoiceDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
            newFileName = `${numero} ${formattedInvoiceDate}${extension}`;
          } else {
            // Mark claimed: Rename to "Claimed (DD-MM-YYYY) Número.ext"
            newFileName = `Claimed (${formattedToday}) ${numero}${extension}`;
          }

          file.setName(newFileName);
          Logger.log(`IVA Row ${row}: ✅ Renamed file to "${newFileName}"`);

        } catch (fileError) {
          Logger.log(`IVA Row ${row}: ⚠️ Could not rename file - ${fileError.toString()}`);
          // Don't fail the whole operation if file rename fails
        }
      }
    }

    // Send email only when marking as claimed (not on undo)
    if (!isClaimed && fileUrl) {
      try {
        const fileId = extractFileId(fileUrl);
        if (fileId) {
          const file = DriveApp.getFileById(fileId);
          const fileName = file.getName();
          const recipient = "jacqueline.eaton@nato.int";
          const subject = fileName.replace(/\.[^.]+$/, ""); // Remove extension for subject
          const fileLink = `https://drive.google.com/file/d/${fileId}/view`;

          const body = [
            `Hi,`,
            ``,
            `Here is an IVA claim receipt.`,
            ``,
            `Número: ${numero}`,
            ``,
            `You can also access the file here:`,
            `${fileLink}`,
            ``,
            `Regards,`,
            `Automated System`
          ].join("\n");

          GmailApp.sendEmail(recipient, subject, body, { attachments: [file.getBlob()] });
          Logger.log(`IVA Row ${row}: ✅ Email sent to ${recipient}`);
        }
      } catch (emailError) {
        Logger.log(`IVA Row ${row}: ⚠️ Could not send email - ${emailError.toString()}`);
        // Don't fail the whole operation if email fails
      }
    }

    return {
      success: true,
      sheetRow: row,
      newStatus: newStatus,
      action: isClaimed ? "undo" : "claimed"
    };

  } catch (error) {
    Logger.log(`Error in toggleIvaClaimStatus: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Extract file ID from a Google Drive URL
 */
function extractFileId(fileUrl) {
  if (!fileUrl) return null;
  const idMatch = fileUrl.match(/[-\w]{25,}/);
  return idMatch ? idMatch[0] : null;
}

/**
 * Toggle Work expense status between "To do" and "Claimed DD-MM-YYYY"
 * - Renames file with "Claimed (DD-MM-YYYY)" prefix when marking claimed
 * - Removes prefix when undoing
 * - No email sent on status change (email already sent on submit)
 */
function toggleWorkClaimStatus(sheetRow) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Work");

    if (!sheet) {
      return { success: false, error: "Work sheet not found" };
    }

    const row = resolveSheetRow(sheet, sheetRow);
    if (!row) {
      return { success: false, error: `Invalid Work row: ${sheetRow}` };
    }

    // Read everything from the sheet rather than trusting the request body
    const rowValues = readRowValues(sheet, row);
    const currentStatus = (rowValues[9] || '').toString().trim(); // J: Status
    const fileUrl = (rowValues[6] || '').toString().trim();       // G: File

    const isClaimed = currentStatus.toLowerCase().startsWith('claimed');
    const today = new Date();
    const formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MM-yyyy");

    // Determine new status
    let newStatus;
    if (isClaimed) {
      newStatus = "To do";
    } else {
      newStatus = `Claimed ${formattedToday}`;
    }

    // Update status in sheet (column J = column 10)
    sheet.getRange(row, 10).setValue(newStatus);
    Logger.log(`Work Row ${row}: Status changed from "${currentStatus}" to "${newStatus}"`);

    // Handle file rename
    if (fileUrl) {
      const fileId = extractFileId(fileUrl);
      if (fileId) {
        try {
          const file = DriveApp.getFileById(fileId);
          const currentName = file.getName();

          let newFileName;
          if (isClaimed) {
            // Undo: Remove "Claimed (DD-MM-YYYY) " prefix
            newFileName = currentName.replace(/^Claimed \(\d{2}-\d{2}-\d{4}\) /, '');
          } else {
            // Mark claimed: Add "Claimed (DD-MM-YYYY) " prefix
            newFileName = `Claimed (${formattedToday}) ${currentName}`;
          }

          file.setName(newFileName);
          Logger.log(`Work Row ${row}: ✅ Renamed file to "${newFileName}"`);

        } catch (fileError) {
          Logger.log(`Work Row ${row}: ⚠️ Could not rename file - ${fileError.toString()}`);
        }
      }
    }

    return {
      success: true,
      sheetRow: row,
      newStatus: newStatus,
      action: isClaimed ? "undo" : "claimed"
    };

  } catch (error) {
    Logger.log(`Error in toggleWorkClaimStatus: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Toggle Health claim status between "To do" and "Claimed DD-MM-YYYY"
 * - On Done: Rename files in Google Drive, send email to trigger iCloud rename Shortcut
 * - On Undo: Rename files back (no email)
 */
function toggleHealthClaimStatus(sheetRow) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Health");

    if (!sheet) {
      return { success: false, error: "Health sheet not found" };
    }

    const row = resolveSheetRow(sheet, sheetRow);
    if (!row) {
      return { success: false, error: `Invalid Health row: ${sheetRow}` };
    }

    // Get row data for patient/provider/date/amount/original filenames.
    // Status and file URLs come from the sheet, not the request body.
    const rowValues = readRowValues(sheet, row);
    const currentStatus = (rowValues[11] || "").toString().trim(); // Column L (Status)
    const fileUrl = (rowValues[9] || "").toString().trim(); // Column J (Receipt file)
    const patient = (rowValues[1] || "").toString().trim(); // Column B
    const provider = (rowValues[3] || "").toString().trim(); // Column D
    const invoiceDate = new Date(rowValues[5]); // Column F (Invoice Date)
    const amount = (rowValues[8] || "").toString().trim(); // Column I (Amount)
    const detailsFileUrl = (rowValues[10] || "").toString().trim(); // Column K (Details file)
    const originalReceiptName = (rowValues[12] || "").toString().trim(); // Column M (Original Receipt Filename)
    const originalDetailsName = (rowValues[13] || "").toString().trim(); // Column N (Original Details Filename)

    const patientInitial = (patient.split(/\s+/)[0] || '').charAt(0) || '';
    const providerFirst = provider.split(/\s+/)[0] || '';

    const isClaimed = (currentStatus || '').toLowerCase().startsWith('claimed');
    const today = new Date();
    const formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MM-yyyy");
    const formattedInvoiceDate = Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), "yyMMdd");

    // Determine new status
    let newStatus;
    if (isClaimed) {
      newStatus = "To do";
    } else {
      newStatus = `Claimed ${formattedToday}`;
    }

    // Update status in sheet (column L = column 12)
    sheet.getRange(row, 12).setValue(newStatus);
    Logger.log(`Health Row ${row}: Status changed from "${currentStatus}" to "${newStatus}"`);

    // Build new filenames for email
    const receiptExt = originalReceiptName.match(/(\.[^.\s]+)$/) ? originalReceiptName.match(/(\.[^.\s]+)$/)[0] : '';
    const detailsExt = originalDetailsName.match(/(\.[^.\s]+)$/) ? originalDetailsName.match(/(\.[^.\s]+)$/)[0] : '';
    const newReceiptName = `${formattedInvoiceDate}_${patientInitial}_${providerFirst}_${amount}_receipt${receiptExt}`;
    const newDetailsName = originalDetailsName ? `${formattedInvoiceDate}_${patientInitial}_${providerFirst}_${amount}_details${detailsExt}` : '';

    // Helper function to rename a Health file in Google Drive
    function renameHealthFile(url, fileType) {
      if (!url) return;
      const fileId = extractFileId(url);
      if (!fileId) return;

      try {
        const file = DriveApp.getFileById(fileId);
        const originalName = file.getName();
        const extMatch = originalName.match(/(\.[^.\s]+)$/);
        const extension = extMatch ? extMatch[0] : "";

        // Base name: yymmdd_initial_provider_amount_type
        const baseName = `${formattedInvoiceDate}_${patientInitial}_${providerFirst}_${amount}_${fileType}`;

        let newFileName;
        if (isClaimed) {
          // Undo: Remove "Claimed (DD-MM-YYYY) " prefix
          newFileName = `${baseName}${extension}`;
        } else {
          // Mark claimed: Add "Claimed (DD-MM-YYYY) " prefix
          newFileName = `Claimed (${formattedToday}) ${baseName}${extension}`;
        }

        file.setName(newFileName);
        Logger.log(`Health Row ${row}: ✅ Renamed ${fileType} file to "${newFileName}"`);

      } catch (fileError) {
        Logger.log(`Health Row ${row}: ⚠️ Could not rename ${fileType} file - ${fileError.toString()}`);
      }
    }

    // Rename Receipt file (column J)
    renameHealthFile(fileUrl, 'receipt');

    // Rename Details file (column K)
    renameHealthFile(detailsFileUrl, 'details');

    // Columns M and N record the CURRENT filename of the iCloud copies, which
    // is what the next rename has to search for. They are updated only after
    // the Shortcut email actually goes out, because that email is the thing
    // that renames those copies.
    //
    // On Undo no email is sent, so the iCloud files keep their names and M/N
    // must be left exactly as they are. (Previously Undo overwrote them with
    // "Claimed/<name>", a path that matched nothing, breaking the next Done.)

    // Send email when marking as claimed (not on undo) to trigger iCloud rename Shortcut
    let icloudEmailSent = false;
    if (!isClaimed && originalReceiptName) {
      try {
        const recipient = getIcloudEmail(); // iCloud email for Shortcut automation
        const subject = `Health Claim Rename: ${patientInitial} ${formattedInvoiceDate} ${providerFirst} ${amount}`;

        // Email body with structured data for Shortcut to parse
        // Each field on its own line for easy parsing
        const body = [
          `ORIGINAL_RECEIPT=${originalReceiptName}`,
          `NEW_RECEIPT=${newReceiptName}`,
          `ORIGINAL_DETAILS=${originalDetailsName || ''}`,
          `NEW_DETAILS=${newDetailsName || ''}`,
        ].join("\n");

        GmailApp.sendEmail(recipient, subject, body);
        icloudEmailSent = true;
        Logger.log(`Health Row ${row}: ✅ Email sent to ${recipient} for iCloud rename`);

        // The Shortcut renames the iCloud copies to these names, so they become
        // the current names for any future rename.
        sheet.getRange(row, 13).setValue(newReceiptName);
        Logger.log(`Health Row ${row}: Current receipt filename now "${newReceiptName}"`);
        if (newDetailsName) {
          sheet.getRange(row, 14).setValue(newDetailsName);
          Logger.log(`Health Row ${row}: Current details filename now "${newDetailsName}"`);
        }

      } catch (emailError) {
        Logger.log(`Health Row ${row}: ⚠️ Could not send email - ${emailError.toString()}`);
      }
    }

    return {
      success: true,
      sheetRow: row,
      newStatus: newStatus,
      action: isClaimed ? "undo" : "claimed",
      icloudEmailSent: icloudEmailSent
    };

  } catch (error) {
    Logger.log(`Error in toggleHealthClaimStatus: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Toggle Income status between "To do" and "Fatura DD-MM-YYYY"
 * - No file rename (Income has no file upload)
 * - No email sent
 */
function toggleIncomeStatus(sheetRow) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Income");

    if (!sheet) {
      return { success: false, error: "Income sheet not found" };
    }

    const row = resolveSheetRow(sheet, sheetRow);
    if (!row) {
      return { success: false, error: `Invalid Income row: ${sheetRow}` };
    }

    // Read status from the sheet rather than trusting the request body
    const rowValues = readRowValues(sheet, row);
    const currentStatus = (rowValues[7] || '').toString().trim(); // H: Status

    const isFatura = currentStatus.toLowerCase().startsWith('fatura');
    const today = new Date();
    const formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MM-yyyy");

    // Determine new status
    let newStatus;
    if (isFatura) {
      newStatus = "To do";
    } else {
      newStatus = `Fatura ${formattedToday}`;
    }

    // Update status in sheet (column H = column 8)
    sheet.getRange(row, 8).setValue(newStatus);
    Logger.log(`Income Row ${row}: Status changed from "${currentStatus}" to "${newStatus}"`);

    return {
      success: true,
      sheetRow: row,
      newStatus: newStatus,
      action: isFatura ? "undo" : "fatura"
    };

  } catch (error) {
    Logger.log(`Error in toggleIncomeStatus: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * ONE-OFF FUNCTION: Rename all Details files in Health sheet
 * Run this manually from Apps Script editor, then delete it
 * Format: yymmdd_initial_provider_amount_details.ext
 */
function oneOffRenameAllHealthDetails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Health");

  if (!sheet) {
    Logger.log("Health sheet not found");
    return;
  }

  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 14); // Start from row 2, columns A-N
  const data = dataRange.getValues();

  let renamedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;

  data.forEach((row, index) => {
    const rowNum = index + 2; // Actual sheet row number
    const detailsFileUrl = (row[10] || "").toString().trim(); // Column K (index 10)

    if (!detailsFileUrl) {
      Logger.log(`Row ${rowNum}: No details file, skipping`);
      skippedCount++;
      return;
    }

    try {
      // Extract file ID
      const idMatch = detailsFileUrl.match(/[-\w]{25,}/);
      if (!idMatch) {
        Logger.log(`Row ${rowNum}: Could not extract file ID from URL`);
        skippedCount++;
        return;
      }
      const fileId = idMatch[0];

      // Get file
      const file = DriveApp.getFileById(fileId);
      const originalName = file.getName();
      const extMatch = originalName.match(/(\.[^.\s]+)$/);
      const extension = extMatch ? extMatch[0] : "";

      // Get data for new filename
      const invoiceDate = new Date(row[5]); // Column F (index 5)
      const patient = (row[1] || "").toString().trim(); // Column B (index 1)
      const provider = (row[3] || "").toString().trim(); // Column D (index 3)
      const amount = (row[8] || "").toString().trim(); // Column I (index 8)

      const formattedDate = Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), "yyMMdd");
      const patientInitial = (patient.split(/\s+/)[0] || '').charAt(0) || '';
      const providerFirst = provider.split(/\s+/)[0] || '';

      const newFileName = `${formattedDate}_${patientInitial}_${providerFirst}_${amount}_details${extension}`;

      // Check if already renamed
      if (originalName === newFileName) {
        Logger.log(`Row ${rowNum}: Already renamed, skipping`);
        skippedCount++;
        return;
      }

      // Rename
      file.setName(newFileName);
      Logger.log(`Row ${rowNum}: ✅ Renamed "${originalName}" to "${newFileName}"`);
      renamedCount++;

    } catch (error) {
      Logger.log(`Row ${rowNum}: ❌ Error - ${error.toString()}`);
      errorCount++;
    }
  });

  Logger.log(`\n=== SUMMARY ===`);
  Logger.log(`Renamed: ${renamedCount}`);
  Logger.log(`Skipped: ${skippedCount}`);
  Logger.log(`Errors: ${errorCount}`);
}
