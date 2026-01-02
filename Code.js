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
      dateCol: 6,       // Column F (Date)
      descriptionCol: 13, // Column M (calculated: first letter of B + D + I)
      calculateDescription: true,
      calcMethod: 'firstLetters', // Special calculation method
      calcCols: [2, 4, 9], // Columns B, D, I
      fileCol: 10,      // Column J (Receipt)
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

    if (config.calcMethod === 'firstLetters') {
      // Health: First letter of columns B, D, I
      const letters = config.calcCols.map(colNum => {
        const value = (rowValues[colNum - 1] || "").toString().trim();
        return value.charAt(0).toUpperCase();
      }).join('');
      calculatedDesc = letters;
      Logger.log(`${sheetName} Row ${row}: Calculated description "${calculatedDesc}" from first letters and wrote to column M`);
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

    let newFileName;
    if (sheetName === "IVA" || sheetName === "IVA (Responses)") {
      // IVA format: "Número Data.ext" (e.g., "INV-123 2025-01-15.pdf")
      const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
      newFileName = `${description} ${formattedDate}${extension}`;
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

    // Handle different actions
    if (action === "toggleIvaStatus") {
      // Toggle IVA claim status (claimed/undo)
      const result = toggleIvaClaimStatus(
        data.sheetRow,
        data.currentStatus,
        data.numero,
        data.data,  // Invoice date (Data field)
        data.fileUrl
      );
      return createCORSResponse(result);

    } else if (action === "toggleWorkStatus") {
      // Toggle Work expense status (claimed/undo) - no email on status change
      const result = toggleWorkClaimStatus(
        data.sheetRow,
        data.currentStatus,
        data.fileUrl
      );
      return createCORSResponse(result);

    } else if (action === "toggleHealthStatus") {
      // Toggle Health claim status (claimed/undo) - no email
      const result = toggleHealthClaimStatus(
        data.sheetRow,
        data.currentStatus,
        data.fileUrl
      );
      return createCORSResponse(result);

    } else if (action === "toggleIncomeStatus") {
      // Toggle Income status (fatura/undo) - no file, no email
      const result = toggleIncomeStatus(
        data.sheetRow,
        data.currentStatus
      );
      return createCORSResponse(result);

    } else if (action === "addTrip" || action === "addExpenseReason") {
      // Add expense reason to form dropdown
      if (!reasonName || reasonName.trim() === "") {
        return createCORSResponse({
          success: false,
          error: "Expense reason is required"
        });
      }
      const result = addExpenseReasonToForm(reasonName.trim());
      return createCORSResponse(result);

    } else {
      // Default action: delete expense reason (backwards compatibility)
      if (!reasonName || reasonName.trim() === "") {
        return createCORSResponse({
          success: false,
          error: "Expense reason is required"
        });
      }
      const result = deleteExpenseReasonRows(reasonName.trim());
      return createCORSResponse(result);
    }

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

/**
 * Delete all rows from Work sheet matching the given expense reason
 */
function deleteExpenseReasonRows(expenseReason) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Work");

    if (!sheet) {
      return { success: false, error: "Work sheet not found" };
    }

    const data = sheet.getDataRange().getValues();
    let deletedCount = 0;

    // Start from bottom to avoid index shifting issues when deleting
    for (let i = data.length - 1; i > 0; i--) {  // Skip header row (i > 0)
      const rowExpenseReason = data[i][0]; // Column A (Expense Reason)

      // Convert both to strings for comparison (handles numbers like 202511)
      const rowReasonStr = rowExpenseReason ? rowExpenseReason.toString() : "";
      const reasonStr = expenseReason ? expenseReason.toString() : "";

      if (rowReasonStr === reasonStr) {
        sheet.deleteRow(i + 1); // +1 because sheet rows are 1-indexed
        deletedCount++;
      }
    }

    Logger.log(`Deleted ${deletedCount} rows for expense reason: ${expenseReason}`);

    // Get remaining unique expense reasons from spreadsheet for the response
    const remainingData = sheet.getDataRange().getValues();
    const uniqueReasons = new Set();
    for (let i = 1; i < remainingData.length; i++) {
      const reason = remainingData[i][0];
      // Convert to string safely
      if (reason != null && reason !== "") {
        const reasonStr = String(reason);  // Use String() instead of toString()
        if (reasonStr.trim() !== "") {
          uniqueReasons.add(reasonStr);
        }
      }
    }

    Logger.log(`Remaining unique expense reasons: ${Array.from(uniqueReasons).join(', ')}`);

    // Remove only the deleted expense reason from form dropdown
    const removeResult = removeExpenseReasonFromForm(expenseReason);
    Logger.log(`Remove from form result: ${JSON.stringify(removeResult)}`);

    return {
      success: true,
      expenseReason: expenseReason,
      deletedRows: deletedCount,
      remainingReasons: uniqueReasons.size
    };
  } catch (error) {
    Logger.log(`Error in deleteExpenseReasonRows: ${error.toString()}`);
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
function toggleIvaClaimStatus(sheetRow, currentStatus, numero, invoiceDate, fileUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("IVA");

    if (!sheet) {
      return { success: false, error: "IVA sheet not found" };
    }

    const isClaimed = (currentStatus || '').toLowerCase().startsWith('claimed');
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
    sheet.getRange(sheetRow, 10).setValue(newStatus);
    Logger.log(`IVA Row ${sheetRow}: Status changed from "${currentStatus}" to "${newStatus}"`);

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
          Logger.log(`IVA Row ${sheetRow}: ✅ Renamed file to "${newFileName}"`);

        } catch (fileError) {
          Logger.log(`IVA Row ${sheetRow}: ⚠️ Could not rename file - ${fileError.toString()}`);
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
          Logger.log(`IVA Row ${sheetRow}: ✅ Email sent to ${recipient}`);
        }
      } catch (emailError) {
        Logger.log(`IVA Row ${sheetRow}: ⚠️ Could not send email - ${emailError.toString()}`);
        // Don't fail the whole operation if email fails
      }
    }

    return {
      success: true,
      sheetRow: sheetRow,
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
function toggleWorkClaimStatus(sheetRow, currentStatus, fileUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Work");

    if (!sheet) {
      return { success: false, error: "Work sheet not found" };
    }

    const isClaimed = (currentStatus || '').toLowerCase().startsWith('claimed');
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
    sheet.getRange(sheetRow, 10).setValue(newStatus);
    Logger.log(`Work Row ${sheetRow}: Status changed from "${currentStatus}" to "${newStatus}"`);

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
          Logger.log(`Work Row ${sheetRow}: ✅ Renamed file to "${newFileName}"`);

        } catch (fileError) {
          Logger.log(`Work Row ${sheetRow}: ⚠️ Could not rename file - ${fileError.toString()}`);
        }
      }
    }

    return {
      success: true,
      sheetRow: sheetRow,
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
 * - Renames file with "Claimed (DD-MM-YYYY)" prefix when marking claimed
 * - Removes prefix when undoing
 * - No email sent
 */
function toggleHealthClaimStatus(sheetRow, currentStatus, fileUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Health");

    if (!sheet) {
      return { success: false, error: "Health sheet not found" };
    }

    const isClaimed = (currentStatus || '').toLowerCase().startsWith('claimed');
    const today = new Date();
    const formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MM-yyyy");

    // Determine new status
    let newStatus;
    if (isClaimed) {
      newStatus = "To do";
    } else {
      newStatus = `Claimed ${formattedToday}`;
    }

    // Update status in sheet (column L = column 12)
    sheet.getRange(sheetRow, 12).setValue(newStatus);
    Logger.log(`Health Row ${sheetRow}: Status changed from "${currentStatus}" to "${newStatus}"`);

    // Handle file rename (Receipt file in column J = column 10)
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
          Logger.log(`Health Row ${sheetRow}: ✅ Renamed file to "${newFileName}"`);

        } catch (fileError) {
          Logger.log(`Health Row ${sheetRow}: ⚠️ Could not rename file - ${fileError.toString()}`);
        }
      }
    }

    return {
      success: true,
      sheetRow: sheetRow,
      newStatus: newStatus,
      action: isClaimed ? "undo" : "claimed"
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
function toggleIncomeStatus(sheetRow, currentStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Income");

    if (!sheet) {
      return { success: false, error: "Income sheet not found" };
    }

    const isFatura = (currentStatus || '').toLowerCase().startsWith('fatura');
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
    sheet.getRange(sheetRow, 8).setValue(newStatus);
    Logger.log(`Income Row ${sheetRow}: Status changed from "${currentStatus}" to "${newStatus}"`);

    return {
      success: true,
      sheetRow: sheetRow,
      newStatus: newStatus,
      action: isFatura ? "undo" : "fatura"
    };

  } catch (error) {
    Logger.log(`Error in toggleIncomeStatus: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}
