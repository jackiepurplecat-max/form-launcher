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
 * Install trigger for IVA form submissions
 * Run this once to set up automatic status setting
 */
function installIVATrigger() {
  // Remove existing triggers for handleIVA to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'handleIVA') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('handleIVA')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log('✅ Form submit trigger installed for handleIVA()');
}

/**
 * Handle IVA form submissions - set default status to "to do"
 */
function handleIVA(e) {
  const sheetName = "IVA";
  const statusCol = 10;  // Column J (Status)

  const row = e.range.getRow();
  const sheet = e.source.getSheetByName(sheetName);
  if (row === 1) return; // Skip header

  // Set default status to "to do" if empty
  const statusCell = sheet.getRange(row, statusCol);
  const currentStatus = statusCell.getValue();

  if (!currentStatus || currentStatus === "") {
    statusCell.setValue("to do");
    Logger.log(`Row ${row}: Set default status to "to do"`);
  }
}

function handleTravel(e) {
  const sheetName = "Claims Form (Responses)";
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
 * Web App endpoint for handling POST requests (delete expense reason, add expense reason)
 * Supports CORS for localhost and GitHub Pages
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const { action, tripName, apiKey } = data;

    // Verify API key
    const expectedKey = getDeleteApiKey();
    if (apiKey !== expectedKey) {
      return createCORSResponse({
        success: false,
        error: "Invalid API key"
      });
    }

    // Handle different actions
    if (action === "addTrip") {
      // Add expense reason to form dropdown
      if (!tripName || tripName.trim() === "") {
        return createCORSResponse({
          success: false,
          error: "Expense reason is required"
        });
      }
      const result = addTripToForm(tripName.trim());
      return createCORSResponse(result);

    } else {
      // Default action: delete expense reason (backwards compatibility)
      if (!tripName || tripName.trim() === "") {
        return createCORSResponse({
          success: false,
          error: "Expense reason is required"
        });
      }
      const result = deleteTripRows(tripName.trim());
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
 * Delete all rows from Claims Form (Responses) sheet matching the given expense reason
 */
function deleteTripRows(tripName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Claims Form (Responses)");

    if (!sheet) {
      return { success: false, error: "Claims Form (Responses) sheet not found" };
    }

    const data = sheet.getDataRange().getValues();
    let deletedCount = 0;

    // Start from bottom to avoid index shifting issues when deleting
    for (let i = data.length - 1; i > 0; i--) {  // Skip header row (i > 0)
      const rowTripName = data[i][0]; // Column A (Expense Reason)

      // Convert both to strings for comparison (handles numbers like 202511)
      const rowTripStr = rowTripName ? rowTripName.toString() : "";
      const tripNameStr = tripName ? tripName.toString() : "";

      if (rowTripStr === tripNameStr) {
        sheet.deleteRow(i + 1); // +1 because sheet rows are 1-indexed
        deletedCount++;
      }
    }

    Logger.log(`Deleted ${deletedCount} rows for expense reason: ${tripName}`);

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
    const removeResult = removeTripFromForm(tripName);
    Logger.log(`Remove from form result: ${JSON.stringify(removeResult)}`);

    return {
      success: true,
      trip: tripName,
      deletedRows: deletedCount,
      remainingTrips: uniqueReasons.size
    };
  } catch (error) {
    Logger.log(`Error in deleteTripRows: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Remove an expense reason from the Google Form dropdown
 */
function removeTripFromForm(tripName) {
  try {
    const formId = getFormId();
    const form = FormApp.openById(formId);

    // Find the Expense Reason dropdown question (searches for "trip" or "expense" in title)
    const items = form.getItems();
    let tripQuestion = null;

    for (let item of items) {
      if (item.getType() === FormApp.ItemType.LIST) {
        const listItem = item.asListItem();
        const title = listItem.getTitle().toLowerCase();
        if (title.includes('trip') || title.includes('expense') || title.includes('reason')) {
          tripQuestion = listItem;
          break;
        }
      }
    }

    if (!tripQuestion) {
      Logger.log("Warning: Expense reason dropdown not found in form");
      return { success: false, error: "Expense reason dropdown not found in form" };
    }

    // Get existing choices and remove the specified expense reason
    const existingChoices = tripQuestion.getChoices().map(c => c.getValue());
    // Convert both to strings for comparison (handles numbers like 202511)
    const tripNameStr = tripName ? tripName.toString() : "";
    const updatedChoices = existingChoices.filter(t => {
      const tStr = t ? t.toString() : "";
      return tStr !== tripNameStr;
    });

    // Update form with filtered list
    tripQuestion.setChoices(updatedChoices.map(c => tripQuestion.createChoice(c)));

    Logger.log(`Removed "${tripName}" from form dropdown. ${updatedChoices.length} expense reasons remain.`);

    return {
      success: true,
      tripName: tripName,
      totalTrips: updatedChoices.length
    };

  } catch (error) {
    Logger.log("Error removing expense reason from form: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Add a new expense reason to the Google Form dropdown
 */
function addTripToForm(tripName) {
  try {
    const formId = getFormId();
    const form = FormApp.openById(formId);

    // Find the Expense Reason dropdown question (searches for "trip" or "expense" in title)
    const items = form.getItems();
    let tripQuestion = null;

    for (let item of items) {
      if (item.getType() === FormApp.ItemType.LIST) {
        const listItem = item.asListItem();
        const title = listItem.getTitle().toLowerCase();
        if (title.includes('trip') || title.includes('expense') || title.includes('reason')) {
          tripQuestion = listItem;
          break;
        }
      }
    }

    if (!tripQuestion) {
      return { success: false, error: "Expense reason dropdown not found in form" };
    }

    // Get existing choices
    const existingChoices = tripQuestion.getChoices().map(c => c.getValue());

    // Check if expense reason already exists
    if (existingChoices.includes(tripName)) {
      return { success: false, error: `Expense reason "${tripName}" already exists in form` };
    }

    // Add new expense reason and sort
    const newChoices = [...existingChoices, tripName].sort();
    tripQuestion.setChoices(newChoices.map(c => tripQuestion.createChoice(c)));

    return {
      success: true,
      tripName: tripName,
      totalTrips: newChoices.length
    };

  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Update the Google Form dropdown with current expense reasons from sheet
 */
function updateFormDropdown(trips) {
  try {
    const formId = getFormId();
    const form = FormApp.openById(formId);

    // Find the Expense Reason dropdown question (searches for "trip" or "expense" in title)
    const items = form.getItems();
    let tripQuestion = null;

    for (let item of items) {
      if (item.getType() === FormApp.ItemType.LIST) {
        const listItem = item.asListItem();
        const title = listItem.getTitle().toLowerCase();
        if (title.includes('trip') || title.includes('expense') || title.includes('reason')) {
          tripQuestion = listItem;
          break;
        }
      }
    }

    if (!tripQuestion) {
      Logger.log("Warning: Expense reason dropdown not found in form");
      return;
    }

    // Update choices with sorted expense reason list
    const sortedTrips = trips.sort();
    tripQuestion.setChoices(sortedTrips.map(t => tripQuestion.createChoice(t)));

    Logger.log(`Form dropdown updated with ${sortedTrips.length} expense reasons`);

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


