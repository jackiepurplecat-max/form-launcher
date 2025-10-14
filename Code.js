// Configuration stored in Script Properties for security
// To set: Run setupScriptProperties() once, or set manually in Project Settings > Script Properties

function getDeleteApiKey() {
  const props = PropertiesService.getScriptProperties();
  let apiKey = props.getProperty('DELETE_API_KEY');

  // Fallback for initial setup - remove after setting Script Properties
  if (!apiKey) {
    apiKey = 'K9mP2xR7nQ4vL8wT6hY3jF5bN1cZ9sD4';
    Logger.log('⚠️ Using hardcoded API key. Run setupScriptProperties() to store securely.');
  }

  return apiKey;
}

function getRecipientEmail() {
  const props = PropertiesService.getScriptProperties();
  let email = props.getProperty('RECIPIENT_EMAIL');

  // Fallback for initial setup
  if (!email) {
    email = 'you@example.com';
    Logger.log('⚠️ Using placeholder email. Run setupScriptProperties() to set your email.');
  }

  return email;
}

/**
 * One-time setup function to store secrets in Script Properties
 * Run this once from the Apps Script editor, then delete or comment out
 *
 * IMPORTANT: Update these values before running!
 */
function setupScriptProperties() {
  const props = PropertiesService.getScriptProperties();

  // Update these values!
  props.setProperty('DELETE_API_KEY', 'K9mP2xR7nQ4vL8wT6hY3jF5bN1cZ9sD4');
  props.setProperty('RECIPIENT_EMAIL', 'jacqueline.eaton@nato.int');
  props.setProperty('FORM_ID', '1Ug8DFOA2lKhrSJZ58ctjCXQMTwkw7UU-DfkNBVpaEpw');

  Logger.log('✅ Configuration stored in Script Properties');
  Logger.log('   - DELETE_API_KEY: Set');
  Logger.log('   - RECIPIENT_EMAIL: ' + props.getProperty('RECIPIENT_EMAIL'));
  Logger.log('   - FORM_ID: ' + props.getProperty('FORM_ID'));
}

function handleTravel(e) {
  const sheetName = "Travel";
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

    // Subject: travel claim receipt <trip> <description>
    const subject = `travel claim receipt ${trip || ""} ${description || ""}`.trim();

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
 * Web App endpoint for handling POST requests (delete trip, add trip)
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
      // Add trip to form dropdown
      if (!tripName || tripName.trim() === "") {
        return createCORSResponse({
          success: false,
          error: "Trip name is required"
        });
      }
      const result = addTripToForm(tripName.trim());
      return createCORSResponse(result);

    } else {
      // Default action: delete trip (backwards compatibility)
      if (!tripName || tripName.trim() === "") {
        return createCORSResponse({
          success: false,
          error: "Trip name is required"
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
 * Delete all rows from Travel sheet matching the given trip name
 */
function deleteTripRows(tripName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Travel");

  if (!sheet) {
    return { success: false, error: "Travel sheet not found" };
  }

  const data = sheet.getDataRange().getValues();
  let deletedCount = 0;

  // Start from bottom to avoid index shifting issues when deleting
  for (let i = data.length - 1; i > 0; i--) {  // Skip header row (i > 0)
    const rowTripName = data[i][0]; // Column A (Trip name)

    if (rowTripName === tripName) {
      sheet.deleteRow(i + 1); // +1 because sheet rows are 1-indexed
      deletedCount++;
    }
  }

  // Get remaining unique trips for the response
  const remainingData = sheet.getDataRange().getValues();
  const uniqueTrips = new Set();
  for (let i = 1; i < remainingData.length; i++) {
    const trip = remainingData[i][0];
    if (trip && trip.trim() !== "") {
      uniqueTrips.add(trip);
    }
  }

  // Update form dropdown after deletion
  updateFormDropdown(Array.from(uniqueTrips));

  return {
    success: true,
    trip: tripName,
    deletedRows: deletedCount,
    remainingTrips: uniqueTrips.size
  };
}

/**
 * Add a new trip to the Google Form dropdown
 */
function addTripToForm(tripName) {
  try {
    const formId = getFormId();
    const form = FormApp.openById(formId);

    // Find the Trip dropdown question
    const items = form.getItems();
    let tripQuestion = null;

    for (let item of items) {
      if (item.getType() === FormApp.ItemType.LIST) {
        const listItem = item.asListItem();
        if (listItem.getTitle().toLowerCase().includes('trip')) {
          tripQuestion = listItem;
          break;
        }
      }
    }

    if (!tripQuestion) {
      return { success: false, error: "Trip dropdown not found in form" };
    }

    // Get existing choices
    const existingChoices = tripQuestion.getChoices().map(c => c.getValue());

    // Check if trip already exists
    if (existingChoices.includes(tripName)) {
      return { success: false, error: `Trip "${tripName}" already exists in form` };
    }

    // Add new trip and sort
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
 * Update the Google Form dropdown with current trips from sheet
 */
function updateFormDropdown(trips) {
  try {
    const formId = getFormId();
    const form = FormApp.openById(formId);

    // Find the Trip dropdown question
    const items = form.getItems();
    let tripQuestion = null;

    for (let item of items) {
      if (item.getType() === FormApp.ItemType.LIST) {
        const listItem = item.asListItem();
        if (listItem.getTitle().toLowerCase().includes('trip')) {
          tripQuestion = listItem;
          break;
        }
      }
    }

    if (!tripQuestion) {
      Logger.log("Warning: Trip dropdown not found in form");
      return;
    }

    // Update choices with sorted trip list
    const sortedTrips = trips.sort();
    tripQuestion.setChoices(sortedTrips.map(t => tripQuestion.createChoice(t)));

    Logger.log(`Form dropdown updated with ${sortedTrips.length} trips`);

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
    formId = 'YOUR_FORM_ID_HERE';
    Logger.log('⚠️ Using placeholder form ID. Run setupScriptProperties() to set your form ID.');
  }

  return formId;
}


