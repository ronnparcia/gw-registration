function addTicketNumber() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var properties = PropertiesService.getScriptProperties();
  var lastTicketNumber = properties.getProperty("lastTicketNumber");
  lastTicketNumber = lastTicketNumber ? parseInt(lastTicketNumber) : 0;

  var ticketColumn = 1; 
  var timestampColumn = 2;
  
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var ticketCell = sheet.getRange(i + 1, ticketColumn);
    var timestampCell = row[timestampColumn - 1];
    
    if (!ticketCell.getValue() && timestampCell) {
      
      // itcket no increment
      lastTicketNumber++; 
      
      var timestamp = new Date(timestampCell);
      var month = String(timestamp.getMonth() + 1).padStart(2, "0");
      var day = String(timestamp.getDate()).padStart(2, "0");
      var dateCode = month + day;
      
      var ticketNumber = String(lastTicketNumber).padStart(5, "0") + "-" + dateCode;
      ticketCell.setValue(ticketNumber);

      // Generate card
      try {
        generateCardForRow(i + 1);
      } catch (error) {
        Logger.log("Failed to generate card for row " + (i + 1) + ": " + error.message);
      }
    }
  }
  
  // saves last number used
  properties.setProperty("lastTicketNumber", lastTicketNumber); 
}


/**
 * Generates a card for the specified row.
 * This function retrieves the header mapping from the first row of the sheet,
 * extracts values from the row using the header names, converts text to uppercase,
 * and creates a card (by copying a template) with the placeholders replaced.
 *
 * Expected header names:
 *   "Account Number", "OR Number", "Last Name", "First Name",
 *   "Middle Name", "ID Number", "College", "Degree Code",
 *   "Alt. Deg. Code", "Chosen Package", "Term of Payment"
 *
 * @param {Number} rowNumber - The row number (1-indexed) in the sheet.
 */
function generateCardForRow(rowNumber) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Retrieve header row (first row) and build header mapping.
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerMapping = {};
  for (var col = 0; col < headerRow.length; col++) {
    headerMapping[headerRow[col]] = col;
  }
  
  // Get the row data for the specified row number.
  var row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Helper: Get cell value by header name (converted to uppercase).
  function getCellValue(headerName) {
    var value = row[headerMapping[headerName]];
    return value ? value.toString().toUpperCase() : "UNKNOWN";
  }
  
  // Extract values.
  var accountNumber = getCellValue("Account Number");
  var orNumber = getCellValue("OR Number");
  var lastName = getCellValue("Last Name");
  var firstName = getCellValue("First Name");
  var middleName = getCellValue("Middle Name"); // New field.
  var idNumber = getCellValue("ID Number");
  var college = getCellValue("College");
  
  // Use "Alt. Deg. Code" if "Degree Code" equals "MY DEGREE CODE ISN'T IN THE LIST".
  var degree = getCellValue("Degree Code");
  if (degree === "MY DEGREE CODE ISN'T IN THE LIST" && headerMapping["Alt. Deg. Code"] !== undefined) {
    degree = getCellValue("Alt. Deg. Code");
  }
  
  var chosenPackage = getCellValue("Chosen Package");
  var termOfPayment = getCellValue("Term of Payment");
  
  // Determine package price based on chosen package.
  var packagePrice = "";
  if (chosenPackage === "PACKAGE A (BUSINESS)") {
    packagePrice = "P5,000";
  } else if (chosenPackage === "PACKAGE B (CREATIVE)") {
    packagePrice = "P5,150";
  } else if (chosenPackage === "PACKAGE C (A+B)") {
    packagePrice = "P5,300";
  } else if (chosenPackage === "PACKAGE D (SCHOLARS)") {
    packagePrice = "P4,800";
  } else {
    packagePrice = "UNKNOWN";
  }
  
  // Template and folder settings.
  var templateId = "1kPmaYz7pR0SvOKA_3qL85KJN5rHTcHTl6lau8h_l180";
  var folderId = "1ictsEDej7qYc2sI4_udzKp6SPxi3qoYR";
  
  // Create date string and build copyName: "Card - dateStr - LAST NAME - ACCOUNT NUMBER"
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var copyName = accountNumber + " - " + idNumber + " - " + lastName;
  
  // Retrieve template file and destination folder.
  var templateFile = DriveApp.getFileById(templateId);
  var folder = DriveApp.getFolderById(folderId);
  
  // Make a copy of the template in the specified folder.
  var cardFile = templateFile.makeCopy(copyName, folder);
  
  // Open the new document to modify its content.
  var doc = DocumentApp.openById(cardFile.getId());
  var body = doc.getBody();
  
  // Replace placeholders with the actual (all-caps) values.
  body.replaceText("{{AccountNumber}}", accountNumber);
  body.replaceText("{{ORNumber}}", orNumber);
  body.replaceText("{{LastName}}", lastName);
  body.replaceText("{{FirstName}}", firstName);
  body.replaceText("{{MiddleName}}", middleName);
  body.replaceText("{{IDNumber}}", idNumber);
  body.replaceText("{{College}}", college);
  body.replaceText("{{Degree}}", degree);
  body.replaceText("{{Package}}", chosenPackage);
  body.replaceText("{{TermOfPayment}}", termOfPayment);
  body.replaceText("{{PackagePrice}}", packagePrice);
  
  // Save and close the document.
  doc.saveAndClose(); 
  
  Logger.log("Card created for " + accountNumber + " at row " + rowNumber);
}

/**
 * onOpen() adds a custom menu item "Regenerate Card" to the Google Sheets UI.
 */
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Regenerate Card')
      .addItem('Regenerate Card', 'regenerateCard')
      .addToUi();
  }
  
  /**
   * regenerateCard() prompts the user for the Account Number,
   * searches for the matching row in the sheet,
   * and then calls regenerateCardForRow() to create a new edited card.
   */
  function regenerateCard() {
    var ui = SpreadsheetApp.getUi();
    
    // Prompt the user to enter the Account Number.
    var response = ui.prompt('Regenerate Card', 'Please enter the ACCOUNT NUMBER:', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) {
      ui.alert('Operation cancelled.');
      return;
    }
    
    var inputAccount = response.getResponseText().trim();
    if (!inputAccount) {
      ui.alert('No Account Number entered.');
      return;
    }
    
    // Retrieve all data from the active sheet.
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    
    // Build header mapping from the first row (headers) to column indexes.
    var headers = data[0];
    var headerMapping = {};
    for (var col = 0; col < headers.length; col++) {
      headerMapping[headers[col]] = col;
    }
    
    // Ensure required headers exist.
    if (headerMapping["Account Number"] === undefined || headerMapping["Last Name"] === undefined) {
      ui.alert('Required headers ("Account Number" and "Last Name") not found.');
      return;
    }
    
    // Search for the row where the "Account Number" matches the input.
    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var acctVal = row[headerMapping["Account Number"]].toString().trim();
      if (acctVal.toUpperCase() === inputAccount.toUpperCase()) {
        foundRow = i + 1; // Use 1-indexed row number
        break;
      }
    }
    
    if (foundRow === -1) {
      ui.alert('Account Number "' + inputAccount + '" not found in the sheet.');
      return;
    }
    
    // Call the helper function to generate a new edited card for the found row.
    try {
      regenerateCardForRow(foundRow);
      ui.alert('New edited card generated successfully for Account Number: ' + inputAccount);
    } catch (error) {
      ui.alert('Error regenerating card: ' + error);
      Logger.log('Error regenerating card for Account Number ' + inputAccount + ': ' + error);
    }
  }
  
  /**
   * regenerateCardForRow() creates a new card for the specified row.
   * It uses the row’s data to build the file name and performs placeholder replacements.
   * The new file name format is:
   *   "[Account Number] - [ID Number] - [Last Name] - Edited (yyyy-MM-dd HH:mm:ss)"
   */
  function regenerateCardForRow(rowNumber) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Retrieve the header row and build a header mapping.
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var headerMapping = {};
    for (var col = 0; col < headerRow.length; col++) {
      headerMapping[headerRow[col]] = col;
    }
    
    // Get the data row.
    var row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Helper function: get cell value by header name, converted to uppercase.
    function getCell(headerName) {
      var value = row[headerMapping[headerName]];
      return value ? value.toString().toUpperCase() : "UNKNOWN";
    }
    
    // Extract all required fields.
    var accountNumber = getCell("Account Number");
    var orNumber = getCell("OR Number");
    var lastName = getCell("Last Name");
    var firstName = getCell("First Name");
    var middleName = getCell("Middle Name");
    var idNumber = getCell("ID Number");
    var college = getCell("College");
    
    // Use Alt. Deg. Code if "Degree Code" is "MY DEGREE CODE ISN'T IN THE LIST".
    var degree = getCell("Degree Code");
    if (degree === "MY DEGREE CODE ISN'T IN THE LIST" && headerMapping["Alt. Deg. Code"] !== undefined) {
      degree = getCell("Alt. Deg. Code");
    }
    
    var chosenPackage = getCell("Chosen Package");
    var termOfPayment = getCell("Term of Payment");
    
    // Determine package price based on chosen package.
    var packagePrice = "";
    if (chosenPackage === "PACKAGE A (BUSINESS)") {
      packagePrice = "P5,000";
    } else if (chosenPackage === "PACKAGE B (CREATIVE)") {
      packagePrice = "P5,150";
    } else if (chosenPackage === "PACKAGE C (A+B)") {
      packagePrice = "P5,300";
    } else if (chosenPackage === "PACKAGE D (SCHOLARS)") {
      packagePrice = "P4,800";
    } else {
      packagePrice = "UNKNOWN";
    }
    
    // Template and folder settings.
    var templateId = "1kPmaYz7pR0SvOKA_3qL85KJN5rHTcHTl6lau8h_l180";
    var folderId = "1ictsEDej7qYc2sI4_udzKp6SPxi3qoYR";
    
    // Build the new file name with the "Edited" suffix.
    // Format: "[Account Number] - [ID Number] - [Last Name] - Edited (yyyy-MM-dd HH:mm:ss)"
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var copyName = accountNumber + " - " + idNumber + " - " + lastName + " - Edited (" + timestamp + ")";
    
    // Retrieve the template file and destination folder.
    var templateFile = DriveApp.getFileById(templateId);
    var folder = DriveApp.getFolderById(folderId);
    
    // Make a new copy of the template in the folder with the new file name.
    var cardFile = templateFile.makeCopy(copyName, folder);
    
    // Open the new document to modify its content.
    var doc = DocumentApp.openById(cardFile.getId());
    var body = doc.getBody();
    
    // Replace the placeholders with the actual values.
    body.replaceText("{{AccountNumber}}", accountNumber);
    body.replaceText("{{ORNumber}}", orNumber);
    body.replaceText("{{LastName}}", lastName);
    body.replaceText("{{FirstName}}", firstName);
    body.replaceText("{{MiddleName}}", middleName);
    body.replaceText("{{IDNumber}}", idNumber);
    body.replaceText("{{College}}", college);
    body.replaceText("{{Degree}}", degree);
    body.replaceText("{{Package}}", chosenPackage);
    body.replaceText("{{TermOfPayment}}", termOfPayment);
    body.replaceText("{{PackagePrice}}", packagePrice);
    
    // Save and close the document.
    doc.saveAndClose();
    
    Logger.log("New card created for " + accountNumber + " at row " + rowNumber);
  }
  