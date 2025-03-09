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
      generateCardForRow(i + 1);
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
 *   "Middle Name", "Full ID Number", "College", "Degree Code",
 *   "Alternate Degree Code", "Chosen Package", "Term of Payment"
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
  var idNumber = getCellValue("Full ID Number");
  var college = getCellValue("College");
  
  // Use "Alternate Degree Code" if "Degree Code" equals "MY DEGREE CODE ISN'T IN THE LIST".
  var degree = getCellValue("Degree Code");
  if (degree === "MY DEGREE CODE ISN'T IN THE LIST" && headerMapping["Alternate Degree Code"] !== undefined) {
    degree = getCellValue("Alternate Degree Code");
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
