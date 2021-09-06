/*
This script acrchives the cellar log from a tab on a google sheet to a new
file on google drive.  The archive is organized by batch number here:
https://drive.google.com/drive/folders/1LFc9T7tti7M300D997j8XbM9FqUgpXyM

The script finishes by placing an empty cellar log in place of the one
that was archived.
*/

function ArchiveCellarLog() {

  // Identify cellar log sheet to archive
  var sheetToCopy = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var theLongSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();

  // Get beer name
  var theBeerName = sheetToCopy.getRange('B2:B2').getValue();

  // Remove "FV " from theLongFermentorNumber
  var theSheetName = theLongSheetName.replace("FV ", "") ;

  // Prompt user for the fermentor number to archive
  var theUserEnteredFermentorNumber = displayPrompt();
  
  // Exit the program if the user entered a different fermentor number from the active sheet
  if (theSheetName != theUserEnteredFermentorNumber) {
    SpreadsheetApp.getUi().alert('Please type a number matching the active sheet. Exiting...');
    return;
  }

  // Get batch number of the brew in this fermentor
  var theLongBatchNumber = sheetToCopy.getRange('D2:D2').getValue();

  // Remove the trailing .0 from theBatchNumber
  var theBatchNumber = theLongBatchNumber.replace(".0", "") ;

  // Create archive folder using the batch number
  // Move it to the 2021 archive and trash the original
  DriveApp.createFolder(theBatchNumber)
  var folder = DriveApp.getFoldersByName(theBatchNumber).next();
  var destination = DriveApp.getFoldersByName("2021").next();
  destination.addFolder(folder);
  folder.getParents().next().removeFolder(folder);

  // Set the destination folder for the archive (the folder must already exist)
  //var destFolder = DriveApp.getFoldersByName("21001").next();
  var destFolder = DriveApp.getFoldersByName(theBatchNumber).next();

  // Get the Id for the template file that will be copied to the archive folder
  // and appended with a sheet containing the cellar log for this batch and fermentor  
  // The template file is named CellarLogArchive Template
  // It lives in the Batch Archive folder here:
  //https://docs.google.com/spreadsheets/d/1JHcsQIpgM2eCM06_ukLQfT2xJGKCIK91FATQn2bOwEo/edit?usp=sharing
  
  var templateSheetId = "1JHcsQIpgM2eCM06_ukLQfT2xJGKCIK91FATQn2bOwEo"

  // Make the copy
  //DriveApp.getFileById(templateSheetId).makeCopy(theBatchNumber & " " & theBeerName, destFolder);
  DriveApp.getFileById(templateSheetId).makeCopy(theBatchNumber, destFolder); 

  // Get Id for the newly made copy
  var destinationId = DriveApp.getFilesByName(theBatchNumber).next().getId();
  
  // Set destination to receive the copy of the cellar log
  var destination = SpreadsheetApp.openById(destinationId);

  // Make the cellar log copy in the batch archive
  sheetToCopy.copyTo(destination);

  // Delete sheet "Blank"

 //var files = DriveApp.getFilesByName('Finances');
 var files = DriveApp.getFilesByName(theBatchNumber);
  while (files.hasNext()) {
  var spreadsheet = SpreadsheetApp.open(files.next());
  //var sheet = spreadsheet.getSheets()[0];
  var sheet = spreadsheet.getSheetByName("Blank")
  //Logger.log(sheet.getName());
  spreadsheet.deleteSheet(sheet)
}

// Remove "Copy of" from sheet name
var sheet = spreadsheet.getSheets()[0];
sheet.setName("FV " + theUserEnteredFermentorNumber)

// Append beer name to the spreadsheet
spreadsheet.rename(theBatchNumber + " " + theBeerName)

  /**********************************************************
  Reset the cellar log with an empty sheet for this fermentor
  ***********************************************************/

  // Get batch number in this fermentor
  var theLongFermentorNumber = sheetToCopy.getRange('A2:A2').getValue();

  // Remove "FV " from theLongFermentorNumber
  var theFermentorNumber = theLongFermentorNumber.replace("FV ", "") ;

  // Delete the old fermentor sheet (active sheet )and store the new active sheet in a variable
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet();

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Template'), true);
  spreadsheet.duplicateActiveSheet();
  spreadsheet.getActiveSheet().setName('FV ' + theFermentorNumber);
  
  // Very important to set the Active Sheet
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('FV ' + theFermentorNumber), true);
 
  // Move the new sheet to its proper place
  switch (theFermentorNumber)  {
        case "A1":
            spreadsheet.moveActiveSheet(30);
            break;
        case "A2":
            spreadsheet.moveActiveSheet(31);
            break;
        case "A3":
            spreadsheet.moveActiveSheet(32);
            break;
         case "A4":
            spreadsheet.moveActiveSheet(33);
            break;
        default:
           spreadsheet.moveActiveSheet(theFermentorNumber);  
            break;
    }
  
  // Write fermentor number on the sheet
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('FV ' + theFermentorNumber), true);
  spreadsheet.getRange('A2').activate();
  spreadsheet.getCurrentCell().setValue('FV ' + theFermentorNumber);

};
