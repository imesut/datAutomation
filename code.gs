var MasterDataSpreadsheetId = "Spreadsheet_address"; // Summary sheet address
var MasterDataSheetId = "Sheet_Name"; // Where summary data will be inserted
var SummaryReportId = "Summary_Sheet"; // From summary mail will be parsed 
var DataFolderId = "folder_id"; // Where individual data files will be moved

function doGet(e) {
  if(typeof e !== 'undefined'){
    Logger.log(e.parameter["file"]);
    file = UrlFetchApp.fetch(e.parameter["file"]).getBlob();
    var excelfile = UrlFetchApp.fetch(e.parameter["file"]).getBlob();
    var filename = Utilities.formatDate(new Date(), "GMT", "yyMMdd");
    var fileInfo = {
      title: filename,
      mimeType: "MICROSOFT_EXCEL", // it'll be converted to Google Sheets
      "parents": [{'id': DataFolderId}], //Folder id for Reports folder
    };
    file = Drive.Files.insert(fileInfo, excelfile, {convert: true});
    addQueryToSheet(file.id);
    cleanFormula(file.id);
    appendToDataSheet(file.id);
    sendSummaryMailTo("mail@mail.com", "Report Summary of " + filename);
    return ContentService.createTextOutput(JSON.stringify(e.parameter));
    }
}

function addQueryToSheet(SpreadsheetId){
  var source = SpreadsheetApp.openById(""); //query sheet address
  var sheet = source.getSheets()[0];
  var destination = SpreadsheetApp.openById(SpreadsheetId); //target sheet id
  sheet.copyTo(destination);
}

function cleanFormula(SpreadsheetId){ // Instead of this formula, .getValue() method is also be valid. But for my project it is required to store data as value only
  var spreadsheet = SpreadsheetApp.openById(SpreadsheetId);
  var sheetNumber = spreadsheet.getNumSheets();
  var lastSheet = spreadsheet.getSheets()[sheetNumber-1];
  var source = lastSheet.getRange("A1:X12");
  source.copyTo(lastSheet.getRange("A1"), {contentsOnly: true});
}

function appendToDataSheet(SpreadsheetId){
  var usageDataSheet = SpreadsheetApp.openById(MasterDataSpreadsheetId).getSheetByName(MasterDataSheetId); //Address for usage data sheet
  var spreadsheet = SpreadsheetApp.openById(SpreadsheetId);
  var sheetNumber = spreadsheet.getNumSheets();
  var lastSheet = spreadsheet.getSheets()[sheetNumber-1];
  var data = lastSheet.getRange("G3:X3").getValues()[0]; //Address for usage data sheet
  usageDataSheet.appendRow(data);
}

function sendSummaryMailTo(address, subject){
  var message = SpreadsheetApp.openById(MasterDataSpreadsheetId).getSheetByName(SummaryReportId).getRange("E1").getValue();
  MailApp.sendEmail(address, subject, message);
}















