function folderSetup() {
  var longtermFeedback = DriveApp.createFolder('Longterm Feedback');
  var masterFolder = longtermFeedback.createFolder('Master Sheet');
  var folder=DriveApp.getFoldersByName(masterFolder).next();
  var file=SpreadsheetApp.create('Master Sheet')
  var copyFile=DriveApp.getFileById(file.getId());
  folder.addFile(copyFile);
  DriveApp.getRootFolder().removeFile(copyFile);
  var yearFolder = longtermFeedback.createFolder(new Date().getFullYear());
  masterSheetSetup(file);
  }
function masterSheetSetup(spreadsheet){
  spreadsheet.getActiveSheet().setName('Hub');
  spreadsheet.getActiveSheet().setColumnWidth(1, 185).setColumnWidth(2, 568).setColumnWidth(3, 197).setColumnWidth(4, 137).setColumnWidth(5, 203).deleteColumns(6,21);
  spreadsheet.getActiveSheet().setRowHeight(1, 61).setRowHeight(2, 61).setRowHeight(3, 61).deleteRows(121,880);  spreadsheet.getRange('A1').activate().setFontSize(36).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setValue("Step 1:");
  spreadsheet.getRange('B1').activate().setFontSize(36).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setValue("insert student name");
  spreadsheet.getRange('C1').activate().setFontSize(24).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setValue("Year:");
  spreadsheet.getRange('D1').activate().setFontSize(24).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setValue(new Date().getFullYear());
  spreadsheet.getRange('A2').activate().setFontSize(36).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setValue("Step 2:");
  spreadsheet.getRange('B2').activate().setFontSize(36).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setValue("Insert Year");
  spreadsheet.getRange('D2').activate().setFontSize(24).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setValue("Student Name");
  spreadsheet.getRange('E2').activate().setFontSize(24).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setValue("Student Email");
  spreadsheet.getRange('A3').activate().setFontSize(36).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setValue("Step 3:");
  spreadsheet.getRange('B3').activate().setFontSize(36).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setValue("'Press Button");
  newYear(spreadsheet);
  }
function newYear(spreadsheet){
  yearlySpreadsheet = spreadsheet.insertSheet(1).setName('2021');
  yearlySpreadsheet.getRange(1, 1, yearlySpreadsheet.getMaxRows(), yearlySpreadsheet.getMaxColumns()).activate().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  yearlySpreadsheet.setColumnWidth(1, 150).setRowHeight(1, 50).setRowHeights(93, 3, 190);
  yearlySpreadsheet.getRange('A1').setValue("Name:");
  yearlySpreadsheet.getRange('A93:A95').activate().merge().setFontSize(100).setTextRotation(90).setValue("QR Code");
  yearlySpreadsheet.getRange('A7:A92').activate().merge().setFontSize(100).setTextRotation(90).setValue("DATA DATA DATA DATA DATA");
  yearlySpreadsheet.getRange('A1:A6').activate().setFontSize(9);
  yearlySpreadsheet.getRange('A2').activate().setValue("editable google formb:");
  yearlySpreadsheet.getRange('A3').activate().setValue("response google form:");
  yearlySpreadsheet.getRange('A4').activate().setValue("Sheet link:");
  yearlySpreadsheet.getRange('A5').activate().setValue("Email:");
  yearlySpreadsheet.getRange('A6').activate().setValue("Data Type");
  spreadsheet.setFrozenRows(6);
  spreadsheet.getActiveSheet().deleteRows(109, 892);
  spreadsheet.getActiveSheet().deleteColumns(3, 24);
}
