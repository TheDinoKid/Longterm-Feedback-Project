function folderSetup() {
  var longtermFeedback = DriveApp.createFolder('Longterm Feedback');
  var masterFolder = longtermFeedback.createFolder('Master Sheet');
  var file=SpreadsheetApp.create('Master Sheet')
  masterFolder.addFile(DriveApp.getFileById(file.getId()));
  DriveApp.getRootFolder().removeFile(DriveApp.getFileById(file.getId()));
  longtermFeedback.createFolder(new Date().getFullYear());
  masterSheetSetup(file);
  }
function masterSheetSetup(spreadsheet){
  spreadsheet.getActiveSheet().setName('Hub').setColumnWidth(1, 185).setColumnWidth(2, 568).setColumnWidth(3, 197).setColumnWidth(4, 137).setColumnWidth(5, 203).setRowHeight(1, 61).setRowHeight(2, 61).setRowHeight(3, 61).deleteColumns(6,21);
  spreadsheet.getActiveSheet().deleteRows(121,880); 
  var values = [["Step 1:","insert student name","Year:",(new Date().getFullYear()),""],["Step 2:","Insert Year","","Student Name","Student Email"],["Step 3:","'Press Button","","",""]]
  spreadsheet.getRange('A1:E3').setValues(values); 
  spreadsheet.getRange('A1:B3').activate().setFontSize(36).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  spreadsheet.getRange('C1:E2').activate().setFontSize(24).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  newYear(spreadsheet);
  }
function newYear(spreadsheet){
  yearlySpreadsheet = spreadsheet.insertSheet(1).setName('2021');
  yearlySpreadsheet.getRange(1, 1, yearlySpreadsheet.getMaxRows(), yearlySpreadsheet.getMaxColumns()).activate().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  yearlySpreadsheet.setColumnWidth(1, 150).setRowHeight(1, 50).setRowHeights(93, 3, 190).setFrozenRows(6);
  yearlySpreadsheet.getRange('A93:A95').activate().merge().setFontSize(100).setTextRotation(90).setValue("QR Code");
  yearlySpreadsheet.getRange('A7:A92').activate().merge().setFontSize(100).setTextRotation(90).setValue("DATA DATA DATA DATA DATA");
  yearlySpreadsheet.getRange('A1:A6').setFontSize(9);
  var values = [["Name:"],["editable google form:"],["response google form:"],["Sheet link:"],["Email:"],["Data Type"]]
  yearlySpreadsheet.getRange('A1:A6').setValues(values);
  spreadsheet.getActiveSheet().deleteRows(109, 892);
  spreadsheet.getActiveSheet().deleteColumns(3, 24);
}
