/** @NotOnlyCurrentDoc */
function Create_New_Table() {
  var Startspreadsheet = SpreadsheetApp.getActive();
  

  var Profiles = []
  var Row = 3
  var cellN = 'd3'
  while ((Startspreadsheet.getRange(cellN).getValue()).length > 1){
    var cellN = 'd'+ Row
    var cellE = 'e' + Row
    var Profile = []
    if ((Startspreadsheet.getRange(cellN).getValue()).length > 0);
      var Name = (Startspreadsheet.getRange(cellN).getValue())
      var Email = (Startspreadsheet.getRange(cellE).getValue())
    var Row = Row + 1
    Profile.push(Name,Email)
    Profiles.push(Profile)
  }
  Profiles.splice(-1, 1);

  console.log(Profiles)
  var sheetname = Startspreadsheet.getRange("D1").getValue();
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName(sheetname);
  while (Profiles.length > 0){
    console.log(Profiles.length)
    var studentName = Profiles[0][0]
    var spreadsheet = SpreadsheetApp.getActive().getSheetByName(sheetname);
    spreadsheet.getRange('B1').activate();
    var spreadsheet = SpreadsheetApp.getActive()
    spreadsheet.getActiveSheet().insertColumnsBefore(spreadsheet.getActiveRange().getColumn(), 3);
    spreadsheet.getActiveRange().offset(0, 0, spreadsheet.getActiveRange().getNumRows(), 3).activate();
    spreadsheet.getRange('B1:D5').activate().mergeAcross();
    spreadsheet.getRange('B106:D108').activate().merge();
    spreadsheet.getRange('B1:D1').activate().setValue(studentName + "'s Longterm");
    spreadsheet.getRange('B5').activate().setValue(Profiles[0][1]);
    spreadsheet.getRange('B6').activate().setValue('Timestamp');
    spreadsheet.getRange('C6').activate().setValue('How would you rate the presentation?');
    spreadsheet.getRange('D6').activate().setValue('Do you have any feedback for the presenter?');
    spreadsheet.getRange('B7').activate().setFormula('=importrange(b4, "a2:c100")');
    spreadsheet.getActiveSheet().setColumnWidth(3, 238).setColumnWidth(3, 238);
    spreadsheet.getActiveSheet().setColumnWidth(4, 284);
    var yearFolder = DriveApp.getFoldersByName(new Date().getFullYear()).next();
    var studentFolder = yearFolder.createFolder(studentName);
    FormApp.getActiveForm()
    var form = FormApp.create(studentName + "'s Longterm");
    var formId = form.getId()

    var copyFile=DriveApp.getFileById(formId);
    studentFolder.addFile(copyFile);
    DriveApp.getRootFolder().removeFile(copyFile);
    form.addScaleItem().setTitle('How would you rate the presentation').setBounds(1, 10);
    form.addParagraphTextItem().setTitle('Do you have any feedback for the presenter?');
    spreadsheet.getRange('B2').activate().setValue(form.getEditUrl()); 
    spreadsheet.getRange('B3').activate().setValue(form.getPublishedUrl()); 	
    var studentSpreadsheet = SpreadsheetApp.create(studentName + "'s Longterm (Responses)").addEditor(Profiles[0][1]);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, studentSpreadsheet.getId());
    var copyFile=DriveApp.getFileById(studentSpreadsheet.getId());
    studentFolder.addFile(copyFile);
    DriveApp.getRootFolder().removeFile(copyFile);

    GmailApp.sendEmail(Profiles[0][1], "Longterm Feedback", "Here is your link to open your Longterm Feedback Spreadsheet and your QR code. Please put the QR code at the end of your presentation. \n\nSpreadsheet: " + studentSpreadsheet.getUrl() + '\n\nQR Code: "https://image-charts.com/chart?chs=150x150&cht=qr&choe=UTF-8&chl=' + form.getPublishedUrl());
    spreadsheet.getRange('B4').activate().setValue(studentSpreadsheet.getUrl()); 
    spreadsheet.getRange('B193').activate().setValue('=if(isblank(B3), "No url, this is very bad and means Sean probably messed up. Way to go Sean.", image("https://image-charts.com/chart?chs=150x150&cht=qr&choe=UTF-8&chl="&ENCODEURL(B3)))');
    Profiles.splice(0, 1);
  }
};
