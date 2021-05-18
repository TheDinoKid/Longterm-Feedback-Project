/** @NotOnlyCurrentDoc */

//1k tests//(3s-6s)getNames 35 test average = 4.3
function getNames() {
  var times = []
  times.push(new Date().getTime())
  var hubSheet = SpreadsheetApp.getActive().getSheetByName("hub");
  for (row = 3; (hubSheet.getRange('d'+ row).getValue()).length != 0 ; row++){}
  profiles = [hubSheet.getRange('d3:d'+(row-1)).getValues()]
  profiles.push(hubSheet.getRange('e3:e'+(row-1)).getValues())
  times.push(new Date().getTime())
  format(profiles,hubSheet,times)
  }
//1k tests//
function format(profiles,hubSheet,times){
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName(hubSheet.getRange("D1").getValue());
    spreadsheet.insertColumnsBefore(2,3*profiles[0].length)
    var valuesB = [['Timestamp'],['=importrange(b4, "a2:c100")']]
  for (i = 0; profiles[0].length != i; i++){
    spreadsheet.getRange(1,2+(i*3),5,3).activate().mergeAcross();
    spreadsheet.getRange(93,2+(i*3),3,3).activate().merge();
    spreadsheet.setColumnWidth(3+(i*3), 238).setColumnWidth(4+(i*3), 284)
    spreadsheet.getRange(6,2+(i*3),2,1).setValues(valuesB);
    spreadsheet.getRange(1,2+(i*3),1,3).setValue(profiles[0][i] + "'s Longterm");
    spreadsheet.getRange(5,2+(i*3),1,3).setValue(profiles[1][i]);
    spreadsheet.getRange(6,3+(i*3),1,1).setValue('How would you rate the presentation?');
    spreadsheet.getRange(6,4+(i*3),1,1).setValue('Do you have any feedback for the presenter?');
    spreadsheet.getRange(93,2+(i*3),3,3).setValue('=if(isblank(B3), "No url, this is very bad and means Sean probably messed up. Way to go Sean.", image("https://image-charts.com/chart?chs=150x150&cht=qr&choe=UTF-8&chl="&ENCODEURL(B3))');}
    times.push(new Date().getTime())
    form(profiles,hubSheet,spreadsheet,times)}


function form(profiles,hubSheet,spreadsheet,times){
      //need FormApp.getActiveForm()??
      FormApp.getActiveForm()
      for (i = 0; profiles[0].length != i; i++){
        var studentFolder = (DriveApp.getFoldersByName(new Date().getFullYear()).next().createFolder(profiles[0][i]))
        var form = ((FormApp.create(profiles[0][i] + "'s Longterm")));
        form.addScaleItem().setTitle('How would you rate the presentation').setBounds(1, 10);
        form.addParagraphTextItem().setTitle('Do you have any feedback for the presenter?');
        studentFolder.getId().addFile(DriveApp.getFileById(form.getId()));
        DriveApp.getRootFolder().removeFile(DriveApp.getFileById(form.getId()));
        
        }
      //var studentSpreadsheet = SpreadsheetApp.create(studentName + "'s Longterm (Responses)").addEditor(Profiles[0][1]);
      //form.setDestination(FormApp.DestinationType.SPREADSHEET, studentSpreadsheet.getId());
      //studentFolder.addFile(DriveApp.getFileById(studentSpreadsheet.getId()));
      //DriveApp.getRootFolder().removeFile(DriveApp.getFileById(studentSpreadsheet.getId()));
      times.push(new Date().getTime())
      console.log('getNames:',(times[2]-times[1])/1000,'s      format:',(times[3]-times[2])/1000,'s      forms:',(times[4]-times[3])/1000,'s      total:',(times[4]-times[1])/1000)}
