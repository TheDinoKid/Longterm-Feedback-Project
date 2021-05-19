/** @NotOnlyCurrentDoc */
function getNames() {
  var times = []
  times.push(new Date().getTime())
  var hubSheet = SpreadsheetApp.getActive().getSheetByName("hub");
  for (row = 3; (hubSheet.getRange('d'+ row).getValue()).length != 0 ; row++){}
  var profiles = [hubSheet.getRange('d3:d'+(row-1)).getValues()]
  profiles.push(hubSheet.getRange('e3:e'+(row-1)).getValues())
  times.push(new Date().getTime())
  forms(profiles,hubSheet,times)
  }

function forms(profiles,hubSheet,times){
      var editFormUrl = []
      var pubFormUrl = []
      var sheetUrl = []
      for (i = 0; profiles[0].length != i; i++){
        var studentFolder = (DriveApp.getFoldersByName(new Date().getFullYear()).next().createFolder(profiles[0][i]))
        var form = ((FormApp.create(profiles[0][i] + "'s Longterm")));
        form.addScaleItem().setTitle('How would you rate the presentation').setBounds(1, 10)
        form.addParagraphTextItem().setTitle('Do you have any feedback for the presenter?')
        form.setRequireLogin(false);
        var formid = DriveApp.getFileById(form.getId())
        studentFolder.addFile(formid).removeFile(formid);
        editFormUrl.push(form.getEditUrl())
        pubFormUrl.push(form.getPublishedUrl())
        }
        times.push(new Date().getTime())
        for (i = 0; profiles[0].length != i; i++){
        var studentSheetId = DriveApp.getFileById((SpreadsheetApp.create((profiles[0][i]) + "'s Longterm (Responses)")).getId());
        studentFolder.addFile(studentSheetId).removeFile(studentSheetId);
        sheetUrl.push(studentSheetId.getUrl())
        }
      times.push(new Date().getTime())
      profiles.push([editFormUrl],[pubFormUrl],[sheetUrl])
      format(profiles,hubSheet,times)}

function format(profiles,hubSheet,times){
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName(hubSheet.getRange("D1").getValue());
    spreadsheet.insertColumnsBefore(2,3*profiles[0].length)
    var valuesB2 = [['Timestamp'],['=importrange(b4, "a2:c100")']]
  for (i = 0; profiles[0].length != i; i++){
    spreadsheet.getRange(1,2+(i*3),5,3).activate().mergeAcross();
    spreadsheet.getRange(93,2+(i*3),3,3).activate().merge();
    spreadsheet.setColumnWidth(3+(i*3), 238).setColumnWidth(4+(i*3), 284)
    var valuesB1 = [[profiles[2][i]],[profiles[3][i]],[profiles[4][i]]]
    spreadsheet.getRange(2,2+(i*3),3,1).setValues(valuesB1);
    spreadsheet.getRange(6,2+(i*3),2,1).setValues(valuesB2);
    spreadsheet.getRange(1,2+(i*3),1,3).setValue(profiles[0][i] + "'s Longterm");
    spreadsheet.getRange(5,2+(i*3),1,3).setValue(profiles[1][i]);
    spreadsheet.getRange(6,3+(i*3),1,1).setValue('How would you rate the presentation?');
    spreadsheet.getRange(6,4+(i*3),1,1).setValue('Do you have any feedback for the presenter?');
    var qrcode = ('=image("https://image-charts.com/chart?chs=150x150&cht=qr&choe=UTF-8&chl="&ENCODEURL("'+profiles[3][i]+'"))')
    spreadsheet.getRange(93,2+(i*3),3,3).setValue(qrcode);}
    times.push(new Date().getTime())
    console.log('getNames:',(times[1]-times[0]),'ms      forms:',(times[2]-times[1])/1000,'s      sheets:',(times[3]-times[2])/1000,'s      format:',(times[4]-times[3]),'ms      total:',(times[4]-times[0])/1000,'s')}
