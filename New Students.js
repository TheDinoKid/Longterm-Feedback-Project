/** @NotOnlyCurrentDoc */
function getNames() {
  var profiles = []
  var hubSheet = SpreadsheetApp.getActive().getSheetByName("hub");
  var times = [new Date().getTime()]
  var names = []
  var emails = []
  var data =(hubSheet.getRange('d3:e').getValues())
  for (i=0;i!=data.length;i++){
    names.push(data[i][0])
    emails.push(data[i][1])
  }
  var lastRow = names.filter(String).length
  names.splice(lastRow)
  emails.splice(lastRow)
  profiles.push(names,emails)
  times.push(new Date().getTime())
  console.log (times[1]-times[0])
  }
