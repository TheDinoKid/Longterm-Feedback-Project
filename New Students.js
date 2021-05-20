/** @NotOnlyCurrentDoc */
function getNames() {
  //Variable Defining
  var profiles = []
  var names = []
  var emails = []
  var hubSheet = SpreadsheetApp.getActive().getSheetByName("hub");
  
  //pulls all possible slots where names and emails could be
  var data =(hubSheet.getRange('d3:e').getValues())
  
  //splits them into seperate lists
  for (i=0;i!=data.length;i++){
    names.push(data[i][0])
    emails.push(data[i][1])
  }
  
  //removes all blank space
  var lastRow = names.filter(String).length
  names.splice(lastRow)
  emails.splice(lastRow)
  profiles.push(names,emails)
  }
