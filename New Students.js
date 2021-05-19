/** @NotOnlyCurrentDoc */
function getNames() {
  var profiles = [[],[]]
  var hubSheet = SpreadsheetApp.getActive().getSheetByName("hub");
  var times = [new Date().getTime()]
  var values = hubSheet.getRange('D:D').getValues();
  for (row = 3; values[row-1] != '' && row-1 != values.length;row++){}
  var values = hubSheet.getRange('d3:e'+(row-1)).getValues()
  for (i = 0; values.length !=i; i++){
    profiles[0].push(values[i][0])
    profiles[1].push(values[i][1])}
  times.push(new Date().getTime())
  //forms(profiles,hubSheet,times)
  console.log (times[1]-times[0])
  }
  function rerun(){for (c=0;c!=25;c++){getNames()}}
