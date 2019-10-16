link_answer = 'https://docs.google.com/spreadsheets/d/13cpAjYAH21OkqMSqNmXm0VjbYnTFTqX1QOZdyMyolpU/edit?usp=sharing';
function getNewNames(){
  var form = FormApp.getActiveForm();
  var items = form.getItems();

  var ss = SpreadsheetApp.openByUrl(link_answer);
    //var assSheet1 = ss.getSheetByName('Sheet1');
    var assSheet1 = ss.getSheets()[1];
    var assValues = assSheet1.getDataRange().getValues();
    var values = assValues.slice(1);// get col 2
    var valSort = values.sort(); // sort data 

  var ss2 = SpreadsheetApp.openByUrl(link_answer);
     var playerSheet1 = ss2.getSheets()[0];
     var playerValues = playerSheet1.getDataRange().getValues(); 
     var player = playerValues.slice(0);
     var plySort = player.sort();

  var names = [];

  for(var p = 0; p < plySort.length; p++){
    names.push(plySort[p][0])
  }

  var pList = items[0].asListItem();
  pList.setChoiceValues(names).setRequired(true).setHelpText('Please Select Player Name')

  var areas = [];
  for (var i = 0; i < valSort.length; i++) {
      areas.push(valSort[i][1])
    }

  var aList = items[1].asMultipleChoiceItem();
      aList.setChoiceValues(areas).setRequired(true).setHelpText('Select Area of Assessment')

}