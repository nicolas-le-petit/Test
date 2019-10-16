link_Data_base = 'https://docs.google.com/spreadsheets/d/1Inc1YFK9WGrqyUNMLNvv3syiQIjFJJ80ZasZWke9hnw/edit?usp=sharing';
sheet_tutor = 'Tutor_register';
sheet_student = 'Student_register';

function getNewNames(){
  var form = FormApp.getActiveForm();
  var items = form.getItems();
  
//    var ss = SpreadsheetApp.openByUrl(
//     'https://docs.google.com/spreadsheets/d/1Inc1YFK9WGrqyUNMLNvv3syiQIjFJJ80ZasZWke9hnw/edit');
//    var assSheet1 = ss.getSheetByName('Form Responses 1');
//    var assValues = assSheet1.getRange("C:C").getDataRegion().getValues();
//    var values = assValues.slice(0);
//    var valSort = values.sort();
//
//  var ss2 = SpreadsheetApp.openByUrl(
//     'https://docs.google.com/spreadsheets/d/1Inc1YFK9WGrqyUNMLNvv3syiQIjFJJ80ZasZWke9hnw/edit');
//     var playerSheet1 = ss2.getSheetByName('Form Responses 1');
//     var playerValues = playerSheet1.getRange("D:D").getDataRegion().getValues();
//     var player = playerValues.slice(0);
//     var plySort = player.sort();

  var names = [];

  for(var p = 0; p < plySort.length; p++){
    names.push(plySort[p][0])
  }

  var pList = items[6].asListItem();
  pList.setChoiceValues(names).setRequired(true).setHelpText('Please Select Name');

  var areas = [];
  for (var i = 0; i < valSort.length; i++) {
      areas.push(valSort[i][0])
    }

  var aList = items[7].asListItem();
      aList.setChoiceValues(areas).setRequired(true).setHelpText('Select ID');

}

function getColbyTitle(link, sheetName, colTitle){
  var ss = SpreadsheetApp.openByUrl(link);
  var s = ss.getSheetByName(sheetName);
  var dataRange = s.getDataRange();
  var data = dataRange.getValues();
  
  for (var i=0; i<data[0].length; i++){
    if (data[0][i] == colTitle)
    {
      return i;
    }
  }
}

function getDatabyCol(link, sheetName, colTitle){
  var ss = SpreadsheetApp.openByUrl(link);
  var s = ss.getSheetByName(sheetName);
  var dataRange = s.getDataRange();
  var data = dataRange.getValues();
  
  var coldata = [];
  
  for (var i=0; i<data[0].length; i++){
    if (data[0][i] == colTitle)
    {
      for (var j=1; j<data.length; j++){
        if (data[j][i]){
          coldata.push(data[j][i]);
        }
      }
    }
  } 
  return coldata;
}

function merge_infor (link, sheetName)
{
  var ss = SpreadsheetApp.openByUrl(link);
  var s = ss.getSheetByName(sheetName);
  var dataRange = s.getDataRange();
  var data = dataRange.getValues();

  for (var i = 1; i < data.length ; i++) {
    if (data[i][4]=='') {
      //data[i][4] = data[i][1] + ' - ' + data[i][2].dateformat ('dd-m-yy') + ' - ' + data[i][3];
      var giasu = data[i][1];
      var ngay = data[i][2].getDate()+1;
      var thang = data[i][2].getMonth()+1;
      var gio = data[i][3].getHours()+1 + 'h';
      data[i][4] = giasu + ' - ' + ngay +'/'+ thang + ' - ' + gio;
      s.getRange(i+1, 5).setValue(data[i][4]);
    }
  }
}

function update_form(){
  merge_infor(link_Data_base,sheet_tutor);
  var form_tutor = FormApp.getActiveForm();
  
  var tutor = getDatabyCol(link_Data_base,sheet_tutor,"Merge Information");
 
  var form_student = FormApp.openByUrl('https://docs.google.com/forms/d/1vhPbzWI9C8VXkQprNzgpSJ8m0OV2LGMNfcD2WZRDRic/edit');
  var items = form_student.getItems();
  
  for (var i=0; i<items.length; i++){
    if (items[i].getTitle() == "Please Select Your Tutor"){
      var pList = items[i].asMultipleChoiceItem();
      pList.setChoiceValues(tutor).setRequired(true);
      //Browser.msgBox(i);
      break;
    }
  }


//  var pList = items[6].asListItem();
//  pList.setChoiceValues(names).setRequired(true).setHelpText('Please Select Name');

//  var areas = getDatabyCol(link_Data_base,'Form Responses 1','Please input ID');
//
//  var aList = items[7].asListItem();
//      aList.setChoiceValues(areas).setRequired(true).setHelpText('Select ID');
}