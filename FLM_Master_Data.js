/*
@Date: 20/6/2018
@Author: DungTT
@brief: for L&D Section - HVB
*/
LAF = ["Coaching","Time Management","Communication","Influence","Delegation"];
APF = ["Motivation","Stress Management","Teamwork","Performance Management","Decision Making","Meeting skill"]; 

function get_data_Learning_Application(){
  
  //var APF = ["Motivation","Stress Management","Teamwork","Performance Management","Decision Making","Meeting skill"];        
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var staff_name = sheet.getRange("C2").getValue();
  
  
  //get data Learning Application 
  var offset = 0;
  for (var row = 0; row < 5; row++){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LAF[row]);
  var dataRange = sheet.getDataRange().getValues();
  
  //check staff name
  var temp = 0;
  for (var rowdata = 0; rowdata<dataRange.length; rowdata++){
    if (dataRange[rowdata][0] == staff_name){
      temp = 1;
      //pasteSheet.getRange("G2").setValue(row);
      //tim duoc row chua data nhan vien 
      break;
    }
  }
  
  var check1=0;
  var check2=0;
  //copy data
  if (sheet.getRange(rowdata+1,13).getValue()==""){
    check1=1;//manager chua cmt
    var data_range1 = sheet.getRange(rowdata+1,6,3,7).getValues();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Data").getRange(offset+38,3,3,7).setValues(data_range1);
  }
  else{
    var data_range1 = sheet.getRange(rowdata+1,6,3,8).getValues();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Data").getRange(offset+38,3,3,8).setValues(data_range1);
  }
  
  if (sheet.getRange(rowdata+4,21).getValue()==""){
    check2=1;//manager chua cmt
    var data_range2 = sheet.getRange(rowdata+4,14,3,7).getValues();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Data").getRange(offset+42,3,3,7).setValues(data_range2);
  }
  else{
    var data_range2 = sheet.getRange(rowdata+4,14,3,8).getValues();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Data").getRange(offset+42,3,3,8).setValues(data_range2);
  }
    
    offset+= 8;
 }  
}

function get_data_Action_Plan(){       
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var staff_name = sheet.getRange("C2").getValue();
 
  //get data Action Plan
  var offset2 = 0;
  for (var row = 0; row < 6; row++){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APF[row]);
  var dataRange = sheet.getDataRange().getValues();
  
  //check staff name
  var temp = 0;
  for (var rowdata = 0; rowdata<dataRange.length; rowdata++){
    if (dataRange[rowdata][0] == staff_name){
      temp = 1;
      //pasteSheet.getRange("G2").setValue(row);
      //tim duoc row chua data nhan vien 
      break;
    }
  }
  
  var check3=0;
  //copy data
  if (sheet.getRange(rowdata+1,13).getValue()==""){
    check3=1;//manager chua cmt
    var data_range1 = sheet.getRange(rowdata+1,6,4,7).getValues();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Data").getRange(offset2+10,3,4,7).setValues(data_range1);
  }
  else{
    var data_range1 = sheet.getRange(rowdata+1,6,4,8).getValues();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Data").getRange(offset2+10,3,4,8).setValues(data_range1);
  }
    
    offset2+= 4;
 }  
}


function check_update_action_plan()
{
  var Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_remind = Spreadsheet.getSheetByName("Remind_List");

  
}

/*======================================================================================
function check_name(){ 
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Remind_List");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Coaching");
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  var lastRow = data.length; 
  var row = 1;
  for (var i = lastRow; i>=0; i--){
    masterSheet.getRange(row, 1).setValue(data[i][0]);
    row++;
  }
  
  removeDuplicates("B:B");
}

function removeDuplicates() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getRange("B:B").getValues();
  var newData = new Array();
  for(i in data){
    var row = data[i];
    var duplicate = false;
    for(j in newData){
      if(row.join() == newData[j].join()){
        duplicate = true;
      }
    }
    if(!duplicate){
      newData.push(row);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length)
      .setValues(newData);
}

=================================================================================*/

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menu = [ 
    {name: "1. Get Data of Learning Application", functionName: "get_data_Learning_Application"},
    {name: "2. Get Data of Action Plan", functionName: "get_data_Action_Plan"},
    null,
  ];  
  ss.addMenu("L&D Tab", menu);
}
