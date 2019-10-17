/*
@author: Dung Tran
@date  : 1 Oct 19
@define link to sheet & form relate
@NEED TO CLEAN CODE
*/

link_Data_base = 'https://docs.google.com/spreadsheets/d/1BO9dr4feFbfD4vAcB-h_ltF9GHKL3pZKhVdtR5O4gEM/edit?usp=sharing';
sheet_tutor = 'Tutor_register';
sheet_student = 'Learner_register';
form_student_id = 'https://docs.google.com/forms/d/1Cit5s4LO12IFWOi910oJuY2SNsTayv5xTJsHv15YtMg/edit';

/*
@author: Dung Tran
@date  : 1 Oct 19
@brief : to get num of col by header
@para  : link to sheet, sheet name, name of header  
*/
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

/*
@author: Dung Tran
@date  : 1 Oct 19
@brief : to get data from col by header
@para  : link to sheet, sheet name, name of header  
*/
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

/*
@author: Dung Tran
@date  : 1 Oct 19
@brief : to merge data from several cells
@para  : link: link to sheet,
         sheetName, 
blank -> 1
not blank + not book -> 0
not blank + book -> 1
not blank + booked -> 0
*/

function merge_infor (link, sheetName)
{
  var ss = SpreadsheetApp.openByUrl(link);
  var s = ss.getSheetByName(sheetName);
  var dataRange = s.getDataRange();
  var data = dataRange.getValues();

  for (var i = 1; i < data.length ; i++) {
    //not blank + not book -> 0
    if (data[i][6].search("Booked by") != (-1)) {continue;}
    //not blank + booked -> 0
    if (data[i][6] != "" & data[i][5] =="") {continue;}
    if (data[i][6] != "" & data[i][5] !="") {data[i][6]="";}
    data[i][2] = (data[i][2].getDate()) + '/' + (data[i][2].getMonth()+1);
    for (var j = 1; j <= 5; j++) {
      if (j==3) {data[i][6] += data[i][j] + ' - ';}
      else if (j==5) {data[i][6] += data[i][j];} 
      else {data[i][6] += data[i][j] + ' | ';}
    }
    s.getRange(i+1, 7).setValue(data[i][6].trim());
      // var giasu = data[i][1];
      // var ngay = data[i][2].getDate()+1;
      // var thang = data[i][2].getMonth()+1;
      // var gio_start = data[i][3];
      // var gio_end = data[i][4];
      // var lack = data[i][6];
      // var dau_cach = " | ";
      // data[i][6] = giasu + dau_cach + ngay +'/'+ thang + dau_cach + gio_start + '-' + gio_end + dau_cach + lack;
      // s.getRange(i+1, 7).setValue(data[i][6]);
  }
  //s.getRange(i+1, 7).setValue(data[i][6]);
}

/*
@author: Dung Tran
@date  : 1 Oct 19
@brief : to update form by data from sheet
@para  : none 
*/

function update_form(){
  merge_infor(link_Data_base,sheet_tutor);
  var form_tutor = FormApp.getActiveForm();
  
  var tutor = getDatabyCol(link_Data_base,sheet_tutor,"Merge Information");
 
  var form_student = FormApp.openByUrl(form_student_id);
  var items = form_student.getItems();
  
  for (var i=0; i<items.length; i++){
    if (items[i].getTitle() == "Please Select Your Tutor"){
      var pList = items[i].asMultipleChoiceItem();
      pList.setChoiceValues(tutor).setRequired(true);
      break;
    }
  }
}

/*
@author: Dung Tran
@date  : 2 Oct 19
@brief : to send mail from data
@para  : link to sheet
         sheet name
         reference data 
*/
function send_mail(link, sheetName, reference)
{
  var ss = SpreadsheetApp.openByUrl(link);
  var s = ss.getSheetByName(sheetName);
  var dataRange = ss.getDataRange();
  var data = dataRange.getValues();

  for (var rowdata = 0; rowdata<data.length; rowdata++){
    if (data[rowdata][1] == coacher){
      var email1 = data3[rowdata][3];//coacher
      var email2 = data3[rowdata][4];//SH
      var email3 = data3[rowdata][5];//TPM/POD?
              
      if (email2 != "" && email3 != "")
        var email_cc = email2 + "," + email3 + "," + email4;
              
      if (email2 == "" && email3 != "")
        var email_cc = email4 + "," + email3;
              
      if (email2 != "" && email3 == "")
        var email_cc = email2 + "," + email4;
              
      if (email2 == "" && email3 == "")
        var email_cc = email4;
                
      MailApp.sendEmail(email1, subject, message, {
                            cc: email_cc});//gui mail cho hoc vien
        break;
    }
  }
}