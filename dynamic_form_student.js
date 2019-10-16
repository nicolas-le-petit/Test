link_Data_base = 'https://docs.google.com/spreadsheets/d/1BO9dr4feFbfD4vAcB-h_ltF9GHKL3pZKhVdtR5O4gEM/edit?usp=sharing';
sheet_tutor = 'Tutor_register';
sheet_student = 'Learner_register';
sheet_email = 'Email_address';
form_student_id = 'https://docs.google.com/forms/d/1Cit5s4LO12IFWOi910oJuY2SNsTayv5xTJsHv15YtMg/edit';
//message = "Test";
ld_email = 'dinhngoc.phuoc@heineken.com';

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

function find_data_lastrow(link, sheetname, col){
  var ss = SpreadsheetApp.openByUrl(link);
  var s = ss.getSheetByName(sheetname);
  var dataRange = s.getDataRange();
  var data = dataRange.getValues();
  return data[data.length-1][col];
//  for (var i = 1; i < data.length ; i++) {
//    if (data[i][3] == ''){;
//      s.getRange(i+1, 4).setValue("Done");
//      return data[i][2];
//    }
//  }  
}

function find_feedback(link, sheetname, col_refer, reference, col_find){
  var ss = SpreadsheetApp.openByUrl(link);
  var s = ss.getSheetByName(sheetname);
  var dataRange = s.getDataRange();
  var data = dataRange.getValues();

  for (var i = 1; i < data.length ; i++) {
   if (data[i][col_refer] == reference){;
//     s.getRange(i, 7).setValue(data[i][col_refer]);
//     s.getRange(i, 8).setValue(data[i][col_find]);
     return data[i][col_find];
   }
 }  
}

function test()
{
  var student_booked = find_data_lastrow(link_Data_base,sheet_student,1);
  var tutor_booked = find_data_lastrow(link_Data_base,sheet_student,2);
  find_feedback(link_Data_base,sheet_email,1,2,student_booked);
  find_feedback(link_Data_base,sheet_email,1,2,tutor_booked);
}

//book tutor
function booking(){  
  var student_booked = find_data_lastrow(link_Data_base,sheet_student,1);//Student name
  var tutor_booked_infor = find_data_lastrow(link_Data_base,sheet_student,2);//Merge Tutor name + date + time
  var ss = SpreadsheetApp.openByUrl(link_Data_base);
  var s = ss.getSheetByName(sheet_tutor);
  var dataRange = s.getDataRange();
  var data = dataRange.getValues();

  var student_email = find_feedback(link_Data_base,sheet_email,1,student_booked,2);

  for (var i = 1; i < data.length ; i++) {

    //chua co nguoi book
    if (data[i][6] == tutor_booked_infor & data[i][5] == ""){
      //send mail here
      s.getRange(i+1, 6).setValue("Booked by "+student_booked);

      var subject = "Booking Success!";
      var tutor_booked = data[i][1];//Tutor name
      var tutor_email = find_feedback(link_Data_base,sheet_email,1,tutor_booked,2);//Tutor email
      var to_email = tutor_email + ',' + student_email;
      var cc_email = ld_email;
      var message = 'Dear Mr/Ms. ' + tutor_booked + ',\n\n' +
                    'Please be informed that '+ student_booked + ' has booked your calendar as schedule below:' + '\n' +
                    tutor_booked_infor + '\n' + '\n' +

                    'Please arrange your time to be OTIF.' + '\n' + '\n' + 
                    'Regards,' + '\n' + 'HVBVT English Club';
      MailApp.sendEmail(to_email, subject, message, {
                            cc: cc_email});//gui mail cho hoc vien
      s.getRange(i+1, 8).setValue("Sent mail");
    }

    //da co nguoi book roi
    else if (data[i][6] == tutor_booked_infor & data[i][5] != "")
    { 
      var subject = "Booking Fail!";
      var tutor_booked = data[i][1];//Tutor name
      var to_email = student_email;
      var cc_email = ld_email;
      var message = 'Dear Mr/Ms. ' + student_booked + ',\n\n' +
                    'You can not book Mr/Ms. '+ tutor_booked + ' because he/she had been booked by another one.' + '\n' +
                    'Please try again with another Tutor.' + '\n' + '\n' +
 
                    'Regards,' + '\n' + 'HVBVT English Club';
      MailApp.sendEmail(to_email, subject, message, {
                            cc: cc_email});//gui mail cho hoc vien
      s.getRange(i+1, 8).setValue("Sent mail");
      break;
    } 
      //SpreadsheetApp.getActiveSheet().getSheetName(sheet_student).getRange(i+1, 4).setValue("Done");
    
  }
  update_form();
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
    //not blank + booked -> 0
    if (data[i][6].search("Booked by") != (-1)) {continue;}
    //not blank + not book -> 0 
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