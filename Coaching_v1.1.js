/*
@Author: DungTT
@Date  : 8 Mar 2019
@brief : to digitalize coaching
@Name  : Coaching Online Form
*/
link = 'https://docs.google.com/spreadsheets/d/1kToXp0gAVSAsUgJVeLufA-EQudmbgOAbANVQ-kt4k_8/edit?usp=sharing'; 

Module = ["Data submit detail"];

col_time       = 0;
col_coacher    = 1;
col_coachee_1st= 2;
col_coachee_end= 6;
col_content    = 7;
col_action     = 8;
col_deadline   = 9;

offset=3;
/* Create time-driven triggers based on Gmail send schedule */
/*
@Date  : 8 Mar 2019
@brief : auto send email by content
@parameter  : none
*/
function send_email() {
    
  for (var j = 0; j < Module.length; j++){
  
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Module[j]);
    var data = sheet.getDataRange().getValues();
    var time = new Date().getTime();
  
    var last_row = data.length;
    var last_col = sheet.getLastColumn();
    
    for (var i = last_row-1; i>=0; i--){
      if (data[i][last_col-1] != ""){      
        break;
      }
    
    /*for (var i = 0; i<last_row ; i++){
      if (data[i][18] == "Yes"){      
        return;
      }*/
      
      else{
        var schedule = data[i][col_time];
        //Browser.msgBox(schedule.getTime());
        
        if ((schedule != "") && (schedule <= time)){
          var coacher = data[i][col_coacher];//cho nay quan trong - lay chi so
          //tach chuoi -> list coachee
          for (var j = col_coachee_1st; j <= col_coachee_end; j++) {
            if (data[i][j]) 
              {var coachee_list = data[i][j].split(', ');        
                break;}
            }
          }
          
          var subject = 'POD Pillar | Coaching';
          var content = data[i][col_content];
          var message1 = coacher + " vừa thực hiện coaching tại site vào lúc:";
          var message2 = schedule;
          var message3 = "- Người được coach: " + coachee_list;
          var message4 = '- Nội dung coach: ' + content;
          var signature = "Best regards," + '\n' + "POD Pillar";
          
          //var message = 'Dear' + ',\n\n' + message1 + '\n' + message2 + '\n\n' + message3 + '\n' + message4 + '\n\n' + signature;
          
          sheet.getRange(i+1,last_col).setValue("Done");
          //sheet.getRange(i+1,1).setValue(Number(coacher_id) + brewery_id);
         
         /*===================================================== Total Report ================================================*/
          var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Total Report');
          var data2 = sheet2.getDataRange().getValues();
          //var hyperlink = sheet2.getRange("G1").getValue();
          //check date if in the month
          var date = new Date();
          var month = date.getMonth();
          for (var coldata=0;coldata<12;coldata++){
            if (coldata==month)
            {
              break;
            }           
          }
            
          //check coacher id
          for (var rowdata = 0; rowdata<data2.length; rowdata++){
            if (data2[rowdata][1] == coacher){
                  var message5 = "- Số lần đã coach trong tháng " + (coldata+1) + ": " + data2[rowdata][coldata+offset] + " lần";
                  break;
                }
          } 
          
          var message6 = "Vui lòng truy cập đường link bên dưới để xem thông tin chi tiết:" + '\n' + link;
          var message = 'Dear' + ',\n\n' + message1 + '\n' + message2 + '\n\n' + message3 + '\n' 
                        + message4 + '\n' + message5 + '\n' + message6 + '\n\n' + signature ;                        

         /*===================================================== Coachee list ================================================*/
          var sheet4 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Coachee list');
          var data4 = sheet4.getDataRange().getValues();
          var email4 = '';
          
          //check coachee
          for (var pax = 0; pax<coachee_list.length; pax++)
          {
            for (var rowdata = 0; rowdata<data4.length; rowdata++){
              if (data4[rowdata][1] == coachee_list[pax]){
                  email4 += data4[rowdata][3]+',';//coachee
                  //var message5 = "- Số lần đã coach trong tháng " + (coldata+1) + ": " + data2[rowdata][coldata+offset];
                  break;
                }
            }
          }
  
         /*==========================Feedback to 2nd Person==========================*/
          var sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Route');
          var data3 = sheet3.getDataRange().getValues();
          //check coacher name
          for (var rowdata = 0; rowdata<data3.length; rowdata++){
            if (data3[rowdata][1] == coacher){
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
      }
    }
  
}

function test()
{
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Follow up");
  var data2 = sheet2.getDataRange().getValues();
//  
//  var date = new Date();
//  var date1 = new Date('4/1/2019');
//  var date1_m = date1.getMonth();
//  Browser.msgBox("today is: "+ date);
//  Browser.msgBox("date1 is: "+date1);
//  Browser.msgBox("thang cua date1: "+date1_m);
//  
//  if (date.valueOf() < date1.valueOf())
//    
//    Browser.msgBox("today is early than date1");
//  else if (date >= date1)
//    Browser.msgBox("today is later than date1");
  var date = new Date();
  var month = date.getMonth();
          for (var k=0;k<12;k++){
            if (k==month)
            {
              Browser.msgBox("k= "+k);
              Browser.msgBox("today is in: "+month);
              break;
            }
            Browser.msgBox("here");
          }
  
}

function find_col_by_tittle(sheet,tittle){
  var sheet = SpreadsheetApp.getActiveSpreadsheet.getSheetByName(sheet);
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  var last_col = dataRange.getLastColumn();
  for (var i=0; i<=last_col; i++){
    if (data[0][i] == tittle) {
      Browser.msgBox("column no: "+i);
      return i;
    }
  }
}

/*
@Date  : 30 Jul 2019
@brief : to tracking deadline of tasks & remind
@parameter  : none
*/
function remind_task(offset_time){
  var sheet = SpreadsheetApp.getActiveSpreadsheet.getSheetByName("Data submit detail");
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  var time = new Date().getTime();

  var last_row = data.length;
  
  //listdown status & find blank
  for (var i =1; i<last_row; i++){
    if (data[i][11] == "Done" || data[i][11] == "Cancel"){      
      continue;
    }
    
    else{
      var schedule = data[i][15];//check lại stt col
      
      if ((schedule != "") && ((schedule + offset_time) <= time)){
        
        var owner = data[i][17];
        var cc_email = data[i][18];
        var subject = "[Training_Plan]";
        var message1 = "Khóa học: " + data[i][2] + ' đã được lên kế hoạch vào lúc:';
        var message2 = data[i][12];
        var message3 = 'Nhắc nhở!';

        var message = 'Dear,' + '\n\n' + message1 + '\n' + message2 + '\n\n' + message3;
  
        MailApp.sendEmail(owner, subject, message,{cc: cc_email});
      }
    }
  }
}

function abc()
{
 var t = MailApp.getRemainingDailyQuota();
 Browser.msgBox(t);
//          var sheet4 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data submit detail');
//          var string = sheet4.getRange(139,3).getValue();
//          Browser.msgBox(string);
//          var string_o = string.split(', ');
//          Browser.msgBox(string_o);
//          Browser.msgBox(string_o.length);
//          Browser.msgBox(string_o[0]);
//          Browser.msgBox(string_o[1]);
//          Browser.msgBox(string_o[2]);
//          Browser.msgBox(string_o[3]);
}