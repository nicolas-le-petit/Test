/*======================================================================
Define to manage HVBQN Library
The function include:
Booking

@ Author: DungTT
@ Date  : 16 Oct 19
@ Version: 1.0

======================================================================*/

/*======================================================================
Including define variable
======================================================================*/

const link_ss = 'https://docs.google.com/spreadsheets/d/1sRPwYmoWTkUu0cvy59EKZMjw1MKuPU1NxEXKYZjXOZo/edit?usp=sharing';
const sheet_register = 'Register';
const sheet_booklist = 'BookList';
const date_off_set = 30*24*60*60;//second

/*======================================================================
Including utilities Function
======================================================================*/

/**
 * Version: 1.0
 * Author : DungTT	
 * Get number of column by header
 * @param : link to GGsheet
 			sheet name
 			header of col
 * @return: number of column
 			-1 if can not find
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
  return -1;
}

/**
 * Version: 1.0
 * Author : DungTT	
 * Get all data of column by header
 * @param : link to GGsheet
 			sheet name
 			header of col
 * @return: array of col data
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

/**
 * Version: 1.0
 * Author : DungTT	
 * Like VlookUp in Exel
 * @param : link to GGsheet
 			sheet name source
 			data to vlook_up
 			sheet name destination
 			number of column refer
 			number of column finding
 			
 * @return: data need to find
 			N/A if can not find
 */

function vlook_up(link, sheetname_src, reference, sheetname_des, col_refer, col_find){
  var ss = SpreadsheetApp.openByUrl(link);

  var s_src = ss.getSheetByName(sheetname_src);
  var dataRange_src = s_src.getDataRange();
  var data_src = dataRange_src.getValues();

  var s_des = ss.getSheetByName(sheetname_des);
  var dataRange_des = s_des.getDataRange();
  var data_des = dataRange_des.getValues();

  for (var i = 1; i < data_des.length ; i++) {
   if (reference == data[i][col_refer]){
     return data[i][col_find];
   }
 }  
 return "N/A";
}

/*======================================================================
Including main Function
======================================================================*/

/**
 * Version: 1.0
 * Author : DungTT	
 * Check the book is available on list
 * @param : link to GGsheet
 			sheet name source
 			data to vlook_up
 			sheet name destination
 			number of column refer
 			number of column finding
 			
 * @return: data need to find
 */

 //open Register
 //get id
 //get timestamp
 //open Booklist
 //vlookup -> row -> yes:no? book tittle/status:add new row
 //send mail
function check_book(){
  var ss_src = SpreadsheetApp.openByUrl(link_ss);
  var s_src  = ss_src.getSheetByName(sheet_register);
  var dataRange_src = s_src.getDataRange();
  var data_src = dataRange_src.getValues();

  var book_id = 
}