/***********************************************************************
* @file createNewSheet.gs
*
* @brief function of Create new Sheet
*
* @author Shohei.Sugano
*
* @date 2016.12.04
*
* @copyright (c) 2016 aba Co.,Ltd.
*
***********************************************************************/

/**
 * Create new Sheet
 */
function createNewSheet(){
//  var start = new Date();//for 処理速度Check
  var app = UiApp.createApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name;
  var i=1;
  while(true){
    name = "sheet"+i;
    var sheet = ss.getSheetByName(name);
    if(!sheet){
      sheet = ss.insertSheet(name, i);
      break;
    }
    ++i;
  }
  sheet.activate();
  

  /* insert WBS tool */
  insertItems(sheet);
  
  /* insert Calendar */
  createCalendar( sheet );
  
  /* set formula */
//  Browser.msgBox(sheet.getLastRow(), Browser.Buttons.OK);
  setFormula(sheet);
  
//  var end = new Date();//for 処理速度Check
//  var span = end - start;
//  Logger.log('startTime: ' + start);
//  Logger.log('endTime: ' + end);
//  Logger.log('処理時間：' + span + 'ミリ秒'); 
}
    
  
