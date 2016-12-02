function createNewSheet(){
  // var start = new Date();//for 処理速度Check
  var app = UiApp.createApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name;
  var i=1;
  ///* 
  //temporary commentout
  while(true){
    name = "sheet"+i;
    var sheet = ss.getSheetByName(name);
    if(!sheet){
      sheet = ss.insertSheet(name, i);
      break;
    }
    ++i;
  }
  //*/
  //var sheet = ss.getSheetByName("Sheet2");//for debug
  sheet.activate();
  

  /* insert WBS tool */
 var headerRow = 16;
 var items = new Array("No．", "機能", "予定工数", "status", "進捗", "担当", "項目");
 var itemWidth = new Array(40, 70, 70, 70, 70, 50, 200);
 var itemsSize = items.length;
 for(var n = 0; n < itemsSize; ++n){
   sheet.getRange(headerRow,n+1).setValue(items[n]).setHorizontalAlignment("center");
   sheet.setColumnWidth(n+1, itemWidth[n]);
 }
 sheet.getRange(headerRow,1,1,7).setBackgroundColor(COLOR_BORDER).setBorder(true, true, true, true, true, true);
 sheet.deleteRows(40, 500);
 sheet.deleteColumns(46,50);
  
  /* insert Calendar */
  createCalendar( sheet );
  
  // var end = new Date();//for 処理速度Check
  // var span = end - start;
  // Logger.log('startTime: ' + start);
  // Logger.log('endTime: ' + end);
  // Logger.log('処理時間：' + span + 'ミリ秒'); 
}