/***********************************************************************
* @file createCalendar.gs
*
* @brief function of Create calendar
*
* @author Shohei.Sugano
*
* @date 2016.12.04
*
* @copyright (c) 2016 aba Co.,Ltd.
*
***********************************************************************/

/**
 * Create calendar
 * 
 * @param sheet[in] active sheet
 */
function createCalendar( sheet ) {
　var start = new Date();//for 処理速度Check
  var today = new Date();
  var defaultMonthWidth = 2;
  var lastDate = today;
  var startDate = new Date(lastDate.getYear(), lastDate.getMonth(), lastDate.getDate());
  var endDate = new Date(today.getYear(), today.getMonth() + defaultMonthWidth, today.getDate());

  
  var startDateColumn = sheet.getLastColumn() + 1;
  var diffDate = (endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24);
  diffDate = Math.floor(diffDate) - 3;//１日分マイナス+startDateColumnでの+1を相殺
//  Browser.msgBox(diffDate, Browser.Buttons.OK);
  var endDateColumn = startDateColumn + diffDate;
  
  var week = new Array("日", "月", "火", "水", "木", "金", "土");
  var month = new Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec");
  
  // 祝日の日を取得
  var holidayCalendar = CalendarApp.getCalendarsByName("日本の祝日")[0];
  if (holidayCalendar == undefined) {
    holidayCalendar = CalendarApp.getCalendarsByName("Holidays in Japan")[0];
  }
  
  var holidayEvents = holidayCalendar.getEvents(startDate, endDate);
  var holidayDates = new Array(holidayEvents.length);
  for (var i = 0; i < holidayEvents.length; i++) {
    var startTime = holidayEvents[i].getStartTime();
    var holidayDate = new Date(startTime.getYear(), startTime.getMonth(), startTime.getDate());
    holidayDates[i] = holidayDate;
  }
  

  var rowNum = sheet.getMaxRows() - 1;
  if(rowNum > CONST_ROW_SET_MAX){
    rowNum = CONST_ROW_SET_MAX-1;
  }
  sheet.insertColumnsAfter(startDateColumn, diffDate + 1);
  sheet.getRange(2, startDateColumn, rowNum, diffDate + 1).setBorder(true, true, true, true, true, true);
  sheet.getRange(2, startDateColumn, 1, diffDate + 1).setNumberFormat("d");
  
  for (var i = 0, j = 0; i < diffDate + 1; i++) {
    sheet.setColumnWidth(startDateColumn + i, 18);
    var newDate = new Date(startDate.getYear(), startDate.getMonth(), startDate.getDate() + i);
    var day = week[newDate.getDay()];
    if (i==0 || newDate.getDate() == 1) {
      sheet.getRange(1, startDateColumn + i).setValue(month[newDate.getMonth()]+", "+newDate.getYear()).setHorizontalAlignment("left");
    } else {
      // 月のセルを結合
      var mergeRangeWidth = newDate.getDate(); // 結合するセルの幅
      var firstDayColumn = startDateColumn + i - (mergeRangeWidth - 1); // 月の初めのセルの列
      if(firstDayColumn<startDateColumn){
        firstDayColumn = startDateColumn;
        mergeRangeWidth = mergeRangeWidth - startDate.getDate()+1;
      }
      Logger.log(mergeRangeWidth);
      Logger.log(firstDayColumn);
      sheet.getRange(1, firstDayColumn, 1, mergeRangeWidth).mergeAcross();
    }
    if (day == "土" || day == "日") {
      sheet.getRange(2, startDateColumn + i, rowNum, 1).setBackgroundColor(COLOR_HOLIDAY);
    } else {
      sheet.getRange(2, startDateColumn + i, 1, 1).setBackgroundColor(COLOR_DAY);
    }

    if (holidayDates.length != 0) {
      //祝日があれば色づける
      if (newDate.getYear() == holidayDates[j].getYear() && 
          newDate.getMonth() == holidayDates[j].getMonth() && 
          newDate.getDate() == holidayDates[j].getDate()) {
        sheet.getRange(2, startDateColumn + i, rowNum, 1).setBackgroundColor(COLOR_HOLIDAY);
        // 祝日の名前を入力する
        var holidayName = holidayEvents[j].getTitle();
        sheet.getRange(2, startDateColumn + i).setComment(holidayName);
        if (holidayDates.length != j + 1) {
          ++j;
        }
      }
    }
    
    sheet.getRange(2, startDateColumn + i).setFontSize(8);
    sheet.getRange(2, startDateColumn + i).setValue(newDate);
    sheet.getRange(3, startDateColumn + i).setValue(day);
  }

  var end = new Date();//for 処理速度Check
  var span = end - start;
  Logger.log('startTime: ' + start);
  Logger.log('endTime: ' + end);
  Logger.log('処理時間：' + span + 'ミリ秒'); 
}
