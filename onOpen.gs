/* reference */
/*
url1: eye4brain.sakura.ne.jp/wp/blog/category/gascript/
*/

/* 各セルの色(#RGB) */
var COLOR_TARGET_RED ="#F00"
var COLOR_HOLIDAY = "#FA0";
var COLOR_BORDER = "#0CF";
var COLOR_TENCHAN = "#05F";
var COLOR_DAY = "#FFF";
var COLOR_TODAY_BACKGROUND = "#333";
var COLOR_TODAY_TEXT = "#FFF";
var COLOR_DATE_BACKGROUND = "#DDD";
var COLOR_DATE_TEXT = "#000";

/* 共通定数 */


//Browser.msgBox(yourSelections);

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // メニューバーに追加
  var menuEntries = [{name: "新規ガントチャートの挿入", functionName: "createNewSheet"},
                     {name: "日付の追加", functionName: "appendChart"}, 
                     null,
                     {name: "バージョン", functionName: "getVersion"}];
  ss.addMenu("管理ツール", menuEntries);
}
