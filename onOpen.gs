/***********************************************************************
* @file onOpen.gs
*
* @brief function of onOpen event
*
* @author Shohei.Sugano
*
* @date 2016.12.04
*
* @copyright (c) 2016 aba Co.,Ltd.
*
***********************************************************************/

/* reference */
/*
url1: eye4brain.sakura.ne.jp/wp/blog/category/gascript/
*/

/* 各セルの色(#RGB) */
var COLOR_TARGET_RED ="#F00"
var COLOR_HOLIDAY = "#F7BE81";
var COLOR_BORDER = "#0CF";
var COLOR_TENCHAN = "#05F";
var COLOR_DAY = "#FFF";
var COLOR_TODAY_BACKGROUND = "#333";
var COLOR_TODAY_TEXT = "#FFF";
var COLOR_DATE_BACKGROUND = "#DDD";
var COLOR_DATE_TEXT = "#000";

var COLOR_AUTO = "#c9daf8";   //自動表示・計算
var COLOR_MANUAL = "#fff2cc"; //手動入力
var COLOR_SELECT = "#d9ead3"; //選択入力

/* 共通定数 */
var CONST_CONFIG_SHEET_NAME = "config";
var CONST_CONFIG_ROW_START = 4;
var CONST_CONFIG_ROW_END = 9;

var CONST_ROW_SET_MAX = 100;
var CONST_ROW_HEADER = 7;   //16;

//Browser.msgBox(yourSelections);

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // メニューバーに追加
  var menuEntries = [{name: "新規ガントチャートの挿入", functionName: "createNewSheet"},
                     {name: "日付の追加(Not supported)", functionName: "appendChart"}, 
                     null,
                     {name: "バージョン", functionName: "getVersion"}];
  ss.addMenu("管理ツール", menuEntries);
}
