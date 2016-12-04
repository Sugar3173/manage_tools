/***********************************************************************
* @file insertItems.gs
*
* @brief function of insert Items
*
* @author Shohei.Sugano
*
* @date 2016.12.04
*
* @copyright (c) 2016 aba Co.,Ltd.
*
***********************************************************************/

/**
 * Insert items
 * 
 * @param sheet[in] active sheet
 */
function insertItems(sheet) {
  var items = new Array("No．", "機能", "予定工数", "status", "進捗[%]", "担当", "項目");
  var itemWidth = new Array(40, 70, 70, 70, 70, 50, 200);
  var itemsSize = items.length;
  for(var n = 0; n < itemsSize; ++n){
    sheet.getRange(CONST_ROW_HEADER,n+1).setValue(items[n]).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.setColumnWidth(n+1, itemWidth[n]);
  }
  var maxCols = sheet.getMaxColumns();
  sheet.getRange(CONST_ROW_HEADER,1,1,maxCols).setBackgroundColor(COLOR_BORDER).setBorder(true, true, true, true, true, true);
  sheet.deleteRows(40, 500);
  sheet.deleteColumns(itemsSize+1,maxCols-(itemsSize+1));
  
  sheet.getRange(4,itemsSize).setValue("工数計").setHorizontalAlignment("center").setFontWeight("bold").setBackgroundColor(COLOR_AUTO);
  sheet.getRange(5,itemsSize).setValue("イベント1").setHorizontalAlignment("center").setFontWeight("bold");
  sheet.getRange(6,itemsSize).setValue("イベント2").setHorizontalAlignment("center").setFontWeight("bold");
}
