/***********************************************************************
* @file setFormula.gs
*
* @brief function of setting formula
*
* @author Shohei.Sugano
*
* @date 2016.12.04
*
* @copyright (c) 2016 aba Co.,Ltd.
*
***********************************************************************/

/**
 * Setting formula
 * 
 * @param sheet[in] active sheet
 */
function setFormula(sheet) {
  
  var refsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONST_CONFIG_SHEET_NAME);
  var refRangeList = new Array(
    "A"+CONST_CONFIG_ROW_START+":A"+CONST_CONFIG_ROW_END, 
    "B"+CONST_CONFIG_ROW_START+":B"+CONST_CONFIG_ROW_END,
    "C"+CONST_CONFIG_ROW_START+":C"+CONST_CONFIG_ROW_END
  );
  var setSelectRangeList = new Array(
    "B"+(CONST_ROW_HEADER+1)+":B"+CONST_ROW_SET_MAX, 
    "D"+(CONST_ROW_HEADER+1)+":D"+CONST_ROW_SET_MAX,
    "F"+(CONST_ROW_HEADER+1)+":F"+CONST_ROW_SET_MAX
  );
  var setManualRangeList = new Array(
    "C"+(CONST_ROW_HEADER+1)+":C"+CONST_ROW_SET_MAX, 
    "E"+(CONST_ROW_HEADER+1)+":E"+CONST_ROW_SET_MAX,
    "G"+(CONST_ROW_HEADER+1)+":G"+CONST_ROW_SET_MAX
  );
  
  var calendarRange = sheet.getRange("2:4").getValues();
  var maxCols = sheet.getMaxColumns();
  
  //工数計
  for(var i = 0, j = 0; i < maxCols; ++i){
    if(calendarRange[0][i] != ""){
      var totalManHours = "=SUM(R[4]C["+j+"]:R["+(CONST_ROW_SET_MAX-4)+"]C["+j+"])";
      sheet.getRange(4,i+1).setFormulaR1C1(totalManHours).setBackgroundColor(COLOR_AUTO);
      ++j;
    }
  }
  
  //No.(col:A)
  //=IF(G9="","",MAX($A$8)+1)
  var rangeA = sheet.getRange("A"+(CONST_ROW_HEADER+1)+":A"+CONST_ROW_SET_MAX);
  var tmp = rangeA.getValues();
  for(i = 0; i < (CONST_ROW_SET_MAX - CONST_ROW_HEADER); ++i ){
    if(i == 0 ){
      tmp[i][0] = 1;
    }else{
      tmp[i][0] = "=IF(G"+((CONST_ROW_HEADER+1)+i)+"=\"\",\"\",MAX($A$"+(CONST_ROW_HEADER+1)+":A"+(CONST_ROW_HEADER+i)+")+1)";
    }
  }
  rangeA.setFormulas(tmp).setBackgroundColor(COLOR_AUTO);
  

  //プルダウン(col: B, D, F)
  for(var k = 0; k < refRangeList.length; ++k){
    var rng = sheet.getRange(setSelectRangeList[k]);
    var pullList = refsheet.getRange(refRangeList[k]).getValues();
　  var rule = SpreadsheetApp
　　  　.newDataValidation()
　　　  .requireValueInList(pullList, true)
　　　  .build();

　  rng.setDataValidation(rule).setBackgroundColor(COLOR_SELECT);
  }
      
  //set border color
  sheet.getRange(CONST_ROW_HEADER,1,1,maxCols).setBackgroundColor(COLOR_BORDER);
  // set manual range color
  for(k = 0; k < setManualRangeList.length; ++k){
    var tmpRange = sheet.getRange(setManualRangeList[k]);
　  tmpRange.setBackgroundColor(COLOR_MANUAL);
  }
  
  // explanation of color
  sheet.getRange("A4").setValue("自動表示・計算").setBackgroundColor(COLOR_AUTO);
  sheet.getRange("A5").setValue("手動入力").setBackgroundColor(COLOR_MANUAL);
  sheet.getRange("A6").setValue("選択入力").setBackgroundColor(COLOR_SELECT);
  //cell marge
  sheet.getRange("A4:B4").mergeAcross();
  sheet.getRange("A5:B5").mergeAcross();
  sheet.getRange("A6:B6").mergeAcross();  
  
  //title
  sheet.getRange(1,1,2,7).merge();
  sheet.getRange("A1").setValue("Title").setFontSize(18).setHorizontalAlignment("left").setVerticalAlignment("middle");
}
