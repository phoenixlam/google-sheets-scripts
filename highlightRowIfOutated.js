/*
 * Aim: highlight the row if (now - lastupdate) > day diff
 *			has 1 bug, add 1 column after running
 *			should write "Last Run: ..." after running
 * Version: 1.0
 * Reference: http://stackoverflow.com/questions/3703676/google-spreadsheet-script-to-change-row-color-when-a-cell-changes-text
 */
function highlightRowIfOutdated() {
  // constant
  var lastupdateColumnIndex = 3; // column of 最後更新
  var highlightColumnOffset = 2; // highlight from which column  
  var categoryColumnIndex = 1;
  
  var range = SpreadsheetApp.getActiveSheet().getDataRange();    
  var today = new Date();
  var requiredDayDiff = -1; // skip for -1
  
  for (var i = range.getRow(); i < range.getLastRow(); i++) {
    rowRange = range.offset(i, 0, 1);
        
    // set requiredDayDiff
    category = rowRange.offset(0,categoryColumnIndex,1,1).getValue();    
    if(rowRange.offset(0,categoryColumnIndex,1,1).isBlank()) {
      
    }else if(category == '每日') {
      requiredDayDiff = 1;
    }else if(category == '每週') {
      requiredDayDiff = 7;
    }else if(category == '每月') {
      requiredDayDiff = 30;
    }else if(category == '每年') {
      requiredDayDiff = 365;
    }else if(category == '其他'){
      break;
    }
    
    if(requiredDayDiff >= 0) {          
     lastUpdateStr = rowRange.offset(0,lastupdateColumnIndex,1,1).getValue();    
     lastUpdateDate = new Date(lastUpdateStr);    
     dayDiff = (today.getTime() - lastUpdateDate.getTime()) / (24*60*60*1000);
     if(isNaN(dayDiff)) {
       //rowRange.offset(0, 7, 1,1).setValue('skip');
     }else if(dayDiff > requiredDayDiff) {
       rowRange.offset(0, highlightColumnOffset).setFontColor("#FF0000");
       //rowRange.offset(0, 7, 1,1).setValue('Y');
     }else {
       rowRange.offset(0, highlightColumnOffset).setFontColor("#000000");
       //rowRange.offset(0, 7, 1,1).setValue('N');
     }    
    }    
  }
}
