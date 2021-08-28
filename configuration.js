/** @OnlyCurrentDoc */

let thisSpreadSheet;
function getThisSpreadSheet(){
  if(thisSpreadSheet) return thisSpreadSheet;
  thisSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  return thisSpreadSheet;
}

let settingSheet;
function getSettingSheet(){
  if(settingSheet) return settingSheet;
  settingSheet = getThisSpreadSheet().getSheetByName("settings");
  return settingSheet;
}

let calcDcExportSheetPrefix;
function getCalcDcExportSheetPrefix(){
  if(calcDcExportSheetPrefix) return calcDcExportSheetPrefix;
  calcDcExportSheetPrefix = getSettingSheet().getRange("calc_dc_export_sheet_prefix").getValue();
  return calcDcExportSheetPrefix;
}

let calcDcExportAddr;
function getCalcDcExportAddr(){
  if(calcDcExportAddr) return calcDcExportAddr;
  calcDcExportAddr = getSettingSheet().getRange("calc_dc_export_addr").getValue();
  return calcDcExportAddr;
}
