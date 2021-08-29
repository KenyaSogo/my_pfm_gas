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

let orderedCategoryConfigs;
function getOrderedCategoryConfigs(){
  if(orderedCategoryConfigs) return orderedCategoryConfigs;
  const orderedLargeCategories = getTrimmedColumnValues(getSettingSheet(), "ordered_large_category");
  const orderedMiddleCategories = getTrimmedColumnValues(getSettingSheet(), "ordered_middle_category");
  const categoryLargeClasses = getTrimmedColumnValues(getSettingSheet(), "category_large_class");
  const categoryMiddleClasses = getTrimmedColumnValues(getSettingSheet(), "category_middle_class");
  orderedCategoryConfigs = getIntRangeFromZero(orderedLargeCategories.length).map(
    i => {
      return {
        largeCategoryName:       orderedLargeCategories[i],
        middleCategoryName:      orderedMiddleCategories[i],
        categoryLargeClassName:  categoryLargeClasses[i],
        categoryMiddleClassName: categoryMiddleClasses[i],
      }
    }
  );
  return orderedCategoryConfigs;
}
