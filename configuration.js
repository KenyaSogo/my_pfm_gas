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

let slackReportingBotWebhookUrl;
function getSlackReportingBotWebhookUrl(){
  if(slackReportingBotWebhookUrl) return slackReportingBotWebhookUrl;
  slackReportingBotWebhookUrl = getSettingSheet().getRange("slack_reporting_bot_webhook_url").getValue();
  return slackReportingBotWebhookUrl;
}

let slackLoggingBotWebhookUrl;
function getSlackLoggingBotWebhookUrl(){
  if(slackLoggingBotWebhookUrl) return slackLoggingBotWebhookUrl;
  slackLoggingBotWebhookUrl = getSettingSheet().getRange("slack_logging_bot_webhook_url").getValue();
  return slackLoggingBotWebhookUrl;
}

let slackAlertBotWebhookUrl;
function getSlackAlertBotWebhookUrl(){
  if(slackAlertBotWebhookUrl) return slackAlertBotWebhookUrl;
  slackAlertBotWebhookUrl = getSettingSheet().getRange("slack_alert_bot_webhook_url").getValue();
  return slackAlertBotWebhookUrl;
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

let calcMcImportSheetPrefix;
function getCalcMcImportSheetPrefix(){
  if(calcMcImportSheetPrefix) return calcMcImportSheetPrefix;
  calcMcImportSheetPrefix = getSettingSheet().getRange("calc_mc_import_sheet_prefix").getValue();
  return calcMcImportSheetPrefix;
}

let calcMcImportAddr;
function getCalcMcImportAddr(){
  if(calcMcImportAddr) return calcMcImportAddr;
  calcMcImportAddr = getSettingSheet().getRange("calc_mc_import_addr").getValue();
  return calcMcImportAddr;
}

let calcMcExportSheetPrefix;
function getCalcMcExportSheetPrefix(){
  if(calcMcExportSheetPrefix) return calcMcExportSheetPrefix;
  calcMcExportSheetPrefix = getSettingSheet().getRange("calc_mc_export_sheet_prefix").getValue();
  return calcMcExportSheetPrefix;
}

let calcMcExportAddr;
function getCalcMcExportAddr(){
  if(calcMcExportAddr) return calcMcExportAddr;
  calcMcExportAddr = getSettingSheet().getRange("calc_mc_export_addr").getValue();
  return calcMcExportAddr;
}

let calcDcbExportSheetPrefix;
function getCalcDcbExportSheetPrefix(){
  if(calcDcbExportSheetPrefix) return calcDcbExportSheetPrefix;
  calcDcbExportSheetPrefix = getSettingSheet().getRange("calc_dcb_export_sheet_prefix").getValue();
  return calcDcbExportSheetPrefix;
}

let calcDcbExportAddr;
function getCalcDcbExportAddr(){
  if(calcDcbExportAddr) return calcDcbExportAddr;
  calcDcbExportAddr = getSettingSheet().getRange("calc_dcb_export_addr").getValue();
  return calcDcbExportAddr;
}

let calcDabExportSheetPrefix;
function getCalcDabExportSheetPrefix(){
  if(calcDabExportSheetPrefix) return calcDabExportSheetPrefix;
  calcDabExportSheetPrefix = getSettingSheet().getRange("calc_dab_export_sheet_prefix").getValue();
  return calcDabExportSheetPrefix;
}

let calcDabExportAddr;
function getCalcDabExportAddr(){
  if(calcDabExportAddr) return calcDabExportAddr;
  calcDabExportAddr = getSettingSheet().getRange("calc_dab_export_addr").getValue();
  return calcDabExportAddr;
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
