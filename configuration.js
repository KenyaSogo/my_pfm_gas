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

let createTargetSheetConfigs;
function getCreateTargetSheetConfigs(){
  if(createTargetSheetConfigs) return createTargetSheetConfigs;
  createTargetSheetConfigs = fetchDetailsFromColumns(
    ["create_target_sheet_prefix", "needs_next_month"],
    (colValueArrays, i) => {
      return {
        createTargetSheetPrefix: colValueArrays[0][i],
        needsNextMonth:          colValueArrays[1][i],
      };
    }
  );
  return createTargetSheetConfigs;
}

let assetConfigs;
function getAssetConfigs(){
  if(assetConfigs) return assetConfigs;
  assetConfigs = fetchDetailsFromColumns(
    ["asset_name", "is_cash"],
    (colValueArrays, i) => {
      return {
        assetName: colValueArrays[0][i],
        isCash:    colValueArrays[1][i],
      };
    }
  );
  return assetConfigs;
}

let endBalances;
function getEndBalances(){
  if(endBalances) return endBalances;
  endBalances = fetchDetailsFromColumns(
    ["calc_year", "calc_month", "cash_end_balance", "asset_total_end_balance"],
    (colValueArrays, i) => {
      return {
        calcYear:             colValueArrays[0][i],
        calcMonth:            colValueArrays[1][i],
        cashEndBalance:       colValueArrays[2][i],
        assetTotalEndBalance: colValueArrays[3][i],
      };
    }
  );
  return endBalances;
}

let orderedCategoryConfigs;
function getOrderedCategoryConfigs(){
  if(orderedCategoryConfigs) return orderedCategoryConfigs;
  orderedCategoryConfigs = fetchDetailsFromColumns(
    ["ordered_large_category", "ordered_middle_category", "category_large_class", "category_middle_class"],
    (colValueArrays, i) => {
      return {
        largeCategoryName:       colValueArrays[0][i],
        middleCategoryName:      colValueArrays[1][i],
        categoryLargeClassName:  colValueArrays[2][i],
        categoryMiddleClassName: colValueArrays[3][i],
      };
    }
  );
  return orderedCategoryConfigs;
}
