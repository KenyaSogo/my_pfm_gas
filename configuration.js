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

let phantomJsKey;
function getPhantomJsKey(){
  if(phantomJsKey) return phantomJsKey;
  phantomJsKey = getSettingSheet().getRange("phantom_js_key").getValue();
  return phantomJsKey;
}

let rawDataHeadPattern;
function getRawDataHeadPattern(){
  if(rawDataHeadPattern) return rawDataHeadPattern;
  rawDataHeadPattern = getSettingSheet().getRange("raw_data_head_pattern").getValue();
  return rawDataHeadPattern;
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

let aggreImportSheetPrefix;
function getAggreImportSheetPrefix(){
  if(aggreImportSheetPrefix) return aggreImportSheetPrefix;
  aggreImportSheetPrefix = getSettingSheet().getRange("aggre_import_sheet_prefix").getValue();
  return aggreImportSheetPrefix;
}

let aggreImportAddrCol;
function getAggreImportAddrCol(){
  if(aggreImportAddrCol) return aggreImportAddrCol;
  aggreImportAddrCol = getSettingSheet().getRange("aggre_import_addr_col").getValue();
  return aggreImportAddrCol;
}

let calcDcImportSheetPrefix;
function getCalcDcImportSheetPrefix(){
  if(calcDcImportSheetPrefix) return calcDcImportSheetPrefix;
  calcDcImportSheetPrefix = getSettingSheet().getRange("calc_dc_import_sheet_prefix").getValue();
  return calcDcImportSheetPrefix;
}

let calcDcImportAddrBeforeClosingDay;
function getCalcDcImportAddrBeforeClosingDay(){
  if(calcDcImportAddrBeforeClosingDay) return calcDcImportAddrBeforeClosingDay;
  calcDcImportAddrBeforeClosingDay = getSettingSheet().getRange("calc_dc_import_addr_before_closing_day").getValue();
  return calcDcImportAddrBeforeClosingDay;
}

let calcDcImportAddrAfterClosingDay;
function getCalcDcImportAddrAfterClosingDay(){
  if(calcDcImportAddrAfterClosingDay) return calcDcImportAddrAfterClosingDay;
  calcDcImportAddrAfterClosingDay = getSettingSheet().getRange("calc_dc_import_addr_after_closing_day").getValue();
  return calcDcImportAddrAfterClosingDay;
}

let calcDabImportSheetPrefix;
function getCalcDabImportSheetPrefix(){
  if(calcDabImportSheetPrefix) return calcDabImportSheetPrefix;
  calcDabImportSheetPrefix = getSettingSheet().getRange("calc_dab_import_sheet_prefix").getValue();
  return calcDabImportSheetPrefix;
}

let calcDabImportAddrBeforeClosingDay;
function getCalcDabImportAddrBeforeClosingDay(){
  if(calcDabImportAddrBeforeClosingDay) return calcDabImportAddrBeforeClosingDay;
  calcDabImportAddrBeforeClosingDay = getSettingSheet().getRange("calc_dab_import_addr_before_closing_day").getValue();
  return calcDabImportAddrBeforeClosingDay;
}

let calcDabImportAddrAfterClosingDay;
function getCalcDabImportAddrAfterClosingDay(){
  if(calcDabImportAddrAfterClosingDay) return calcDabImportAddrAfterClosingDay;
  calcDabImportAddrAfterClosingDay = getSettingSheet().getRange("calc_dab_import_addr_after_closing_day").getValue();
  return calcDabImportAddrAfterClosingDay;
}

let calcDcbImportSheetPrefix;
function getCalcDcbImportSheetPrefix(){
  if(calcDcbImportSheetPrefix) return calcDcbImportSheetPrefix;
  calcDcbImportSheetPrefix = getSettingSheet().getRange("calc_dcb_import_sheet_prefix").getValue();
  return calcDcbImportSheetPrefix;
}

let calcDcbImportAddrBeforeClosingDay;
function getCalcDcbImportAddrBeforeClosingDay(){
  if(calcDcbImportAddrBeforeClosingDay) return calcDcbImportAddrBeforeClosingDay;
  calcDcbImportAddrBeforeClosingDay = getSettingSheet().getRange("calc_dcb_import_addr_before_closing_day").getValue();
  return calcDcbImportAddrBeforeClosingDay;
}

let calcDcbImportAddrAfterClosingDay;
function getCalcDcbImportAddrAfterClosingDay(){
  if(calcDcbImportAddrAfterClosingDay) return calcDcbImportAddrAfterClosingDay;
  calcDcbImportAddrAfterClosingDay = getSettingSheet().getRange("calc_dcb_import_addr_after_closing_day").getValue();
  return calcDcbImportAddrAfterClosingDay;
}

let aggreExportSheetPrefix;
function getAggreExportSheetPrefix(){
  if(aggreExportSheetPrefix) return aggreExportSheetPrefix;
  aggreExportSheetPrefix = getSettingSheet().getRange("aggre_export_sheet_prefix").getValue();
  return aggreExportSheetPrefix;
}

let aggreExportAddr;
function getAggreExportAddr(){
  if(aggreExportAddr) return aggreExportAddr;
  aggreExportAddr = getSettingSheet().getRange("aggre_export_addr").getValue();
  return aggreExportAddr;
}

let templateSheetSuffix;
function getTemplateSheetSuffix(){
  if(templateSheetSuffix) return templateSheetSuffix;
  templateSheetSuffix = getSettingSheet().getRange("template_sheet_suffix").getValue();
  return templateSheetSuffix;
}

let todayForCalenderRange;
function getTodayForCalenderRange(){
  if(todayForCalenderRange) return todayForCalenderRange;
  todayForCalenderRange = getSettingSheet().getRange("today_for_calender");
  return todayForCalenderRange;
}

let isCalenderNeedsMaintenance;
function getIsCalenderNeedsMaintenance(){
  if(isCalenderNeedsMaintenance) return isCalenderNeedsMaintenance;
  isCalenderNeedsMaintenance = getSettingSheet().getRange("is_calender_needs_maintenance").getValue();
  return isCalenderNeedsMaintenance;
}

let isMonthlyStartDayInvalid;
function getIsMonthlyStartDayInvalid(){
  if(isMonthlyStartDayInvalid) return isMonthlyStartDayInvalid;
  isMonthlyStartDayInvalid = getSettingSheet().getRange("is_monthly_start_day_invalid").getValue();
  return isMonthlyStartDayInvalid;
}

let isTomorrowMonthlyStart;
function getIsTomorrowMonthlyStart(){
  if(isTomorrowMonthlyStart) return isTomorrowMonthlyStart;
  isTomorrowMonthlyStart = getSettingSheet().getRange("is_tomorrow_monthly_start").getValue();
  return isTomorrowMonthlyStart;
}

let assetTotalEndBalanceRange;
function getAssetTotalEndBalanceRange(){
  if(assetTotalEndBalanceRange) return assetTotalEndBalanceRange;
  assetTotalEndBalanceRange = getSettingSheet().getRange("asset_total_end_balance");
  return assetTotalEndBalanceRange;
}

let cashEndBalanceRange;
function getCashEndBalanceRange(){
  if(cashEndBalanceRange) return cashEndBalanceRange;
  cashEndBalanceRange = getSettingSheet().getRange("cash_end_balance");
  return cashEndBalanceRange;
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
