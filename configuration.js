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
  settingSheet = getThisSpreadSheet().getSheetByName(getIsTestEnv() ? "test_settings" : "settings");
  return settingSheet;
}

let phantomJsKey;
function getPhantomJsKey(){
  if(phantomJsKey) return phantomJsKey;
  phantomJsKey = getRangeValueFrom(getSettingSheet(), "phantom_js_key");
  return phantomJsKey;
}

let rawDataHeadPattern;
function getRawDataHeadPattern(){
  if(rawDataHeadPattern) return rawDataHeadPattern;
  rawDataHeadPattern = getSettingSheet().getRange("raw_data_head_pattern").getValue();
  return rawDataHeadPattern;
}

let rawDataNoDetailValue;
function getRawDataNoDetailValue(){
  if(rawDataNoDetailValue) return rawDataNoDetailValue;
  rawDataNoDetailValue = getSettingSheet().getRange("raw_data_no_detail_value").getValue();
  return rawDataNoDetailValue;
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
  aggreImportSheetPrefix = getRangeValueFrom(getSettingSheet(), "aggre_import_sheet_prefix");
  return aggreImportSheetPrefix;
}

let aggreImportAddrCol;
function getAggreImportAddrCol(){
  if(aggreImportAddrCol) return aggreImportAddrCol;
  aggreImportAddrCol = getRangeValueFrom(getSettingSheet(), "aggre_import_addr_col");
  return aggreImportAddrCol;
}

let calcDcImportSheetPrefix;
function getCalcDcImportSheetPrefix(){
  if(calcDcImportSheetPrefix) return calcDcImportSheetPrefix;
  calcDcImportSheetPrefix = getRangeValueFrom(getSettingSheet(), "calc_dc_import_sheet_prefix");
  return calcDcImportSheetPrefix;
}

let calcDcImportAddrBeforeClosingDay;
function getCalcDcImportAddrBeforeClosingDay(){
  if(calcDcImportAddrBeforeClosingDay) return calcDcImportAddrBeforeClosingDay;
  calcDcImportAddrBeforeClosingDay = getRangeValueFrom(getSettingSheet(), "calc_dc_import_addr_before_closing_day");
  return calcDcImportAddrBeforeClosingDay;
}

let calcDcImportAddrAfterClosingDay;
function getCalcDcImportAddrAfterClosingDay(){
  if(calcDcImportAddrAfterClosingDay) return calcDcImportAddrAfterClosingDay;
  calcDcImportAddrAfterClosingDay = getRangeValueFrom(getSettingSheet(), "calc_dc_import_addr_after_closing_day");
  return calcDcImportAddrAfterClosingDay;
}

let calcDabImportSheetPrefix;
function getCalcDabImportSheetPrefix(){
  if(calcDabImportSheetPrefix) return calcDabImportSheetPrefix;
  calcDabImportSheetPrefix = getRangeValueFrom(getSettingSheet(), "calc_dab_import_sheet_prefix");
  return calcDabImportSheetPrefix;
}

let calcDabImportAddrBeforeClosingDay;
function getCalcDabImportAddrBeforeClosingDay(){
  if(calcDabImportAddrBeforeClosingDay) return calcDabImportAddrBeforeClosingDay;
  calcDabImportAddrBeforeClosingDay = getRangeValueFrom(getSettingSheet(), "calc_dab_import_addr_before_closing_day");
  return calcDabImportAddrBeforeClosingDay;
}

let calcDabImportAddrAfterClosingDay;
function getCalcDabImportAddrAfterClosingDay(){
  if(calcDabImportAddrAfterClosingDay) return calcDabImportAddrAfterClosingDay;
  calcDabImportAddrAfterClosingDay = getRangeValueFrom(getSettingSheet(), "calc_dab_import_addr_after_closing_day");
  return calcDabImportAddrAfterClosingDay;
}

let calcDcbImportSheetPrefix;
function getCalcDcbImportSheetPrefix(){
  if(calcDcbImportSheetPrefix) return calcDcbImportSheetPrefix;
  calcDcbImportSheetPrefix = getRangeValueFrom(getSettingSheet(), "calc_dcb_import_sheet_prefix");
  return calcDcbImportSheetPrefix;
}

let calcDcbImportAddrBeforeClosingDay;
function getCalcDcbImportAddrBeforeClosingDay(){
  if(calcDcbImportAddrBeforeClosingDay) return calcDcbImportAddrBeforeClosingDay;
  calcDcbImportAddrBeforeClosingDay = getRangeValueFrom(getSettingSheet(), "calc_dcb_import_addr_before_closing_day");
  return calcDcbImportAddrBeforeClosingDay;
}

let calcDcbImportAddrAfterClosingDay;
function getCalcDcbImportAddrAfterClosingDay(){
  if(calcDcbImportAddrAfterClosingDay) return calcDcbImportAddrAfterClosingDay;
  calcDcbImportAddrAfterClosingDay = getRangeValueFrom(getSettingSheet(), "calc_dcb_import_addr_after_closing_day");
  return calcDcbImportAddrAfterClosingDay;
}

let aggreExportSheetPrefix;
function getAggreExportSheetPrefix(){
  if(aggreExportSheetPrefix) return aggreExportSheetPrefix;
  aggreExportSheetPrefix = getRangeValueFrom(getSettingSheet(), "aggre_export_sheet_prefix");
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

let aggregateYearRange;
function getAggregateYearRange(){
  if(aggregateYearRange) return aggregateYearRange;
  aggregateYearRange = getSettingSheet().getRange("aggregate_year");
  return aggregateYearRange;
}

let aggregateMonthRange;
function getAggregateMonthRange(){
  if(aggregateMonthRange) return aggregateMonthRange;
  aggregateMonthRange = getSettingSheet().getRange("aggregate_month");
  return aggregateMonthRange;
}

let todayForCalenderRange;
function getTodayForCalenderRange(){
  if(todayForCalenderRange) return todayForCalenderRange;
  todayForCalenderRange = getSettingSheet().getRange("today_for_calender");
  return todayForCalenderRange;
}

let isTodayMonthlyStart;
function getIsTodayMonthlyStart(){
  if(isTodayMonthlyStart) return isTodayMonthlyStart;
  isTodayMonthlyStart = getSettingSheet().getRange("is_today_monthly_start").getValue();
  return isTodayMonthlyStart;
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
  assetTotalEndBalanceRange = getRangeFrom(getSettingSheet(), "asset_total_end_balance");
  return assetTotalEndBalanceRange;
}

let cashEndBalanceRange;
function getCashEndBalanceRange(){
  if(cashEndBalanceRange) return cashEndBalanceRange;
  cashEndBalanceRange = getRangeFrom(getSettingSheet(), "cash_end_balance");
  return cashEndBalanceRange;
}

let pfmAccountConfigs;
function getPfmAccountConfigs(){
  if(pfmAccountConfigs) return pfmAccountConfigs;
  pfmAccountConfigs = fetchSettingDetailsFromColumns(
    getColumnRangeNames(["pfm_id", "pfm_pass"]),
    (colValueArrays, i) => {
      return {
        pfmId:   colValueArrays[0][i],
        pfmPass: colValueArrays[1][i],
      };
    }
  );
  return pfmAccountConfigs;
}

let aggregateYearMonthConfigs;
function getAggregateYearMonthConfigs(){
  if(aggregateYearMonthConfigs) return aggregateYearMonthConfigs;
  aggregateYearMonthConfigs = fetchSettingDetailsFromColumns(
    getColumnRangeNames(["aggregate_year", "aggregate_month"]),
    (colValueArrays, i) => {
      return {
        aggregateYear:  colValueArrays[0][i],
        aggregateMonth: colValueArrays[1][i],
      };
    }
  );
  return aggregateYearMonthConfigs;
}

let createTargetSheetConfigs;
function getCreateTargetSheetConfigs(){
  if(createTargetSheetConfigs) return createTargetSheetConfigs;
  createTargetSheetConfigs = fetchSettingDetailsFromColumns(
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
  assetConfigs = fetchSettingDetailsFromColumns(
    getColumnRangeNames(["asset_name", "is_cash"]),
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
  endBalances = fetchSettingDetailsFromColumns(
    getColumnRangeNames(["calc_year", "calc_month", "cash_end_balance", "asset_total_end_balance"]),
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
  orderedCategoryConfigs = fetchSettingDetailsFromColumns(
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

let testAggreRawDataDetailGetCount = 0;
function getTestAggreRawDataDetail(){
  if(!getIsTestEnv()) return null;
  testAggreRawDataDetailGetCount++;
  const targetRangeName = testAggreRawDataDetailGetCount == 1 ? "test_aggre_raw_data_detail_1" : "test_aggre_raw_data_detail_2";
  return getSettingSheet().getRange(targetRangeName).getValue();
}

let testAggreResultGot1; // TODO: シート関数使わずに result を取得 (シートとロジックの分離 (シート=DBに))
function getTestAggreResultGot1(){
  if(!getIsTestEnv()) return null;
  if(testAggreResultGot1) return testAggreResultGot1;
  testAggreResultGot1 = getSettingSheet().getRange("test_aggre_result_got_1").getValue();
  return testAggreResultGot1;
}

let testAggreResultGot2;
function getTestAggreResultGot2(){
  if(!getIsTestEnv()) return null;
  if(testAggreResultGot2) return testAggreResultGot2;
  testAggreResultGot2 = getSettingSheet().getRange("test_aggre_result_got_2").getValue();
  return testAggreResultGot2;
}

let testAggreResultExpected1;
function getTestAggreResultExpected1(){
  if(!getIsTestEnv()) return null;
  if(testAggreResultExpected1) return testAggreResultExpected1;
  testAggreResultExpected1 = getSettingSheet().getRange("test_aggre_result_expected_1").getValue();
  return testAggreResultExpected1;
}

let testAggreResultExpected2;
function getTestAggreResultExpected2(){
  if(!getIsTestEnv()) return null;
  if(testAggreResultExpected2) return testAggreResultExpected2;
  testAggreResultExpected2 = getSettingSheet().getRange("test_aggre_result_expected_2").getValue();
  return testAggreResultExpected2;
}
