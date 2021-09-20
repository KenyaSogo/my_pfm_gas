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
  calcDcExportSheetPrefix = getRangeValueFrom(getSettingSheet(), "calc_dc_export_sheet_prefix");
  return calcDcExportSheetPrefix;
}

let calcDcExportAddr;
function getCalcDcExportAddr(){
  if(calcDcExportAddr) return calcDcExportAddr;
  calcDcExportAddr = getRangeValueFrom(getSettingSheet(), "calc_dc_export_addr");
  return calcDcExportAddr;
}

let calcMcImportSheetPrefix;
function getCalcMcImportSheetPrefix(){
  if(calcMcImportSheetPrefix) return calcMcImportSheetPrefix;
  calcMcImportSheetPrefix = getRangeValueFrom(getSettingSheet(), "calc_mc_import_sheet_prefix");
  return calcMcImportSheetPrefix;
}

let calcMcImportAddr;
function getCalcMcImportAddr(){
  if(calcMcImportAddr) return calcMcImportAddr;
  calcMcImportAddr = getRangeValueFrom(getSettingSheet(), "calc_mc_import_addr");
  return calcMcImportAddr;
}

let calcMcExportSheetPrefix;
function getCalcMcExportSheetPrefix(){
  if(calcMcExportSheetPrefix) return calcMcExportSheetPrefix;
  calcMcExportSheetPrefix = getRangeValueFrom(getSettingSheet(), "calc_mc_export_sheet_prefix");
  return calcMcExportSheetPrefix;
}

let calcMcExportAddr;
function getCalcMcExportAddr(){
  if(calcMcExportAddr) return calcMcExportAddr;
  calcMcExportAddr = getRangeValueFrom(getSettingSheet(), "calc_mc_export_addr");
  return calcMcExportAddr;
}

let calcDcbExportSheetPrefix;
function getCalcDcbExportSheetPrefix(){
  if(calcDcbExportSheetPrefix) return calcDcbExportSheetPrefix;
  calcDcbExportSheetPrefix = getRangeValueFrom(getSettingSheet(), "calc_dcb_export_sheet_prefix");
  return calcDcbExportSheetPrefix;
}

let calcDcbExportAddr;
function getCalcDcbExportAddr(){
  if(calcDcbExportAddr) return calcDcbExportAddr;
  calcDcbExportAddr = getRangeValueFrom(getSettingSheet(), "calc_dcb_export_addr");
  return calcDcbExportAddr;
}

let calcDabExportSheetPrefix;
function getCalcDabExportSheetPrefix(){
  if(calcDabExportSheetPrefix) return calcDabExportSheetPrefix;
  calcDabExportSheetPrefix = getRangeValueFrom(getSettingSheet(), "calc_dab_export_sheet_prefix");
  return calcDabExportSheetPrefix;
}

let calcDabExportAddr;
function getCalcDabExportAddr(){
  if(calcDabExportAddr) return calcDabExportAddr;
  calcDabExportAddr = getRangeValueFrom(getSettingSheet(), "calc_dab_export_addr");
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
  aggreExportAddr = getRangeValueFrom(getSettingSheet(), "aggre_export_addr");
  return aggreExportAddr;
}

let templateSheetSuffix;
function getTemplateSheetSuffix(){
  if(templateSheetSuffix) return templateSheetSuffix;
  templateSheetSuffix = getSettingSheet().getRange("template_sheet_suffix").getValue();
  return templateSheetSuffix;
}

let graphDcbImportSheetName;
function getGraphDcbImportSheetName(){
  if(graphDcbImportSheetName) return graphDcbImportSheetName;
  graphDcbImportSheetName = getRangeValueFrom(getSettingSheet(), "graph_dcb_import_sheet_name");
  return graphDcbImportSheetName;
}

let graphDcbImportAddr;
function getGraphDcbImportAddr(){
  if(graphDcbImportAddr) return graphDcbImportAddr;
  graphDcbImportAddr = getRangeValueFrom(getSettingSheet(), "graph_dcb_import_addr");
  return graphDcbImportAddr;
}

let graphDcbImageUrl;
function getGraphDcbImageUrl(){
  if(graphDcbImageUrl) return graphDcbImageUrl;
  graphDcbImageUrl = getSettingSheet().getRange("graph_dcb_image_url").getValue();
  return graphDcbImageUrl;
}

let graphDabImportSheetName;
function getGraphDabImportSheetName(){
  if(graphDabImportSheetName) return graphDabImportSheetName;
  graphDabImportSheetName = getRangeValueFrom(getSettingSheet(), "graph_dab_import_sheet_name");
  return graphDabImportSheetName;
}

let graphDabImportAddr;
function getGraphDabImportAddr(){
  if(graphDabImportAddr) return graphDabImportAddr;
  graphDabImportAddr = getRangeValueFrom(getSettingSheet(), "graph_dab_import_addr");
  return graphDabImportAddr;
}

let graphDabImageUrl;
function getGraphDabImageUrl(){
  if(graphDabImageUrl) return graphDabImageUrl;
  graphDabImageUrl = getSettingSheet().getRange("graph_dab_image_url").getValue();
  return graphDabImageUrl;
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
    getColumnRangeNames(["ordered_large_category", "ordered_middle_category", "category_large_class", "category_middle_class"]),
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

let testCalenderDateYear;
function getTestCalenderDateYear(){
  if(!getIsTestEnv()) return null;
  if(testCalenderDateYear) return testCalenderDateYear;
  testCalenderDateYear = getSettingSheet().getRange("test_calender_date_year").getValue();
  return testCalenderDateYear;
}

let testCalenderDateMonth;
function getTestCalenderDateMonth(){
  if(!getIsTestEnv()) return null;
  if(testCalenderDateMonth) return testCalenderDateMonth;
  testCalenderDateMonth = getSettingSheet().getRange("test_calender_date_month").getValue();
  return testCalenderDateMonth;
}

let testCalenderDateDay;
function getTestCalenderDateDay(){
  if(!getIsTestEnv()) return null;
  if(testCalenderDateDay) return testCalenderDateDay;
  testCalenderDateDay = getSettingSheet().getRange("test_calender_date_day").getValue();
  return testCalenderDateDay;
}

let testCalenderDate;
function getTestCalenderDate(){
  if(!getIsTestEnv()) return null;
  if(testCalenderDate) return testCalenderDate;
  testCalenderDate = new Date(getTestCalenderDateYear(), getTestCalenderDateMonth() - 1, getTestCalenderDateDay()); // the month is 0-indexed
  return testCalenderDate;
}

let testAggreRawDataDetail1
function getTestAggreRawDataDetail1(){
  if(!getIsTestEnv()) return null;
  if(testAggreRawDataDetail1) return testAggreRawDataDetail1;
  testAggreRawDataDetail1 = getSettingSheet().getRange("test_aggre_raw_data_detail_1").getValue();
  return testAggreRawDataDetail1;
}

let testAggreRawDataDetail2
function getTestAggreRawDataDetail2(){
  if(!getIsTestEnv()) return null;
  if(testAggreRawDataDetail2) return testAggreRawDataDetail2;
  testAggreRawDataDetail2 = getSettingSheet().getRange("test_aggre_raw_data_detail_2").getValue();
  return testAggreRawDataDetail2;
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

let testCalcDcbResultGotEarlier;
function getTestCalcDcbResultGotEarlier(){
  if(!getIsTestEnv()) return null;
  if(testCalcDcbResultGotEarlier) return testCalcDcbResultGotEarlier;
  testCalcDcbResultGotEarlier = getSettingSheet().getRange("test_calc_dcb_result_got_earlier").getValue();
  return testCalcDcbResultGotEarlier;
}

let testCalcDcbResultGotLater;
function getTestCalcDcbResultGotLater(){
  if(!getIsTestEnv()) return null;
  if(testCalcDcbResultGotLater) return testCalcDcbResultGotLater;
  testCalcDcbResultGotLater = getSettingSheet().getRange("test_calc_dcb_result_got_later").getValue();
  return testCalcDcbResultGotLater;
}

let testCalcDcbResultExpectedEarlier;
function getTestCalcDcbResultExpectedEarlier(){
  if(!getIsTestEnv()) return null;
  if(testCalcDcbResultExpectedEarlier) return testCalcDcbResultExpectedEarlier;
  testCalcDcbResultExpectedEarlier = getSettingSheet().getRange("test_calc_dcb_result_expected_earlier").getValue();
  return testCalcDcbResultExpectedEarlier;
}

let testCalcDcbResultExpectedLater;
function getTestCalcDcbResultExpectedLater(){
  if(!getIsTestEnv()) return null;
  if(testCalcDcbResultExpectedLater) return testCalcDcbResultExpectedLater;
  testCalcDcbResultExpectedLater = getSettingSheet().getRange("test_calc_dcb_result_expected_later").getValue();
  return testCalcDcbResultExpectedLater;
}

let testCalcDabResultGotEarlier;
function getTestCalcDabResultGotEarlier(){
  if(!getIsTestEnv()) return null;
  if(testCalcDabResultGotEarlier) return testCalcDabResultGotEarlier;
  testCalcDabResultGotEarlier = getSettingSheet().getRange("test_calc_dab_result_got_earlier").getValue();
  return testCalcDabResultGotEarlier;
}

let testCalcDabResultGotLater;
function getTestCalcDabResultGotLater(){
  if(!getIsTestEnv()) return null;
  if(testCalcDabResultGotLater) return testCalcDabResultGotLater;
  testCalcDabResultGotLater = getSettingSheet().getRange("test_calc_dab_result_got_later").getValue();
  return testCalcDabResultGotLater;
}

let testCalcDabResultExpectedEarlier;
function getTestCalcDabResultExpectedEarlier(){
  if(!getIsTestEnv()) return null;
  if(testCalcDabResultExpectedEarlier) return testCalcDabResultExpectedEarlier;
  testCalcDabResultExpectedEarlier = getSettingSheet().getRange("test_calc_dab_result_expected_earlier").getValue();
  return testCalcDabResultExpectedEarlier;
}

let testCalcDabResultExpectedLater;
function getTestCalcDabResultExpectedLater(){
  if(!getIsTestEnv()) return null;
  if(testCalcDabResultExpectedLater) return testCalcDabResultExpectedLater;
  testCalcDabResultExpectedLater = getSettingSheet().getRange("test_calc_dab_result_expected_later").getValue();
  return testCalcDabResultExpectedLater;
}

let testCalcDcResultGotEarlier;
function getTestCalcDcResultGotEarlier(){
  if(!getIsTestEnv()) return null;
  if(testCalcDcResultGotEarlier) return testCalcDcResultGotEarlier;
  testCalcDcResultGotEarlier = getSettingSheet().getRange("test_calc_dc_result_got_earlier").getValue();
  return testCalcDcResultGotEarlier;
}

let testCalcDcResultGotLater;
function getTestCalcDcResultGotLater(){
  if(!getIsTestEnv()) return null;
  if(testCalcDcResultGotLater) return testCalcDcResultGotLater;
  testCalcDcResultGotLater = getSettingSheet().getRange("test_calc_dc_result_got_later").getValue();
  return testCalcDcResultGotLater;
}

let testCalcDcResultExpectedEarlier;
function getTestCalcDcResultExpectedEarlier(){
  if(!getIsTestEnv()) return null;
  if(testCalcDcResultExpectedEarlier) return testCalcDcResultExpectedEarlier;
  testCalcDcResultExpectedEarlier = getSettingSheet().getRange("test_calc_dc_result_expected_earlier").getValue();
  return testCalcDcResultExpectedEarlier;
}

let testCalcDcResultExpectedLater;
function getTestCalcDcResultExpectedLater(){
  if(!getIsTestEnv()) return null;
  if(testCalcDcResultExpectedLater) return testCalcDcResultExpectedLater;
  testCalcDcResultExpectedLater = getSettingSheet().getRange("test_calc_dc_result_expected_later").getValue();
  return testCalcDcResultExpectedLater;
}

let testCalcMcResultGot;
function getTestCalcMcResultGot(){
  if(!getIsTestEnv()) return null;
  if(testCalcMcResultGot) return testCalcMcResultGot;
  testCalcMcResultGot = getSettingSheet().getRange("test_calc_mc_result_got").getValue();
  return testCalcMcResultGot;
}

let testCalcMcResultExpected;
function getTestCalcMcResultExpected(){
  if(!getIsTestEnv()) return null;
  if(testCalcMcResultExpected) return testCalcMcResultExpected;
  testCalcMcResultExpected = getSettingSheet().getRange("test_calc_mc_result_expected").getValue();
  return testCalcMcResultExpected;
}

let testReportResultGotRange;
function getTestReportResultGotRange(){
  if(!getIsTestEnv()) return null;
  if(testReportResultGotRange) return testReportResultGotRange;
  testReportResultGotRange = getSettingSheet().getRange("test_report_result_got");
  return testReportResultGotRange;
}

let testReportResultGot;
function getTestReportResultGot(){
  if(!getIsTestEnv()) return null;
  if(testReportResultGot) return testReportResultGot;
  testReportResultGot = getTestReportResultGotRange().getValue();
  return testReportResultGot;
}

let testReportResultExpected;
function getTestReportResultExpected(){
  if(!getIsTestEnv()) return null;
  if(testReportResultExpected) return testReportResultExpected;
  testReportResultExpected = getSettingSheet().getRange("test_report_result_expected").getValue();
  return testReportResultExpected;
}

let testUpdateGraphDcbResultGot;
function getTestUpdateGraphDcbResultGot(){
  if(!getIsTestEnv()) return null;
  if(testUpdateGraphDcbResultGot) return testUpdateGraphDcbResultGot;
  testUpdateGraphDcbResultGot = getSettingSheet().getRange("test_update_graph_dcb_result_got").getValue();
  return testUpdateGraphDcbResultGot;
}

let testUpdateGraphDcbResultExpected;
function getTestUpdateGraphDcbResultExpected(){
  if(!getIsTestEnv()) return null;
  if(testUpdateGraphDcbResultExpected) return testUpdateGraphDcbResultExpected;
  testUpdateGraphDcbResultExpected = getSettingSheet().getRange("test_update_graph_dcb_result_expected").getValue();
  return testUpdateGraphDcbResultExpected;
}

let testUpdateGraphDabResultGot;
function getTestUpdateGraphDabResultGot(){
  if(!getIsTestEnv()) return null;
  if(testUpdateGraphDabResultGot) return testUpdateGraphDabResultGot;
  testUpdateGraphDabResultGot = getSettingSheet().getRange("test_update_graph_dab_result_got").getValue();
  return testUpdateGraphDabResultGot;
}

let testUpdateGraphDabResultExpected;
function getTestUpdateGraphDabResultExpected(){
  if(!getIsTestEnv()) return null;
  if(testUpdateGraphDabResultExpected) return testUpdateGraphDabResultExpected;
  testUpdateGraphDabResultExpected = getSettingSheet().getRange("test_update_graph_dab_result_expected").getValue();
  return testUpdateGraphDabResultExpected;
}
