/** @OnlyCurrentDoc */

function executeReport(){
  // 報告対象月を取得する
  const targetYear = getCurrentYyyy();
  const targetMonth = getCurrentMm();
  console.log("targetYear, targetMonth: " + targetYear + ", " + targetMonth);

  // 対象月直近日の現預金残高を取得する
  const {lastCashBalanceDate, lastCashBalance} = getLastDateAndCashBalanceAt(targetYear, targetMonth);
  // 対象月の先月末の現預金残高を取得する
  const {lastCashBalanceDateAtPreviousMonth, lastCashBalanceAtPreviousMonth} = getLastDateAndCashBalancePreviousMonthFrom(targetYear, targetMonth);
  // 対象月直近日の総資産残高を取得する
  const {lastTotalAssetBalanceDate, lastTotalAssetBalance} = getLastDateAndTotalAssetBalanceAt(targetYear, targetMonth);
  // 対象月の先月末の総資産残高を取得する
  const {lastTotalAssetBalanceDateAtPreviousMonth, lastTotalAssetBalanceAtPreviousMonth} = getLastDateAndTotalAssetBalancePreviousMonthFrom(targetYear, targetMonth);

  // 対象月の項目別月次集計結果を取得する
  const summaryByCategoryAtCurrentMonth = extractReportInfoFromSummaryByCategory(fetchMonthlySummaryByCategory(targetYear, targetMonth));
  // 対象月の先月分の項目別月次集計結果を取得する
  const {previousMonthYyyy, previousMonthMm} = getPreviousYearMonth(targetYear, targetMonth);
  const summaryByCategoryAtPreviousMonth = extractReportInfoFromSummaryByCategory(fetchMonthlySummaryByCategory(previousMonthYyyy, previousMonthMm));
  console.log(summaryByCategoryAtCurrentMonth);
  console.log(summaryByCategoryAtPreviousMonth);
}

function extractReportInfoFromSummaryByCategory(summaryByCategory){
  return summaryByCategory.filter(s => ["収支合計", "総合計", "小計"].includes(s.middleCategory)).map(
    s => {
      return {
        category: s.largeCategory,
        amount:   s.amount,
      };
    }
  );
}

function fetchMonthlySummaryByCategory(targetYear, targetMonth){
  const targetSheetName = getCalcMcExportSheetPrefix() + "_" + targetYear + targetMonth;
  const rawDetailsValue = getThisSpreadSheet().getSheetByName(targetSheetName).getRange(getCalcMcExportAddr()).getValue();
  return rawDetailsValue.split("¥n").map(
    row => {
      const rowElems = row.split("#&#"); // TODO: split して二次元配列にして返すところまでを共通化する
      return {
        yearMonth:      rowElems[0], // 対象年月
        largeCategory:  rowElems[1], // 大項目
        middleCategory: rowElems[2], // 中項目
        amount:         rowElems[3], // 金額
      };
    }
  );
}

function getLastDateAndTotalAssetBalancePreviousMonthFrom(targetYear, targetMonth){
  const {previousMonthYyyy, previousMonthMm} = getPreviousYearMonth(targetYear, targetMonth);
  const {lastTotalAssetBalanceDate, lastTotalAssetBalance} = getLastDateAndTotalAssetBalanceAt(previousMonthYyyy, previousMonthMm);
  const lastTotalAssetBalanceDateAtPreviousMonth = lastTotalAssetBalanceDate;
  const lastTotalAssetBalanceAtPreviousMonth = lastTotalAssetBalance;
  return {lastTotalAssetBalanceDateAtPreviousMonth, lastTotalAssetBalanceAtPreviousMonth};
}

function getLastDateAndTotalAssetBalanceAt(targetYear, targetMonth){
  const {lastDate, lastBalance} = getLastDateAndBalanceAt(targetYear, targetMonth, fetchDailyTotalAssetBalanceDetails);
  const lastTotalAssetBalanceDate = lastDate;
  const lastTotalAssetBalance = lastBalance;
  return {lastTotalAssetBalanceDate, lastTotalAssetBalance};
}

function fetchDailyTotalAssetBalanceDetails(targetYear, targetMonth){
  const targetSheetName = getCalcDabExportSheetPrefix() + "_" + targetYear + targetMonth;
  const rawDetailsValue = getThisSpreadSheet().getSheetByName(targetSheetName).getRange(getCalcDabExportAddr()).getValue();
  return rawDetailsValue.split("¥n").map(
    row => {
      const rowElems = row.split("#&#");
      return {
        date:    rowElems[0], // 日付
        balance: rowElems[1], // 残高
      };
    }
  );
}

function getLastDateAndCashBalancePreviousMonthFrom(targetYear, targetMonth){
  const {previousMonthYyyy, previousMonthMm} = getPreviousYearMonth(targetYear, targetMonth);
  const {lastCashBalanceDate, lastCashBalance} = getLastDateAndCashBalanceAt(previousMonthYyyy, previousMonthMm);
  const lastCashBalanceDateAtPreviousMonth = lastCashBalanceDate;
  const lastCashBalanceAtPreviousMonth = lastCashBalance;
  return {lastCashBalanceDateAtPreviousMonth, lastCashBalanceAtPreviousMonth};
}

function getLastDateAndCashBalanceAt(targetYear, targetMonth){
  const {lastDate, lastBalance} = getLastDateAndBalanceAt(targetYear, targetMonth, fetchDailyCashBalanceDetails);
  const lastCashBalanceDate = lastDate;
  const lastCashBalance = lastBalance;
  return {lastCashBalanceDate, lastCashBalance};
}

function getLastDateAndBalanceAt(targetYear, targetMonth, fetchDetailsFunc){
  const dailyBalanceDetails = fetchDetailsFunc(targetYear, targetMonth);
  const lastBalanceDetail = dailyBalanceDetails.length == 0 ? null : dailyBalanceDetails[dailyBalanceDetails.length - 1];
  const lastDate = lastBalanceDetail == null ? null : lastBalanceDetail.date;
  const lastBalance = lastBalanceDetail == null ? null : lastBalanceDetail.balance;
  return {lastDate, lastBalance};
}

function fetchDailyCashBalanceDetails(targetYear, targetMonth){
  const targetSheetName = getCalcDcbExportSheetPrefix() + "_" + targetYear + targetMonth;
  const rawDetailsValue = getThisSpreadSheet().getSheetByName(targetSheetName).getRange(getCalcDcbExportAddr()).getValue();
  return rawDetailsValue.split("¥n").map(
    row => {
      const rowElems = row.split("#&#");
      return {
        date:    rowElems[0], // 日付
        balance: rowElems[1], // 残高
      };
    }
  );
}
