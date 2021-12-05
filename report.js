/** @OnlyCurrentDoc */

function executeReport(){
  postConsoleAndSlackJobStart("execute report");
  try {
    doReport();
  } catch(error){
    handleExpectedError(error);
  }
  postConsoleAndSlackJobEnd("execute report");
}

function doReport(){
  // 報告対象月を取得する
  const targetYear = getCurrentYyyy();
  const targetMonth = getCurrentMm();
  console.log("targetYear, targetMonth: " + targetYear + ", " + targetMonth);

  // 対象月直近日の現預金残高を取得する
  const {lastCashBalanceDate, lastCashBalance} = getLastDateAndCashBalanceAt(targetYear, targetMonth);
  // 対象月の先月末の現預金残高を取得する
  const {lastCashBalanceAtPreviousMonth} = getLastDateAndCashBalancePreviousMonthFrom(targetYear, targetMonth);
  // 対象月直近日の総資産残高を取得する
  const {lastTotalAssetBalanceDate, lastTotalAssetBalance} = getLastDateAndTotalAssetBalanceAt(targetYear, targetMonth);
  // 対象月の先月末の総資産残高を取得する
  const {lastTotalAssetBalanceAtPreviousMonth} = getLastDateAndTotalAssetBalancePreviousMonthFrom(targetYear, targetMonth);

  // 対象月の項目別月次集計結果を取得する
  const summaryByCategoryAtCurrentMonth = getMonthlySummaryByCategoryForReportAt(targetYear, targetMonth);
  // 対象月の先月分の項目別月次集計結果を取得する
  const summaryByCategoryAtPreviousMonth = getMonthlySummaryByCategoryForReportPreviousMonthFrom(targetYear, targetMonth);

  // 対象週(直近3週分)の項目別週次集計結果を取得する
  const weeklySummaryByCategories = getIntRangeFromZero(3).map(i => getWeeklySummaryByCategoryForReportAt(i));

  // グラフ画像URLを用意する
  const imageUrls = [
    {
      title: "daily balance: cash",
      value: "Daily transition graph of cash and deposit balance. <" + getGraphDcbImageUrl() + "|link>",
      url: getGraphDcbImageUrl(),
    },
    {
      title: "daily balance: total asset",
      value: "Daily transition graph of total asset balance. <" + getGraphDabImageUrl() + "|link>",
      url: getGraphDabImageUrl(),
    },
  ];

  // レポートをまとめて、slack に通知する
  postSlackReportingChannel(
    getReportMessage(
      lastCashBalance,
      lastCashBalanceDate,
      lastCashBalanceAtPreviousMonth,
      lastTotalAssetBalance,
      lastTotalAssetBalanceDate,
      lastTotalAssetBalanceAtPreviousMonth,
      summaryByCategoryAtCurrentMonth,
      summaryByCategoryAtPreviousMonth,
      weeklySummaryByCategories),
    imageUrls);
}

// レポートするメッセージを編集して返す
function getReportMessage(
  lastCashBalance,
  lastCashBalanceDate,
  lastCashBalanceAtPreviousMonth,
  lastTotalAssetBalance,
  lastTotalAssetBalanceDate,
  lastTotalAssetBalanceAtPreviousMonth,
  summaryByCategoryAtCurrentMonth,
  summaryByCategoryAtPreviousMonth,
  weeklySummaryByCategories){
  let reportMessage = ""; // TODO: ヒアドキュメント化
  reportMessage += "# 残高の概要\n";
  reportMessage += ("- 直近現預金残高: " + formatNumStr(lastCashBalance) + "\n");
  reportMessage += ("  - (日付: " + lastCashBalanceDate + ")\n");
  reportMessage += ("  - (前月末残: " + formatNumStr(lastCashBalanceAtPreviousMonth) + ")\n");
  reportMessage += ("- 直近総資産残高: " + formatNumStr(lastTotalAssetBalance) + "\n");
  reportMessage += ("  - (日付: " + lastTotalAssetBalanceDate + ")\n");
  reportMessage += ("  - (前月末残: " + formatNumStr(lastTotalAssetBalanceAtPreviousMonth) + ")\n");
  reportMessage += "\n";
  reportMessage += "# 収支内訳の概要\n";
  reportMessage += "## 月次\n";
  reportMessage += getMessageAboutSummaryByCategory(summaryByCategoryAtCurrentMonth, "今月");
  reportMessage += getMessageAboutSummaryByCategory(summaryByCategoryAtPreviousMonth, "先月");
  reportMessage += "## 週次\n";
  reportMessage += getMessageAboutSummaryByCategory(weeklySummaryByCategories[0], "今週");
  reportMessage += getMessageAboutSummaryByCategory(weeklySummaryByCategories[1], "先週");
  reportMessage += getMessageAboutSummaryByCategory(weeklySummaryByCategories[2], "先々週");
  return reportMessage;
}

// 収支内訳の概要メッセージを対象月の項目別月次集計結果をもとに編集して返す
function getMessageAboutSummaryByCategory(summaryByCategory, headerTitle){
  let reportMessage = "";
  if(summaryByCategory.length > 0){
    reportMessage += ("## " + headerTitle + " (" + summaryByCategory[0].yearMonth + ")\n");
    summaryByCategory.forEach(
      s => {
        if(s.middleCategory == "収支合計"){
          reportMessage += ("- " + s.largeCategory + ": " + formatNumStr(s.amount) + "\n");
        } else if(s.middleCategory == "総合計"){
          reportMessage += ("  - " + s.largeCategory + ": " + formatNumStr(s.amount) + "\n");
        } else if(["小計", "未定義(大項目)"].includes(s.middleCategory)){
          reportMessage += ("    - " + s.largeCategory + ": " + formatNumStr(s.amount) + "\n");
        }
      }
    );
  }
  return reportMessage;
}

// 項目別月次集計結果からレポーティング対象明細を抽出して返す
function extractReportInfoFromSummaryByCategory(summaryByCategory){
  return summaryByCategory.filter(s => ["収支合計", "総合計", "小計", "未定義(大項目)"].includes(s.middleCategory));
}

// 項目別月次集計結果を取得して返す
function fetchMonthlySummaryByCategory(targetYear, targetMonth){
  return fetchDetailsFromCell(
    getCalcMcExportSheetPrefix() + "_" + targetYear + targetMonth,
    getCalcMcExportAddr(),
    rowElems => {
      return {
        yearMonth:      rowElems[0], // 対象年月
        largeCategory:  rowElems[1], // 大項目
        middleCategory: rowElems[2], // 中項目
        amount:         rowElems[3], // 金額
      };
    }
  );
}

// 項目別週次集計結果を取得して返す
function fetchWeeklySummaryByCategory(weeksAgo){
  return fetchDetailsFromCell(
    getCalcWcExportSheetPrefix() + "_" + weeksAgo,
    getCalcWcExportAddr(),
    rowElems => {
      return {
        startDate:      rowElems[0], // 対象週初日
        largeCategory:  rowElems[1], // 大項目
        middleCategory: rowElems[2], // 中項目
        amount:         rowElems[3], // 金額
      };
    }
  );
}

// 指定月の前月の、項目別月次集計結果を取得する
function getMonthlySummaryByCategoryForReportPreviousMonthFrom(targetYear, targetMonth){
  const {previousMonthYyyy, previousMonthMm} = getPreviousYearMonth(targetYear, targetMonth);
  return getMonthlySummaryByCategoryForReportAt(previousMonthYyyy, previousMonthMm);
}

// 指定月の、項目別月次集計結果を取得する
function getMonthlySummaryByCategoryForReportAt(targetYear, targetMonth){
  return extractReportInfoFromSummaryByCategory(fetchMonthlySummaryByCategory(targetYear, targetMonth));
}

// 指定週の、項目別週次集計結果を取得する
function getWeeklySummaryByCategoryForReportAt(weeksAgo){
  return extractReportInfoFromSummaryByCategory(fetchWeeklySummaryByCategory(weeksAgo));
}

// 指定月の前月末の、総資産残高推移の基準日と残高を取得して返す
function getLastDateAndTotalAssetBalancePreviousMonthFrom(targetYear, targetMonth){
  const {previousMonthYyyy, previousMonthMm} = getPreviousYearMonth(targetYear, targetMonth);
  const {lastTotalAssetBalanceDate, lastTotalAssetBalance} = getLastDateAndTotalAssetBalanceAt(previousMonthYyyy, previousMonthMm);
  const lastTotalAssetBalanceDateAtPreviousMonth = lastTotalAssetBalanceDate;
  const lastTotalAssetBalanceAtPreviousMonth = lastTotalAssetBalance;
  return {lastTotalAssetBalanceDateAtPreviousMonth, lastTotalAssetBalanceAtPreviousMonth};
}

// 指定月末の、総資産残高推移の基準日と残高を取得して返す
function getLastDateAndTotalAssetBalanceAt(targetYear, targetMonth){
  const {lastDate, lastBalance} = getLastDateAndBalanceAt(targetYear, targetMonth, fetchDailyTotalAssetBalanceDetails);
  const lastTotalAssetBalanceDate = lastDate;
  const lastTotalAssetBalance = lastBalance;
  return {lastTotalAssetBalanceDate, lastTotalAssetBalance};
}

// 日次総資産残高集計結果を取得して返す
function fetchDailyTotalAssetBalanceDetails(targetYear, targetMonth){
  return fetchDetailsFromCell(
    getCalcDabExportSheetPrefix() + "_" + targetYear + targetMonth,
    getCalcDabExportAddr(),
    rowElems => {
      return {
        date:    rowElems[0], // 日付
        balance: rowElems[1], // 残高
      };
    }
  );
}

// 指定月の前月末の、現預金残高推移の基準日と残高を取得して返す
function getLastDateAndCashBalancePreviousMonthFrom(targetYear, targetMonth){
  const {previousMonthYyyy, previousMonthMm} = getPreviousYearMonth(targetYear, targetMonth);
  const {lastCashBalanceDate, lastCashBalance} = getLastDateAndCashBalanceAt(previousMonthYyyy, previousMonthMm);
  const lastCashBalanceDateAtPreviousMonth = lastCashBalanceDate;
  const lastCashBalanceAtPreviousMonth = lastCashBalance;
  return {lastCashBalanceDateAtPreviousMonth, lastCashBalanceAtPreviousMonth};
}

// 指定月末の、現預金残高推移の基準日と残高を取得して返す
function getLastDateAndCashBalanceAt(targetYear, targetMonth){
  const {lastDate, lastBalance} = getLastDateAndBalanceAt(targetYear, targetMonth, fetchDailyCashBalanceDetails);
  const lastCashBalanceDate = lastDate;
  const lastCashBalance = lastBalance;
  return {lastCashBalanceDate, lastCashBalance};
}

// 指定月末の、残高推移の基準日と残高を取得して返す
function getLastDateAndBalanceAt(targetYear, targetMonth, fetchDetailsFunc){
  const dailyBalanceDetails = fetchDetailsFunc(targetYear, targetMonth);
  const lastBalanceDetail = getLastElemFrom(dailyBalanceDetails, {date: null, balance: null});
  return {
    lastDate: lastBalanceDetail.date,
    lastBalance: lastBalanceDetail.balance,
  };
}

// 日次現預金残高集計結果を取得して返す
function fetchDailyCashBalanceDetails(targetYear, targetMonth){
  return fetchDetailsFromCell(
    getCalcDcbExportSheetPrefix() + "_" + targetYear + targetMonth,
    getCalcDcbExportAddr(),
    rowElems => {
      return {
        date:    rowElems[0], // 日付
        balance: rowElems[1], // 残高
      };
    }
  );
}
