/** @OnlyCurrentDoc */

function executeReport(){
  postConsoleAndSlackJobStart("execute report");

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
  const summaryByCategoryAtCurrentMonth = extractReportInfoFromSummaryByCategory(fetchMonthlySummaryByCategory(targetYear, targetMonth));
  // 対象月の先月分の項目別月次集計結果を取得する
  const {previousMonthYyyy, previousMonthMm} = getPreviousYearMonth(targetYear, targetMonth);
  const summaryByCategoryAtPreviousMonth = extractReportInfoFromSummaryByCategory(fetchMonthlySummaryByCategory(previousMonthYyyy, previousMonthMm));

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
      summaryByCategoryAtPreviousMonth));

  postConsoleAndSlackJobEnd("execute report");
}

// TODO: 以下、コメントを
function getReportMessage(
  lastCashBalance,
  lastCashBalanceDate,
  lastCashBalanceAtPreviousMonth,
  lastTotalAssetBalance,
  lastTotalAssetBalanceDate,
  lastTotalAssetBalanceAtPreviousMonth,
  summaryByCategoryAtCurrentMonth,
  summaryByCategoryAtPreviousMonth){
  let reportMessage = "";
  reportMessage += "# 残高の概要\n";
  reportMessage += ("- 直近現預金残高: " + formatNumStr(lastCashBalance) + "\n");
  reportMessage += ("  - (日付: " + lastCashBalanceDate + ")\n");
  reportMessage += ("  - (前月末残: " + formatNumStr(lastCashBalanceAtPreviousMonth) + ")\n");
  reportMessage += ("- 直近総資産残高: " + formatNumStr(lastTotalAssetBalance) + "\n");
  reportMessage += ("  - (日付: " + lastTotalAssetBalanceDate + ")\n");
  reportMessage += ("  - (前月末残: " + formatNumStr(lastTotalAssetBalanceAtPreviousMonth) + ")\n");
  reportMessage += "\n";
  reportMessage += "# 収支内訳の概要\n";
  reportMessage += getMessageAboutSummaryByCategory(summaryByCategoryAtCurrentMonth, "今月");
  reportMessage += getMessageAboutSummaryByCategory(summaryByCategoryAtPreviousMonth, "先月");
  return reportMessage;
}

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

function extractReportInfoFromSummaryByCategory(summaryByCategory){
  return summaryByCategory.filter(s => ["収支合計", "総合計", "小計", "未定義(大項目)"].includes(s.middleCategory));
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

// https://qiita.com/ryo-yamaoka/items/7677ee4486cf395ce9bc
// https://qiita.com/84zume/items/69f546e6ddee464e0473
// https://ja.javascript.info/custom-errors
class HogeError extends Error {
  constructor(message, needsMention) {
    super(message);
    this.name = "HogeError";
    this.needsMention = needsMention;
  }
}

// memo: エラーハンドリング:
//   - 処理継続するもの (メンション有無はオプション) (ワーニング to alert チャンネル) (ex. アグリに失敗したとき...たまに起きることを見込んでいる)
//   - 処理中断するが正常終了扱いにするもの (メンション有無はオプション) (想定エラー to alert チャンネル)
//   - 処理中断して異常終了扱いにするもの (メンション有無はオプション) (想定外エラー to alert チャンネル)
function test(){
  // <!channel>
  // postConsoleAndSlackJobStart("<!channel> mention test");
  // error handling
  try {
    // throw new HogeError("failed to hoge", true);
    const hoge = undefined[0];
  } catch(error) {
    if(error instanceof HogeError){
      console.error(printError(error));
      let headerMessage = "expected error occurred: ";
      if(error.needsMention){
        headerMessage = "<!channel> " + headerMessage;
      }
      postConsoleAndSlackJobStart(headerMessage + printError(error));
    } else {
      postConsoleAndSlackJobStart("unexpected error occurred: " + printError(error));
      throw error;
    }
  }
}

function printError(error){
  return "\n" +
         "[name]       " + error.name + "\n" +
         // "[place]      " + error.fileName + " (lines:" + error.lineNumber + ")\n" +
         "[message]    " + error.message + "\n" +
         "[stacktrace] " + error.stack;
}
