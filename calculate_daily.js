/** @OnlyCurrentDoc */

// 最新月分の日次集計を行う
function executeCalcDailySummaryCurrentMonth(){
  postConsoleAndSlackJobStart("execute calc daily summary current month");
  try {
    calcDailySummary(0);
  } catch(error){
    handleError(error);
  }
  postConsoleAndSlackJobEnd("execute calc daily summary current month");
}

// 最新月の前月分の日次集計を行う
function executeCalcDailySummaryPreviousMonth(){
  postConsoleAndSlackJobStart("execute calc daily summary previous month");
  try {
    calcDailySummary(1);
    calcDailySummary(0); // 最新月分も洗い替えておく
  } catch(error){
    handleError(error);
  }
  postConsoleAndSlackJobEnd("execute calc daily summary previous month");
}

// 指定月の日次集計を行う
function calcDailySummary(monthsAgo){
  // 集計対象月を取得する
  const {targetYear, targetMonth} = getAggregateTargetYearAndMonth(monthsAgo);
  console.log("targetYear, targetMonth: " + targetYear + ", " + targetMonth);

  // 集計対象月の明細データを取得する
  const aggregatedDetails = fetchAggregatedDetails(targetYear, targetMonth);
  console.log("fetch and parse aggregatedDetails finished: length: " + aggregatedDetails.length);

  // 日次現預金残高集計を行い、結果シートに出力する
  calcAndExportDailyBalanceCash(aggregatedDetails, targetYear, targetMonth);

  // 日次総資産残高集計を行い、結果シートに出力する
  calcAndExportDailyBalanceTotalAsset(aggregatedDetails, targetYear, targetMonth);

  // 日次項目別集計を行い、結果シートに出力する
  calcAndExportDailySummaryByCategory(aggregatedDetails, targetYear, targetMonth);
}

// 日次項目別集計を行い、結果シートに出力する
function calcAndExportDailySummaryByCategory(aggregatedDetails, targetYear, targetMonth){
  // 明細データを大項目&中項目別に日次で集計する
  console.log("start calcAndExportDailySummaryByCategory");
  // 集計対象明細を取得
  const summaryByCategoryTargetDetails = getSummaryByCategoryTargetDetails(aggregatedDetails);
  // 集計対象の日付を収集
  const summaryByCategoryTargetDates = extractSortedDatesFromDetails(summaryByCategoryTargetDetails);
  // 日次で大項目&中項目別に集計
  const summaryByCategoryAndDates = summaryAmountByCategoryAndDate(summaryByCategoryTargetDates, summaryByCategoryTargetDetails);
  console.log("summaryAmountByCategoryAndDate finished: length: " + summaryByCategoryAndDates.length);
  // 集計結果明細を出力対象シートに貼り付ける
  pasteDailyCalcResultByEachMonth(
    targetYear,
    targetMonth,
    summaryByCategoryAndDates,
    getCalcDcImportSheetPrefix(),
    getCalcDcImportAddrBeforeClosingDay(),
    getCalcDcImportAddrAfterClosingDay());
  console.log("end calcAndExportDailySummaryByCategory");
}

// 日次総資産残高集計を行い、結果シートに出力する
function calcAndExportDailyBalanceTotalAsset(aggregatedDetails, targetYear, targetMonth){
  // 明細データのうち全金融機関分の入出金を集計し、日次の全資産分の残高推移を求める
  console.log("start calcAndExportDailyBalanceTotalAsset");
  // 集計対象明細を取得
  const balanceTotalAssetTargetDetails = getBalanceTotalAssetTargetDetails(aggregatedDetails);
  // 集計対象の日付を収集
  const balanceTotalAssetTargetDates = extractSortedDatesFromDetails(balanceTotalAssetTargetDetails);
  // 前月末残高を取得する
  const previousMonthEndBalanceTotalAsset = Number(getPreviousMonthEndBalance(targetYear, targetMonth).assetTotalEndBalance);
  // 日次で総資産残高を集計
  const totalAssetBalanceByDates = accumulateBalanceByDate(previousMonthEndBalanceTotalAsset, balanceTotalAssetTargetDates, balanceTotalAssetTargetDetails);
  console.log("calc totalAssetBalanceByDates finished: length: " + totalAssetBalanceByDates.length);
  // 対象月の末残高設定を更新する
  const updatedTotalAssetBalanceAtMonthEnd = updateTargetMonthEndBalance(totalAssetBalanceByDates, previousMonthEndBalanceTotalAsset, targetYear, targetMonth, getAssetTotalEndBalanceRange());
  console.log("total asset end balance updated: updatedTotalAssetBalanceAtMonthEnd: " + updatedTotalAssetBalanceAtMonthEnd);
  // 集計結果明細を出力対象シートに貼り付ける
  pasteDailyCalcResultByEachMonth(
    targetYear,
    targetMonth,
    totalAssetBalanceByDates,
    getCalcDabImportSheetPrefix(),
    getCalcDabImportAddrBeforeClosingDay(),
    getCalcDabImportAddrAfterClosingDay());
  console.log("end calcAndExportDailyBalanceTotalAsset");
}

// 日次現預金残高集計を行い、結果シートに出力する
function calcAndExportDailyBalanceCash(aggregatedDetails, targetYear, targetMonth){
  // 明細データのうち現預金の入出金を集計し、日次の残高推移を求める
  console.log("start calcAndExportDailyBalanceCash");
  // 集計対象明細を取得
  const balanceCashTargetDetails = getBalanceCashTargetDetails(aggregatedDetails);
  // 集計対象の日付を収集
  const balanceCashTargetDates = extractSortedDatesFromDetails(balanceCashTargetDetails);
  // 前月末残高を取得する
  const previousMonthEndBalanceCash = Number(getPreviousMonthEndBalance(targetYear, targetMonth).cashEndBalance);
  // 日次で現預金残高を集計
  const cashBalanceByDates = accumulateBalanceByDate(previousMonthEndBalanceCash, balanceCashTargetDates, balanceCashTargetDetails);
  console.log("calc cashBalanceByDates finished: length: " + cashBalanceByDates.length);
  // 対象月の末残高設定を更新する
  const updatedCashBalanceAtMonthEnd = updateTargetMonthEndBalance(cashBalanceByDates, previousMonthEndBalanceCash, targetYear, targetMonth, getCashEndBalanceRange());
  console.log("cash end balance updated: updatedCashBalanceAtMonthEnd: " + updatedCashBalanceAtMonthEnd);
  // 集計結果明細を出力対象シートに貼り付ける
  pasteDailyCalcResultByEachMonth(
    targetYear,
    targetMonth,
    cashBalanceByDates,
    getCalcDcbImportSheetPrefix(),
    getCalcDcbImportAddrBeforeClosingDay(),
    getCalcDcbImportAddrAfterClosingDay());
  console.log("end calcAndExportDailyBalanceCash");
}

// 日次集計結果を、早い月分と遅い月分に振り分け、それぞれを結果シートに貼り付ける
function pasteDailyCalcResultByEachMonth(targetYear, targetMonth, dailyCalcResults, targetSheetPrefix, pasteAddrBeforeClosingDay, pasteAddrAfterClosingDay){
  // 集計した明細を、早い月分と遅い月分に振り分ける
  const earlierYearMonth = targetYear + "/" + targetMonth;
  const earlierMonthDailyCalcResults = dailyCalcResults.filter(s => s.date.indexOf(earlierYearMonth) == 0);
  const laterMonthDailyCalcResults = dailyCalcResults.filter(s => s.date.indexOf(earlierYearMonth) != 0);

  // 集計結果を各月分の結果シートに貼り付け
  console.log("start to paste dailyCalcResultByEachMonth: targetSheetPrefix: " + targetSheetPrefix);
  // 早い月分
  if(earlierMonthDailyCalcResults.length > 0){ // TODO: n件 -> 0件に更新された場合に反映されない問題
    console.log("start to paste earlierMonthDailyCalcResults: earlierMonthDailyCalcResults.length: " + earlierMonthDailyCalcResults.length);
    exportMonthlyResultDetails(earlierMonthDailyCalcResults, targetYear, targetMonth, targetSheetPrefix, pasteAddrAfterClosingDay);
  } else {
    console.log("pasting earlierMonthDailyCalcResults was skipped");
  }
  // 遅い月分
  if(laterMonthDailyCalcResults.length > 0){ // TODO: n件 -> 0件に更新された場合に反映されない問題
    console.log("start to paste laterMonthDailyCalcResults: laterMonthDailyCalcResults.length: " + laterMonthDailyCalcResults.length);
    const {nextMonthYyyy, nextMonthMm} = getNextYearMonth(targetYear, targetMonth);
    exportMonthlyResultDetails(laterMonthDailyCalcResults, nextMonthYyyy, nextMonthMm, targetSheetPrefix, pasteAddrBeforeClosingDay);
  } else {
    console.log("pasting laterMonthDailyCalcResults was skipped");
  }
}

// 対象月の末残高レコードを更新する
function updateTargetMonthEndBalance(accumulatedBalanceByDates, previousMonthEndBalance, targetYear, targetMonth, targetRange){
  // 対象月の末残高を取得
  const balanceAtMonthEnd = getLastElemFrom(accumulatedBalanceByDates, {balance: previousMonthEndBalance}).balance;
  // 対象月の末残高表における index を取得
  const targetMonthIndexAtEndBalances = getEndBalances()
    .map(b => b.calcYear + "/" + b.calcMonth)
    .indexOf(targetYear + "/" + targetMonth);
  // 該当 cell を取得し、更新する
  targetRange
    .getCell(targetMonthIndexAtEndBalances + 2, 1) // getCell は 1 始まりなので +1, また該当 range にはヘッダー行が含まれるのでその分も +1
    .setValue(balanceAtMonthEnd);

  return balanceAtMonthEnd;
}

// 日次で項目別に集計して返す
function summaryAmountByCategoryAndDate(targetDates, targetDetails){
  return targetDates.map(
    date => {
      // 対象日の明細を抽出
      const targetDateDetails = targetDetails.filter(d => d.date == date);
      // 対象日における大項目&中項目を重複排除の上抽出
      const categories = getUniqueArrayFrom(targetDateDetails.map(d => d.largeCategoryName + ">>" + d.middleCategoryName));
      // 対象日の明細において、大項目&中項目別に集計
      return categories.map(
        c => {
          const sumAmountByDateAndCategory = sumAmountFromDetails(targetDateDetails.filter(d => d.largeCategoryName + ">>" + d.middleCategoryName == c));
          return {
            date:         date,                       // 日付
            categoryName: c,                          // 項目 (大項目&中項目)
            amount:       sumAmountByDateAndCategory, // 金額
          };
        }
      );
    }
  ).flat(1);
}

// 日次で残高を集計して返す
function accumulateBalanceByDate(initialBalance, targetDates, targetDetails){
  let accumulatedBalance = initialBalance;
  return targetDates.map(
    date => {
      // 残高を集計
      const sumAmountAtTargetDate = sumAmountFromDetails(targetDetails.filter(d => d.date == date));
      accumulatedBalance += sumAmountAtTargetDate;

      return {
        date:    date,
        balance: accumulatedBalance,
      };
    }
  );
}

// 前月末残高を返す
function getPreviousMonthEndBalance(targetYear, targetMonth){
  const {previousMonthYyyy, previousMonthMm} = getPreviousYearMonth(targetYear, targetMonth);
  return getEndBalances().find(b => b.calcYear == previousMonthYyyy && b.calcMonth == previousMonthMm);
}

// 日次項目別集計の対象明細を返す
function getSummaryByCategoryTargetDetails(aggregatedDetails){
  return aggregatedDetails.filter(d => d.isCalcTarget == 1);
}

// 日次総資産残高集計の対象明細を返す
function getBalanceTotalAssetTargetDetails(aggregatedDetails){
  // 明細データを集計対象のものに絞り込む ※相手勘定の無い振替(クレカ引落のカード側の明細が無い)が一部あり、振替は含めない
  return aggregatedDetails.filter(d => d.isCalcTarget == 1);
}

// 日次現預金残高集計の対象明細を返す
function getBalanceCashTargetDetails(aggregatedDetails){
  // 現預金に該当する保有金融機関名を取得する
  const assetConfigs = getAssetConfigs();
  const cashAssetNames = assetConfigs.filter(a => a.isCash == 1).map(a => a.assetName);

  // 明細データを集計対象のものに絞り込む ※現預金の口座に限定して動きを見るため振替も含める
  return aggregatedDetails
    .filter(d => d.isCalcTarget == 1 || d.isTransfer == 1) // 計算対象もしくは振替の明細に絞る
    .filter(d => cashAssetNames.includes(d.assetName)); // 現預金口座の明細に絞る
}

// 明細から日付 (一意かつ昇順) を収集する
function extractSortedDatesFromDetails(targetDetails) {
  const dates = getUniqueArrayFrom(targetDetails.map(d => d.date)); // TODO: 日が飛び飛びになる場合に間の残高を横置きで埋める考慮
  dates.sort(); // 昇順にソート

  return dates;
}

// 集計対象月の明細データを取得し、 hash の配列に parse したものを返す
function fetchAggregatedDetails(targetYear, targetMonth) {
  return fetchDetailsFromCell(
    getAggreExportSheetPrefix() + "_" + targetYear + targetMonth,
    getAggreExportAddr(),
    rowElems => {
      return {
        isCalcTarget:       rowElems[0], // 計算対象
        date:               rowElems[1], // 日付
        contents:           rowElems[2], // 内容
        amount:             rowElems[3], // 金額(円)
        assetName:          rowElems[4], // 保有金融機関
        largeCategoryName:  rowElems[5], // 大項目
        middleCategoryName: rowElems[6], // 中項目
        memo:               rowElems[7], // メモ
        isTransfer:         rowElems[8], // 振替
        uuid:               rowElems[9], // ID
      };
    }
  );
}
