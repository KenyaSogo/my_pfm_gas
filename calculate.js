/** @OnlyCurrentDoc */

// 最新月分の日次集計を行う
function executeCalcDailySummaryCurrentMonth(){
  calcDailySummary(0);
}

// 最新月の前月分の日次集計を行う
function executeCalcDailySummaryPreviousMonth(){
  calcDailySummary(1);
}

// 指定月の日次集計を行う
function calcDailySummary(monthsAgo){
  const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("settings");

  // 集計対象月を取得する
  const {targetYear, targetMonth} = getAggregateTargetYearAndMonth(settingSheet, monthsAgo);
  console.log("targetYear, targetMonth: " + targetYear + ", " + targetMonth);

  // 集計対象月の明細データを取得する
  const aggregatedDetails = fetchAggregatedDetails(targetYear, targetMonth);
  console.log("fetch and parse aggregatedDetails finished: length: " + aggregatedDetails.length);

  // 日次現預金残高集計を行い、結果シートに出力する
  calcAndExportDailyBalanceCash(settingSheet, aggregatedDetails, targetYear, targetMonth);

  // 日次総資産残高集計を行い、結果シートに出力する
  calcAndExportDailyBalanceTotalAsset(settingSheet, aggregatedDetails, targetYear, targetMonth);

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
  pasteDailyCalcResultByEachMonth(targetYear, targetMonth, summaryByCategoryAndDates, "calc_dc");
  console.log("end calcAndExportDailySummaryByCategory");
}

// 日次総資産残高集計を行い、結果シートに出力する
function calcAndExportDailyBalanceTotalAsset(settingSheet, aggregatedDetails, targetYear, targetMonth){
  // 明細データのうち全金融機関分の入出金を集計し、日次の全資産分の残高推移を求める
  console.log("start calcAndExportDailyBalanceTotalAsset");
  // 集計対象明細を取得
  const balanceTotalAssetTargetDetails = getBalanceTotalAssetTargetDetails(aggregatedDetails);
  // 集計対象の日付を収集
  const balanceTotalAssetTargetDates = extractSortedDatesFromDetails(balanceTotalAssetTargetDetails);
  // 前月末残高を取得する
  const previousMonthEndBalanceTotalAsset = Number(getPreviousMonthEndBalance(settingSheet, targetYear, targetMonth).assetTotalEndBalance);
  // 日次で総資産残高を集計
  const totalAssetBalanceByDates = accumulateBalanceByDate(previousMonthEndBalanceTotalAsset, balanceTotalAssetTargetDates, balanceTotalAssetTargetDetails);
  console.log("calc totalAssetBalanceByDates finished: length: " + totalAssetBalanceByDates.length);
  // 対象月の末残高設定を更新する
  const updatedTotalAssetBalanceAtMonthEnd = updateTargetMonthEndBalance(totalAssetBalanceByDates, previousMonthEndBalanceTotalAsset, settingSheet, targetYear, targetMonth, "asset_total_end_balance");
  console.log("total asset end balance updated: updatedTotalAssetBalanceAtMonthEnd: " + updatedTotalAssetBalanceAtMonthEnd);
  // 集計結果明細を出力対象シートに貼り付ける
  pasteDailyCalcResultByEachMonth(targetYear, targetMonth, totalAssetBalanceByDates, "calc_dab");
  console.log("end calcAndExportDailyBalanceTotalAsset");
}

// 日次現預金残高集計を行い、結果シートに出力する
function calcAndExportDailyBalanceCash(settingSheet, aggregatedDetails, targetYear, targetMonth){
  // 明細データのうち現預金の入出金を集計し、日次の残高推移を求める
  console.log("start calcAndExportDailyBalanceCash");
  // 集計対象明細を取得
  const balanceCashTargetDetails = getBalanceCashTargetDetails(settingSheet, aggregatedDetails);
  // 集計対象の日付を収集
  const balanceCashTargetDates = extractSortedDatesFromDetails(balanceCashTargetDetails);
  // 前月末残高を取得する
  const previousMonthEndBalanceCash = Number(getPreviousMonthEndBalance(settingSheet, targetYear, targetMonth).cashEndBalance);
  // 日次で現預金残高を集計
  const cashBalanceByDates = accumulateBalanceByDate(previousMonthEndBalanceCash, balanceCashTargetDates, balanceCashTargetDetails);
  console.log("calc cashBalanceByDates finished: length: " + cashBalanceByDates.length);
  // 対象月の末残高設定を更新する
  const updatedCashBalanceAtMonthEnd = updateTargetMonthEndBalance(cashBalanceByDates, previousMonthEndBalanceCash, settingSheet, targetYear, targetMonth, "cash_end_balance");
  console.log("cash end balance updated: updatedCashBalanceAtMonthEnd: " + updatedCashBalanceAtMonthEnd);
  // 集計結果明細を出力対象シートに貼り付ける
  pasteDailyCalcResultByEachMonth(targetYear, targetMonth, cashBalanceByDates, "calc_dcb");
  console.log("end calcAndExportDailyBalanceCash");
}

// 日次集計結果を、早い月分と遅い月分に振り分け、それぞれを結果シートに貼り付ける
function pasteDailyCalcResultByEachMonth(targetYear, targetMonth, dailyCalcResults, targetSheetPrefix){
  // 集計した明細を、早い月分と遅い月分に振り分ける
  const earlierYearMonth = targetYear + "/" + targetMonth;
  const earlierMonthDailyCalcResults = dailyCalcResults.filter(s => s.date.indexOf(earlierYearMonth) == 0);
  const laterMonthDailyCalcResults = dailyCalcResults.filter(s => s.date.indexOf(earlierYearMonth) != 0);

  // 集計結果を各月分の結果シートに貼り付け
  console.log("start to paste dailyCalcResultByEachMonth: targetSheetPrefix: " + targetSheetPrefix);
  // 早い月分
  if(earlierMonthDailyCalcResults.length > 0){
    console.log("start to paste earlierMonthDailyCalcResults: earlierMonthDailyCalcResults.length: " + earlierMonthDailyCalcResults.length);
    pasteDailyCalcResult(earlierMonthDailyCalcResults, targetYear, targetMonth, true, targetSheetPrefix);
  } else {
    console.log("pasting earlierMonthDailyCalcResults was skipped");
  }
  // 遅い月分
  if(laterMonthDailyCalcResults.length > 0){
    console.log("start to paste laterMonthDailyCalcResults: laterMonthDailyCalcResults.length: " + laterMonthDailyCalcResults.length);
    const {nextMonthYyyy, nextMonthMm} = getNextYearMonth(targetYear, targetMonth);
    pasteDailyCalcResult(laterMonthDailyCalcResults, nextMonthYyyy, nextMonthMm, false, targetSheetPrefix);
  } else {
    console.log("pasting laterMonthDailyCalcResults was skipped");
  }
}

// 対象月の末残高レコードを更新する
function updateTargetMonthEndBalance(accumulatedBalanceByDates, previousMonthEndBalance, settingSheet, targetYear, targetMonth, targetRangeName){
  // 対象月の末残高を取得
  const balanceAtMonthEnd = accumulatedBalanceByDates.length == 0
    ? previousMonthEndBalance
    : accumulatedBalanceByDates[accumulatedBalanceByDates.length - 1].balance;
  // 対象月の末残高表における index を取得
  const targetMonthIndexAtEndBalances = fetchEndBalances(settingSheet)
    .map(b => b.calcYear + "/" + b.calcMonth)
    .indexOf(targetYear + "/" + targetMonth);
  // 該当 cell を取得し、更新する
  settingSheet.getRange(targetRangeName)
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
          const targetDateAndCategoryDetails = targetDateDetails.filter(d => d.largeCategoryName + ">>" + d.middleCategoryName == c);
          const sumAmountByDateAndCategory = targetDateAndCategoryDetails
            .map(d => Number(d.amount))
            .reduce((sum, amount) => sum + amount, 0);
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
      // 対象日の明細を抽出
      const targetDateDetails = targetDetails.filter(d => d.date == date);
      // 残高を集計
      const sumAmountAtTargetDate = targetDateDetails
        .map(d => Number(d.amount))
        .reduce((sum, amount) => sum + amount, 0);
      accumulatedBalance += sumAmountAtTargetDate;

      return {
        date:    date,
        balance: accumulatedBalance,
      };
    }
  );
}

// 前月末残高を返す
function getPreviousMonthEndBalance(settingSheet, targetYear, targetMonth){
  const {previousMonthYyyy, previousMonthMm} = getPreviousYearMonth(targetYear, targetMonth);
  const endBalances = fetchEndBalances(settingSheet);

  return endBalances.find(b => b.calcYear == previousMonthYyyy && b.calcMonth == previousMonthMm);
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
function getBalanceCashTargetDetails(settingSheet, aggregatedDetails){
  // 現預金に該当する保有金融機関名を取得する
  const assetConfigs = getAssetConfigs(settingSheet);
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
  const targetAggreSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ag_" + targetYear + targetMonth);
  const allRowsData = targetAggreSheet.getRange("V2").getValue(); // TODO: 列アドレスの直書きを回避

  return allRowsData.split("¥n").map(
    row => {
      const rowElems = row.split("#&#");
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

// 保有金融機関表を取得して返す
let assetConfigs;
function getAssetConfigs(settingSheet){
  if(assetConfigs){
    return assetConfigs;
  }

  const assetNameValues = getTrimmedColumnValues(settingSheet, "asset_name");
  const isCashValues = getTrimmedColumnValues(settingSheet, "is_cash");
  assetConfigs = getIntRangeFromZero(assetNameValues.length).map(
    i => {
      return {
        assetName: assetNameValues[i],
        isCash:    isCashValues[i],
      };
    }
  );

  return assetConfigs;
}

// 月末残高表を取得して返す
let endBalances;
function fetchEndBalances(settingSheet){
  if(endBalances){
    return endBalances;
  }

  const calcYearValues = getTrimmedColumnValues(settingSheet, "calc_year");
  const calcMonthValues = getTrimmedColumnValues(settingSheet, "calc_month");
  const cashEndBalanceValues = getTrimmedColumnValues(settingSheet, "cash_end_balance");
  const assetTotalEndBalanceValues = getTrimmedColumnValues(settingSheet, "asset_total_end_balance");
  endBalances = getIntRangeFromZero(calcYearValues.length).map(
    i => {
      return {
        calcYear:             calcYearValues[i],
        calcMonth:            calcMonthValues[i],
        cashEndBalance:       cashEndBalanceValues[i],
        assetTotalEndBalance: assetTotalEndBalanceValues[i],
      };
    }
  );

  return endBalances;
}

// 日次集計結果の calc シートへの貼り付けを行う
function pasteDailyCalcResult(summaryByDates, targetYear, targetMonth, isEarlier, sheetPrefix){
  console.log("pasteDailyCalcResult: start: targetYear, targetMonth, isEarlier: " + [targetYear, targetMonth, isEarlier].join(", "));

  // 集計結果の明細を区切り文字で結合し、貼り付け用に一つの文字列にする
  const mergedSummaryByDatesValue = summaryByDates.map(s => Object.values(s).join("#&#")).join("¥n");

  // 貼り付け対象のセルを取得
  const targetCalcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetPrefix + "_" + targetYear + targetMonth); // TODO: sheet 作成 job に追加
  const pasteTargetCell = targetCalcSheet.getRange(isEarlier ? "A3" : "A2"); // TODO: アドレス直書きの廃止

  // 対象データにつき、更新がなければ、貼り付けをスキップして終了
  const currentValue = pasteTargetCell.getValue();
  if(mergedSummaryByDatesValue == currentValue){
    console.log("calc result has no update: update skipped");
    return;
  }

  // データを対象セルに貼り付ける
  pasteTargetCell.setValue(mergedSummaryByDatesValue);
  console.log("calc result was updated");
}
