/** @OnlyCurrentDoc */

// グラフデータの更新を行う
function executeUpdateGraphSourceData(){
  postConsoleAndSlackJobStart("execute update graph source data");
  try {
    doUpdateGraphSourceData();
  } catch(error){
    handleError(error);
  }
  postConsoleAndSlackJobEnd("execute update graph source data");
}

function doUpdateGraphSourceData(){
  // グラフの終点となる月を取得する
  const targetEndYear = getCurrentYyyy();
  const targetEndMonth = getCurrentMm();
  const targetEndDate = getCurrentDate();
  // 対象月のリストを作成する
  const targetYearMonths = getTargetYearMonthListFromEnd(targetEndYear, targetEndMonth, 4)
  console.log("targetYearMonths: " + targetYearMonths.map(ym => Object.values(ym).join("/")));

  // 現預金残高明細を積み上げ直す
  const graphSourceDailyCashBalances = accumulateGraphSourceBalances(targetYearMonths, fetchDailyCashBalanceDetails, targetEndDate)
  // 現預金残高の集計結果明細を対象シートに export する
  exportResultDetails(graphSourceDailyCashBalances, getGraphDcbImportSheetName(), getGraphDcbImportAddr());
}

// 残高明細をグラフ用に積み上げ直したものを返す
// (現状の日次残高明細は、残高が変動した日の分しか明細が存在しない。
//  が、グラフ描画用にはその間の日の分も埋めておく必要がある。
//  この関数では、間の日分も前日の残高を横置きする形で埋めた明細を作成して返す)
function accumulateGraphSourceBalances(targetYearMonths, fetchDailyBalanceDetailsFunc, targetEndDate){
  // 集計の始点となる月の末残高を取得する TODO: 空振り対応
  const startBalance = getLastElemFrom(fetchDailyBalanceDetailsFunc(targetYearMonths[0].targetYear, targetYearMonths[0].targetMonth), null).balance;
  // 以降の各月の日付x残高リストを作成する
  let lastBalance = startBalance; // 最初の月の末残高をデフォルトの残高とする
  return targetYearMonths.slice(1).map( // 最初の月の翌月以降の月から各月分処理していく
    targetYearMonth => {
      // 該当月の日次残高明細を取得する
      const targetMonthDailyBalances = fetchDailyBalanceDetailsFunc(targetYearMonth.targetYear, targetYearMonth.targetMonth);
      // 該当月初日から月末日まで各日分処理していく
      return getMonthlyDateArrayAt(targetYearMonth.targetYear, targetYearMonth.targetMonth).map(
        targetDate => {
          // 該当日の現預金残高明細を取得する
          const targetDateStr = getYyyyMmDdFrom(targetDate);
          const targetDateBalance = targetMonthDailyBalances.find(b => b.date == targetDateStr);

          // 該当日に明細が存在した場合はそれを適用し、最終残高を更新、そうでない場合は最終残高を適用し、オブジェクトにセットして返す
          if(targetDateBalance){
            lastBalance = targetDateBalance.balance;
            return targetDateBalance;
          } else {
            return {
              date: targetDateStr,
              balance: targetDate.getTime() > targetEndDate.getTime() ? "" : lastBalance,
            };
          }
        }
      );
    }
  ).flat(1); // 二次元配列になっているので、平坦化しておく
}

// 対象月のリストを最終月を指定して作成し返す
function getTargetYearMonthListFromEnd(targetEndYear, targetEndMonth, monthNum){
  return getIntRangeFromZero(monthNum).map(
    i => {
      const {someMonthsAgoYyyy, someMonthsAgoMm} = getYearMonthMonthsAgoFrom(targetEndYear, targetEndMonth, monthNum - i - 1);
      return {
        targetYear: someMonthsAgoYyyy,
        targetMonth: someMonthsAgoMm,
      };
    }
  );
}
