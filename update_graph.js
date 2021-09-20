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
  const monthNum = 4;
  const targetYearMonths = getIntRangeFromZero(monthNum).map(
    i => {
      const monthsAgo = monthNum - i - 1;
      const {someMonthsAgoYyyy, someMonthsAgoMm} = getYearMonthMonthsAgoFrom(targetEndYear, targetEndMonth, monthsAgo);
      return {
        targetYear: someMonthsAgoYyyy,
        targetMonth: someMonthsAgoMm,
      };
    }
  );
  console.log("targetYearMonths: " + targetYearMonths.map(ym => Object.values(ym).join("/")));

  // 現預金残高の集計
  // 集計の始点となる月の末残高を取得する TODO: 空振り対応
  const dailyCashBalanceDetailsStartMonth = fetchDailyCashBalanceDetails(targetYearMonths[0].targetYear, targetYearMonths[0].targetMonth);
  const startCashBalance = dailyCashBalanceDetailsStartMonth[dailyCashBalanceDetailsStartMonth.length - 1].balance
  // 以降の各月の日付x残高リストを作成する
  let lastCashBalance = startCashBalance; // 最初の月の末残高をデフォルトの残高とする
  const graphSourceDailyCashBalances = targetYearMonths.slice(1).map( // 最初の月の翌月以降の月から各月分処理していく
    targetYearMonth => {
      // 該当月の日次現預金残高明細を取得する
      const targetMonthDailyCashBalances = fetchDailyCashBalanceDetails(targetYearMonth.targetYear, targetYearMonth.targetMonth);
      // 該当月初日から月末日まで各日分処理していく
      return getMonthlyDateArrayAt(targetYearMonth.targetYear, targetYearMonth.targetMonth).map(
        targetDate => {
          // 該当日の現預金残高明細を取得する
          const targetDateStr = getYyyyMmDdFrom(targetDate);
          const targetDateCashBalance = targetMonthDailyCashBalances.find(b => b.date == targetDateStr);

          // 該当日に明細が存在した場合はそれを適用し、最終残高を更新、そうでない場合は最終残高を適用し、オブジェクトにセットして返す
          if(targetDateCashBalance){
            lastCashBalance = targetDateCashBalance.balance;
            return targetDateCashBalance;
          } else {
            return {
              date: targetDateStr,
              balance: targetDate.getTime() > targetEndDate.getTime() ? "" : lastCashBalance,
            };
          }
        }
      );
    }
  ).flat(1);

  // TODO: WIP
  console.log(graphSourceDailyCashBalances);
}
