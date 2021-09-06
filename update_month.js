/** @OnlyCurrentDoc */

// aggregate_year / aggregate_month のリストを更新する (これによって最新月/その前月として aggregation しに行く月の設定が切り替わる)
function executeUpdateAggregateYearMonthList(){
  postConsoleAndSlackJobStart("execute update aggregate year month list");
  try {
    doUpdateAggregateYearMonthList();
  } catch(error){
    handleError(error);
  }
  postConsoleAndSlackJobEnd("execute update aggregate year month list: updated");
}

function doUpdateAggregateYearMonthList(){
  // 現在の日付を設定
  const currentDate = getCurrent_yyyy_MM_dd();
  getTodayForCalenderRange().setValue(currentDate);

  // 本日が月初めの日か否かを判定
  if(getIsTodayMonthlyStart() != 1){
    postConsoleAndSlackJobEnd("execute update aggregate year month list: isTodayMonthlyStart: false: push skipped");
    return;
  }

  // 本日が月初めの場合: aggregate_year / aggregate_month のリストの末尾に本日の月日を追加する
  console.log("isTodayMonthlyStart: true");

  // リスト末尾のセル、およびその一つ後ろのセル (= push する対象のセル) を取得しておく
  const {aggregateMonthStack} = fetchAggregateYearMonths();
  const lastCellYear = getAggregateYearRange().getCell(aggregateMonthStack.length + 1, 1);
  const lastCellMonth = getAggregateMonthRange().getCell(aggregateMonthStack.length + 1, 1);
  const pushTargetCellYear = getAggregateYearRange().getCell(aggregateMonthStack.length + 2, 1);
  const pushTargetCellMonth = getAggregateMonthRange().getCell(aggregateMonthStack.length + 2, 1);

  // 本日の月日が追加未済であるかを確認する
  const currentYear = getCurrentYyyy();
  const currentMonth = getCurrentMm();
  const lastYearMonth = lastCellYear.getValue() + lastCellMonth.getValue();
  const currentYearMonth = currentYear + currentMonth;
  if(lastYearMonth == currentYearMonth){
    postConsoleAndSlackJobEnd("execute update aggregate year month list: currentYearMonth is already pushed: push skipped");
    return;
  }

  // 追加未済の場合、本日の月日を追加
  console.log("currentYearMonth is not pushed yet");

  pushTargetCellYear.setValue(currentYear);
  pushTargetCellMonth.setValue(currentMonth);
  console.log("aggregate_year, aggregate_month: pushed to the list: year, month: " + currentYear + ", " + currentMonth);
}
