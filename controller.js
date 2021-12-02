/** @OnlyCurrentDoc */

function executeControl(){
  const processStartTime = Date.now();
  console.log("start executeControl");

  try {
    doControl();
  } catch(error){
    handleError(error);
  }

  // トータルの処理時間を出力
  const processingTime = Date.now() - processStartTime;
  postConsoleAndSlackInfoMessage("execute escalation was finished (" + processingTime/1000 + " sec)");
  // 処理時間がトータル 300 (60*5) 秒を超えた場合には warning を出しておく
  if(processingTime > 300000){
    postConsoleAndSlackWarning("total job processing time exceeded 5 min", true);
  }
  console.log("end executeControl");
}

function doControl(){
  // exec group の trigger 判定用の現在時刻セルに値をセットする
  const currentTime = getCurrent_yyyy_mm_dd_hh_mm_ss();
  getCurrentTimeForExecTriggerRange().setValue(currentTime);

  // trigger が on になっている exec group を起動する
  getExecGroupTriggerFlags().filter(t => t.isTriggerOn == 1).forEach(
    trigger => {
      switch(trigger.execGroupName){
        case execGroupNameCreateSheet:
          executeCreateSheetByMonth();
          break;
        case execGroupNameUpdateMonth:
          executeUpdateAggregateYearMonthList();
          break;
        case execGroupNameAggrePrev:
          if(!executeAggregatePreviousMonth()){
            postConsoleAndSlackInfoMessage("aggregate result has no update: following exec was skipped");
            break;
          }
          executeCalcDailySummaryPreviousMonth();
          executeCalcMonthlySummaryPreviousMonth();
          executeCalcWeeklySummaryLast3Weeks();
          executeUpdateGraphSourceData();
          break;
        case execGroupNameAggreCurr:
          if(!executeAggregateCurrentMonth()){
            postConsoleAndSlackInfoMessage("aggregate result has no update: following exec was skipped");
            break;
          }
          executeCalcDailySummaryCurrentMonth();
          executeCalcMonthlySummaryCurrentMonth();
          executeCalcWeeklySummaryLast3Weeks();
          executeUpdateGraphSourceData();
          break;
        case execGroupNameReport:
          executeReport();
          break;
        case execGroupNameCheckMaintenance:
          executeCheckSheetMaintenanceStatus();
          break;
      }
    }
  );
}
