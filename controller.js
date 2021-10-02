/** @OnlyCurrentDoc */

function executeControl(){
  console.log("start executeControl");
  try {
    doControl();
  } catch(error){
    handleError(error);
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
      let isUpdatedAggregateResult = false;
      switch(trigger.execGroupName){
        case execGroupNameCreateSheet:
          executeCreateSheetByMonth();
          break;
        case execGroupNameUpdateMonth:
          executeUpdateAggregateYearMonthList();
          break;
        case execGroupNameAggrePrev:
          isUpdatedAggregateResult = executeAggregatePreviousMonth();
          if(!isUpdatedAggregateResult){
            postConsoleAndSlackInfoMessage("aggregate result has no update: following exec was skipped");
            break;
          }
          executeCalcDailySummaryPreviousMonth();
          executeCalcMonthlySummaryPreviousMonth();
          executeUpdateGraphSourceData();
          break;
        case execGroupNameAggreCurr:
          isUpdatedAggregateResult = executeAggregateCurrentMonth();
          if(!isUpdatedAggregateResult){
            postConsoleAndSlackInfoMessage("aggregate result has no update: following exec was skipped");
            break;
          }
          executeCalcDailySummaryCurrentMonth();
          executeCalcMonthlySummaryCurrentMonth();
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
