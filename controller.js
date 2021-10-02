/** @OnlyCurrentDoc */

function executeControl(){
  console.log("start executeControl");
  doControl();
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
          executeAggregatePreviousMonth(); // TODO: update なしの場合に skip
          executeCalcDailySummaryPreviousMonth();
          executeCalcMonthlySummaryPreviousMonth();
          executeUpdateGraphSourceData();
          break;
        case execGroupNameAggreCurr:
          executeAggregateCurrentMonth();
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
