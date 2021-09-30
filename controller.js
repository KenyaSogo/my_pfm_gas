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

  // 各 exec group を activate trigger に従って処理する
  getExecGroupTriggerFlags().forEach(
    flag => {
      if(flag.isTriggerOn != 1){
        return;
      }
      switch(flag.execGroupName){
        case execGroupNameCreateSheet:
          executeCreateSheetByMonth();
          break;
        case execGroupNameUpdateMonth:
          executeUpdateAggregateYearMonthList();
          break;
        case execGroupNameAggrePrev:
          executeAggregatePreviousMonth();
          executeCalcDailySummaryPreviousMonth(); // TODO: update なしの場合に skip
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
