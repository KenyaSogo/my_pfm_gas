/** @OnlyCurrentDoc */

// シートのメンテナンス状態をチェックする
function executeCheckSheetMaintenanceStatus(){
  postConsoleAndSlackJobStart("execute check sheet maintenance status");
  try {
    doCheckSheetMaintenanceStatus();
  } catch(error){
    handleError(error);
  }
  postConsoleAndSlackJobEnd("execute check sheet maintenance status");
}

function doCheckSheetMaintenanceStatus(){
  // 休日カレンダーのメンテナンス要否を確認する
  // (現状数年分作っているが、時期が来て足さないといけなくなったタイミングで通知を飛ばす ※祝日のフラグを手入力する以外は関数の拡張のみ)
  const isCalenderNeedsMaintenance = getSettingSheet().getRange("is_calender_needs_maitenance").getValue() == 1; // TODO: typo
  if(isCalenderNeedsMaintenance){
    postConsoleAndSlackWarning("isCalenderNeedsMaintenance: true", true);
  } else {
    console.log("isCalenderNeedsMaintenance: false");
  }

  // 月初めの日 (基本だが25日、その日が休日だと営業日ベースに前倒しさせた日) の設定が正しいか確認する
  // (現状、6連休の最後の日が25日、というパターンまでしか考慮しておらず、その限界を超えるケースの有無をチェックする)
  const isMonthlyStartDayInvalid = getSettingSheet().getRange("is_monthly_start_day_invalid").getValue() == 1;
  if(isMonthlyStartDayInvalid){
    postConsoleAndSlackWarning("isMonthlyStartDayInvalid: true", true);
  } else {
    console.log("isMonthlyStartDayInvalid: false");
  }
}
