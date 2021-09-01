/** @OnlyCurrentDoc */

// シートのメンテナンス状態をチェックする
function executeCheckSheetMaintenanceStatus(){
  console.log("start: execute check sheet maintenance status");
  postSlackLoggingChannel("start: execute check sheet maintenance status");

  const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("settings");

  // 休日カレンダーのメンテナンス要否を確認する
  // (現状数年分作っているが、時期が来て足さないといけなくなったタイミングで通知を飛ばす ※祝日のフラグを手入力する以外は関数の拡張のみ)
  const isCalenderNeedsMaintenance = settingSheet.getRange("is_calender_needs_maitenance").getValue() == 1;
  if(isCalenderNeedsMaintenance){
    console.log("isCalenderNeedsMaintenance: true: notified to slack");
    postSlackLoggingChannel("is calender needs maintenance: true"); // TODO: slack 通知: メンション
  } else {
    console.log("isCalenderNeedsMaintenance: false");
  }

  // 月初めの日 (基本だが25日、その日が休日だと営業日ベースに前倒しさせた日) の設定が正しいか確認する
  // (現状、6連休の最後の日が25日、というパターンまでしか考慮しておらず、その限界を超えるケースの有無をチェックする)
  const isMonthlyStartDayInvalid = settingSheet.getRange("is_monthly_start_day_invalid").getValue() == 1;
  if(isMonthlyStartDayInvalid){
    console.log("isMonthlyStartDayInvalid: true: notified to slack");
    postSlackLoggingChannel("is monthly start day invalid: true"); // TODO: slack 通知: メンション
  } else {
    console.log("isMonthlyStartDayInvalid: false");
  }

  console.log("end: execute check sheet maintenance status");
  postSlackLoggingChannel("end: execute check sheet maintenance status");
}
