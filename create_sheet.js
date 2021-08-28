/** @OnlyCurrentDoc */

// 月別のシートを作成する
function executeCreateSheetByMonth(){
  const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = currentSpreadSheet.getSheetByName("settings");

  // 現在の日付を設定
  const currentDate = getCurrent_yyyy_MM_dd();
  settingSheet.getRange("today_for_calender").setValue(currentDate);

  // 本日が月初めの日の前日か否かを判定
  const isTomorrowMonthlyStart = settingSheet.getRange("is_tomorrow_monthly_start").getValue() == 1;
  if(!isTomorrowMonthlyStart){
    console.log("isTomorrowMonthlyStart: false: sheet create skipped");
    return;
  }

  // 本日が月初め日の前日の場合: 当月分のシートを tmp シートからコピーして追加
  console.log("isTomorrowMonthlyStart: true");

  getTrimmedColumnValues(settingSheet, "create_target_sheet_prefix").forEach(
    sheetPrefix => {
      console.log("start newSheet create: sheetPrefix: " + sheetPrefix);

      // シートが既にコピー済みか否かを判定
      const currentYearMonth = getCurrentYyyy() + getCurrentMm();
      const newSheetName = sheetPrefix + "_" + currentYearMonth;
      const existingNewSheet = currentSpreadSheet.getSheetByName(newSheetName);
      if(existingNewSheet){
        console.log("newSheet is already created: sheet create skipped: newSheetName: " + newSheetName);
        return;
      }

      // コピー未済の場合、当月分のシートをコピーして追加
      console.log("newSheet is not created yet");

      const tmpSheet = currentSpreadSheet.getSheetByName(sheetPrefix + "_tmp");
      const copiedNewSheet = tmpSheet.copyTo(currentSpreadSheet);
      copiedNewSheet.setName(newSheetName);
      console.log("newSheet is created: sheet name: " + newSheetName);
    }
  );
}
