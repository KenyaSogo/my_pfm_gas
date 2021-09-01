/** @OnlyCurrentDoc */

// 月別のシートを作成する
function executeCreateSheetByMonth(){
  console.log("start: execute create sheet by month");
  postSlackLoggingChannel("start: execute create sheet by month");

  const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = currentSpreadSheet.getSheetByName("settings");

  // 現在の日付を設定
  const currentDate = getCurrent_yyyy_MM_dd();
  settingSheet.getRange("today_for_calender").setValue(currentDate);

  // 本日が月初めの日の前日か否かを判定
  const isTomorrowMonthlyStart = settingSheet.getRange("is_tomorrow_monthly_start").getValue() == 1;
  if(!isTomorrowMonthlyStart){
    console.log("isTomorrowMonthlyStart: false: sheet create skipped");
    postSlackLoggingChannel("end: execute create sheet by month: sheet create skipped");
    return;
  }

  // 本日が月初め日の前日の場合: 当月分のシートを tmp シートからコピーして追加
  console.log("isTomorrowMonthlyStart: true");

  getCreateTargetSheetConfig(settingSheet).forEach(
    sheetConfig => {
      const sheetPrefix = sheetConfig.createTargetSheetPrefix;
      const needsNextMonth = sheetConfig.needsNextMonth;
      console.log("start newSheet create: sheetPrefix, needsNextMonth: " + sheetPrefix + ", " + needsNextMonth);

      const currentYearMonth = getCurrentYyyy() + getCurrentMm();
      createSheetAtTargetMonth(currentYearMonth, sheetPrefix, currentSpreadSheet);

      // 翌月分も作る必要があるシートの場合、翌月分も作成する
      if(needsNextMonth == 1){
        const {nextMonthYyyy, nextMonthMm} = getNextYearMonth(getCurrentYyyy(), getCurrentMm());
        const nextYearMonth = nextMonthYyyy + nextMonthMm;
        createSheetAtTargetMonth(nextYearMonth, sheetPrefix, currentSpreadSheet);
      }
    }
  );

  console.log("end: execute create sheet by month");
  postSlackLoggingChannel("end: execute create sheet by month");
}

// 対象月のシートを作成する
function createSheetAtTargetMonth(targetYearMonth, sheetPrefix, targetSpreadSheet){
  console.log("start createSheetAtTargetMonth: targetYearMonth: " + targetYearMonth);

  const newSheetName = sheetPrefix + "_" + targetYearMonth;

  // シートが既にコピー済みか否かを判定
  const existingNewSheet = targetSpreadSheet.getSheetByName(newSheetName);
  if(existingNewSheet){
    console.log("newSheet is already created: sheet create skipped: newSheetName: " + newSheetName);
    return;
  }

  // コピー未済の場合、当月分のシートをコピーして追加
  console.log("newSheet is not created yet");

  const tmpSheet = targetSpreadSheet.getSheetByName(sheetPrefix + "_tmp");
  const copiedNewSheet = tmpSheet.copyTo(targetSpreadSheet);
  copiedNewSheet.setName(newSheetName);
  console.log("newSheet is created: sheet name: " + newSheetName);
}

// 作成対象シートの定義表を取得する
function getCreateTargetSheetConfig(settingSheet){
  const prefixes = getTrimmedColumnValues(settingSheet, "create_target_sheet_prefix");
  const needsNextMonths = getTrimmedColumnValues(settingSheet, "needs_next_month");
  return getIntRangeFromZero(prefixes.length).map(
    i => {
      return {
        createTargetSheetPrefix: prefixes[i],
        needsNextMonth: needsNextMonths[i],
      };
    }
  );
}
