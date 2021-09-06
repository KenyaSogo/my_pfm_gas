/** @OnlyCurrentDoc */

// 月別のシートを作成する
function executeCreateSheetByMonth(){
  postConsoleAndSlackJobStart("execute create sheet by month");
  try {
    doCreateSheetByMonth();
  } catch(error){
    handleError(error);
  }
  postConsoleAndSlackJobEnd("execute create sheet by month");
}

function doCreateSheetByMonth(){
  // 現在の日付を設定
  const currentDate = getCurrent_yyyy_MM_dd();
  getTodayForCalenderRange().setValue(currentDate);

  // 本日が月初めの日の前日か否かを判定
  if(getIsTomorrowMonthlyStart() != 1){
    postConsoleAndSlackJobEnd("execute create sheet by month: isTomorrowMonthlyStart: false: sheet create skipped");
    return;
  }

  // 本日が月初め日の前日の場合: 当月分のシートを tmp シートからコピーして追加
  console.log("isTomorrowMonthlyStart: true");

  getCreateTargetSheetConfigs().forEach(
    sheetConfig => {
      const sheetPrefix = sheetConfig.createTargetSheetPrefix;
      const needsNextMonth = sheetConfig.needsNextMonth;
      console.log("start newSheet create: sheetPrefix, needsNextMonth: " + sheetPrefix + ", " + needsNextMonth);

      const currentYearMonth = getCurrentYyyy() + getCurrentMm();
      createSheetAtTargetMonth(currentYearMonth, sheetPrefix);

      // 翌月分も作る必要があるシートの場合、翌月分も作成する
      if(needsNextMonth == 1){
        const {nextMonthYyyy, nextMonthMm} = getNextYearMonth(getCurrentYyyy(), getCurrentMm());
        const nextYearMonth = nextMonthYyyy + nextMonthMm;
        createSheetAtTargetMonth(nextYearMonth, sheetPrefix);
      }
    }
  );
}

// 対象月のシートを作成する
function createSheetAtTargetMonth(targetYearMonth, sheetPrefix){
  console.log("start createSheetAtTargetMonth: targetYearMonth: " + targetYearMonth);

  const newSheetName = sheetPrefix + "_" + targetYearMonth;

  // シートが既にコピー済みか否かを判定
  const existingNewSheet = getThisSpreadSheet().getSheetByName(newSheetName);
  if(existingNewSheet){
    console.log("newSheet is already created: sheet create skipped: newSheetName: " + newSheetName);
    return;
  }

  // コピー未済の場合、当月分のシートをコピーして追加
  console.log("newSheet is not created yet");

  const tmpSheet = getThisSpreadSheet().getSheetByName(sheetPrefix + "_" + getTemplateSheetSuffix());
  const copiedNewSheet = tmpSheet.copyTo(getThisSpreadSheet());
  copiedNewSheet.setName(newSheetName);
  console.log("newSheet is created: sheet name: " + newSheetName);
}
