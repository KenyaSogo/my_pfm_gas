/** @OnlyCurrentDoc */

// TODO: delete
function testCalcWeeklySummary(){
  calcWeeklySummary(0);
  calcWeeklySummary(1);
  calcWeeklySummary(2);
}

// 指定週の集計を行う
function calcWeeklySummary(weeksAgo){
  // 集計対象日を取得する
  const targetDates = getWeeklyCalcTargetDates(weeksAgo);
  console.log(`weeksAgo, targetDates: ${weeksAgo}, ${targetDates}`);

  // 項目別集計の週次集計を行う
  console.log("start weekly summary by category");

  // 集計対象の週の明細データを取得する
  const targetWeekDetails = getTargetWeekDetails(targetDates);

  console.log(targetWeekDetails);
}

// 集計対象の週の明細データを取得する
function getTargetWeekDetails(targetDates){
  const targetYearMonths = getUniqueArrayFrom(targetDates.map(d => getYyyyFrom(d) + "/" + getMmFrom(d)));
  const targetMonthDetails = targetYearMonths.map(
    ym => {
      const yms = ym.split("/");
      return fetchDailySummaryByCategoryDetails(yms[0], yms[1]);
    }
  ).flat(1);

  const targetYyyyMmDds = targetDates.map(d => getYyyyMmDdFrom(d));
  return targetMonthDetails.filter(d => targetYyyyMmDds.includes(d.date));
}

// 週次集計の集計対象日を返す
function getWeeklyCalcTargetDates(weeksAgo){
  const currentWeekdayIndex = getCurrentDate().getDay();
  const fromDate = getDateAgoFrom(getCurrentDate(), currentWeekdayIndex + weeksAgo * 7);
  return getIntRangeFromZero(7).map(i => getDateLaterFrom(fromDate, i));
}
