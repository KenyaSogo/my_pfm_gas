/** @OnlyCurrentDoc */

// 最新月分の月次集計を行う
function executeCalcMonthlySummaryCurrentMonth(){
  calcMonthlySummary(0);
}

// 最新月の前月分の月次集計を行う
function executeCalcMonthlySummaryPreviousMonth(){
  calcMonthlySummary(1);
}

// 指定月の月次集計を行う
function calcMonthlySummary(monthsAgo){
  // 集計対象月を取得する
  const {targetYear, targetMonth} = getTargetYearMonth(monthsAgo);
  console.log("targetYear, targetMonth: " + targetYear + ", " + targetMonth);

  // 項目別集計の月次集計を行う
  // 集計対象月の明細データを取得する
  const targetDetails = fetchMonthlySummaryByCategoryOriginDetails(targetYear, targetMonth);
  // 明細データから、項目を収集する TODO: sort
  const categories = getUniqueArrayFrom(targetDetails.map(d => d.category));
  // 項目別に集計する
  const summaryByCategories = categories.map(
    c => {
      const targetCategoryDetails = targetDetails.filter(d => d.category == c);
      const sumAmountByCategory = targetCategoryDetails
        .map(d => Number(d.amount))
        .reduce((sum, amount) => sum + amount, 0);
      return {
        yearMonth:    targetYear + targetMonth, // 対象年月
        categoryName: c,                        // 項目 (大項目&中項目)
        amount:       sumAmountByCategory,      // 金額
      };
    }
  );
  console.log(summaryByCategories);
}

// 集計対象月の項目別集計元データを取得する
function fetchMonthlySummaryByCategoryOriginDetails(targetYear, targetMonth){
  const dataOriginSheetName = getCalcDcExportSheetPrefix() + "_" + targetYear + targetMonth;
  const rawDetailData = getThisSpreadSheet().getSheetByName(dataOriginSheetName).getRange(getCalcDcExportAddr()).getValue();

  return rawDetailData.split("¥n").map(
    row => {
      const rowElems = row.split("#&#");
      return {
        date:     rowElems[0], // 日付
        category: rowElems[1], // 項目 (大項目>>中項目)
        amount:   rowElems[2], // 金額
      };
    }
  );
}

// 集計対象月を返す
function getTargetYearMonth(monthsAgo){
  const {someMonthsAgoYyyy, someMonthsAgoMm} = getYearMonthMonthsAgoFrom(getCurrentYyyy(), getCurrentMm(), monthsAgo);
  const targetYear =  someMonthsAgoYyyy;
  const targetMonth = someMonthsAgoMm;
  return {targetYear, targetMonth};
}
