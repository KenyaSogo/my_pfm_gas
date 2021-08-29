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
  // 明細データから項目を収集する
  const targetCategories = getUniqueArrayFrom(targetDetails.map(d => d.category));
  // 項目別に明細を集計する
  const summaryByCategories = targetCategories.map(
    c => {
      const sumAmountByCategory = sumAmountFromDetails(targetDetails.filter(d => d.category == c));
      return getMonthlySummaryFromCategory(targetYear, targetMonth, c, sumAmountByCategory);
    }
  );
  // さらに、項目を定められた順番に整えつつ再集計 (= 小計、合計、総合計の集計) する
  const reSummaryByCategories = getOrderedCategoryConfigs().map(
    c => {
      // 収支合計の場合
      if(c.middleCategoryName == "収支合計"){
        // 全ての明細を集計する
        const sumAmountByCategory = sumAmountFromDetails(targetDetails)
        return getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, c, sumAmountByCategory);
      }
      // 総合計の集計項目の場合
      if(c.middleCategoryName == "総合計"){
        // config に未定義の項目が存在し得ることを考えて、それらを漏らさないよう金額の正負を元に集計する
        if(c.largeCategoryName == "収入"){
          // 金額が正の明細を集計する
          const sumAmountByCategory = sumAmountFromDetails(targetDetails.filter(d => Number(d.amount) > 0))
          return getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, c, sumAmountByCategory);
        } else if(c.largeCategoryName == "支出"){
          // 金額が負の明細を集計する
          const sumAmountByCategory = sumAmountFromDetails(targetDetails.filter(d => Number(d.amount) < 0))
          return getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, c, sumAmountByCategory);
        }
      }
      // TODO: 続き
    }
  );
  console.log(reSummaryByCategories);
}

// 項目定義 (CategoryConfig) を元に集計結果明細を作成し返す
function getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, config, amount){
  return getMonthlySummary(
    targetYear,
    targetMonth,
    config.largeCategoryName,
    config.middleCategoryName,
    amount);
}

// 項目 (大項目>>中項目) を元に集計結果明細を作成し返す
function getMonthlySummaryFromCategory(targetYear, targetMonth, category, amount){
  return getMonthlySummary(
    targetYear,
    targetMonth,
    category.split(">>")[0],
    category.split(">>")[1],
    amount);
}

// 集計結果明細を返す
function getMonthlySummary(targetYear, targetMonth, largeCategoryName, middleCategoryName, amount){
  return {
    yearMonth:          targetYear + targetMonth, // 対象年月
    largeCategoryName:  largeCategoryName,        // 大項目
    middleCategoryName: middleCategoryName,       // 中項目
    amount:             amount,                   // 金額
  };
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
  const targetYear = someMonthsAgoYyyy;
  const targetMonth = someMonthsAgoMm;
  return {targetYear, targetMonth};
}
