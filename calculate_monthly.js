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
  console.log("start monthly summary by category");
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
  // まずは、大項目別に小計を集計する
  const reSummaryByLargeCategories = getOrderedCategoryConfigs().map(
    c => {
      // 今回は集計対象外とする項目について skip
      if(["収支合計", "総合計", "合計", "未定義(収入)", "未定義(支出)", "未定義(中項目)"].includes(c.middleCategoryName)){
        return null;
      }
      // 小計項目の場合、大項目別に集計する
      if(c.middleCategoryName == "小計"){
        const sumAmountByCategory = sumAmountFromDetails(summaryByCategories.filter(s => s.largeCategoryName == c.largeCategoryName));
        return getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, c, sumAmountByCategory);
      }
      // その他の項目の場合、項目別集計結果をそのままマッピングする
      return getMatchedSummaryByCategory(c, summaryByCategories, targetYear, targetMonth);
    }
  ).filter(e => e);
  // 大項目内に config で未定義の中項目がある場合を考慮して、未定義(中項目) を集計
  const reSummaryWithUndefinedMiddleCategories = getOrderedCategoryConfigs().map(
    c => {
      // 今回は集計対象外とする項目について skip
      if(["収支合計", "総合計", "合計", "未定義(収入)", "未定義(支出)"].includes(c.middleCategoryName)){
        return null;
      }
      // 未定義(中項目)項目の場合、小計から定義済み項目の合計値を差し引いて、金額を求める
      if(c.middleCategoryName == "未定義(中項目)"){
        // 該当大項目の小計を取得
        const sumAmountByLargeCategory = reSummaryByLargeCategories
          .find(s => s.largeCategoryName == c.largeCategoryName && s.middleCategoryName == "小計")
          .amount
        // 該当大項目の定義済み中項目の合計値を求める
        const sumAmountByLargeCategoryDefined = sumAmountFromDetails(
          reSummaryByLargeCategories
            .filter(s => s.largeCategoryName == c.largeCategoryName && s.middleCategoryName != "小計"));
        // 上記の差額から未定義(中項目)の金額を求める
        const sumAmountByLargeCategoryUndefined = sumAmountByLargeCategory - sumAmountByLargeCategoryDefined;
        // 結果明細にマップして返す (0円の場合は skip する)
        if(sumAmountByLargeCategoryUndefined == 0){
          return null;
        }
        return getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, c, sumAmountByLargeCategoryUndefined);
      }
      // その他の項目の場合、項目別集計結果をそのままマッピングする
      return getMatchedSummaryByCategory(c, reSummaryByLargeCategories, targetYear, targetMonth);
    }
  ).filter(e => e);
  // 合計 (変動費、固定費等の粒度) を集計する
  const reSummaryByCategoryMiddleClasses = getOrderedCategoryConfigs().map(
    c => {
      // 今回は集計対象外とする項目について skip
      if(["収支合計", "総合計", "未定義(収入)", "未定義(支出)"].includes(c.middleCategoryName)){
        return null;
      }
      // 合計項目の場合、紐付く小計項目を集計する
      if(c.middleCategoryName == "合計"){
        // 集計対象の項目を収集する
        const categoriesInTargetMiddleClass = getOrderedCategoryConfigs()
          .filter(r => r.categoryMiddleClassName == c.largeCategoryName)
          .map(r => r.largeCategoryName + ">>" + r.middleCategoryName);
        const sumAmountByCategoryMiddleClass = sumAmountFromDetails(reSummaryWithUndefinedMiddleCategories
          .filter(s => categoriesInTargetMiddleClass.includes(s.largeCategoryName + ">>" + s.middleCategoryName)));
        return getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, c, sumAmountByCategoryMiddleClass);
      }
      // その他の項目の場合、項目別集計結果をそのままマッピングする
      return getMatchedSummaryByCategory(c, reSummaryWithUndefinedMiddleCategories, targetYear, targetMonth);
    }
  ).filter(e => e);
  // TODO: 残項目の集計, 検算 -> error (収支合計の突合), export
  console.log(reSummaryByCategoryMiddleClasses);
  console.log("end monthly summary by category");
}

// config の項目が一致する既存集計結果をマッピングして返す
function getMatchedSummaryByCategory(categoryConfig, summaryByCategories, targetYear, targetMonth){
  const matchedSummaryByCategory = summaryByCategories.find(s => s.largeCategoryName == categoryConfig.largeCategoryName && s.middleCategoryName == categoryConfig.middleCategoryName);
  if(matchedSummaryByCategory){
    return getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, categoryConfig, matchedSummaryByCategory.amount);
  }
  return null;
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
