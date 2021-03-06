/** @OnlyCurrentDoc */

// 最新月分の月次集計を行う
function executeCalcMonthlySummaryCurrentMonth(){
  postConsoleAndSlackJobStart("execute calc monthly summary current month");
  try {
    calcMonthlySummary(0);
  } catch(error){
    handleExpectedError(error);
  }
  postConsoleAndSlackJobEnd("execute calc monthly summary current month");
}

// 最新月の前月分の月次集計を行う
function executeCalcMonthlySummaryPreviousMonth(){
  postConsoleAndSlackJobStart("execute calc monthly summary previous month");
  try {
    calcMonthlySummary(1);
  } catch(error){
    handleExpectedError(error);
  }
  postConsoleAndSlackJobEnd("execute calc monthly summary previous month");
}

// 指定月の月次集計を行う
function calcMonthlySummary(monthsAgo){
  // 集計対象月を取得する
  const {targetYear, targetMonth} = getMonthlyCalcTargetYearMonth(monthsAgo);
  console.log("targetYear, targetMonth: " + targetYear + ", " + targetMonth);

  // 項目別集計の月次集計を行う
  console.log("start monthly summary by category");

  // 集計対象月の明細について、項目別に集計する
  const summaryByCategories = calcTargetMonthSummaryByCategories(targetYear, targetMonth);
  if(summaryByCategories == null){
    console.log("end monthly summary by category: skip summary and export: there is no data");
    return;
  }

  // 項目を定められた順番に整えつつ再集計 (= 小計、合計、総合計の集計) する
  // 大項目別に小計を集計する
  const reSummaryByLargeCategories = calcSummaryByLargeCategories(summaryByCategories, targetYear, targetMonth);
  // 未定義(中項目) を集計する
  const reSummaryWithUndefinedMiddleCategories = calcSummaryUndefinedMiddleCategories(reSummaryByLargeCategories, targetYear, targetMonth);
  // 合計 (変動費、固定費等の粒度) を集計する
  const reSummaryByCategoryMiddleClasses = calcSummaryByCategoryMiddleClasses(reSummaryWithUndefinedMiddleCategories, targetYear, targetMonth);
  // 総合計 (収入合計、支出合計の粒度) を集計する
  const reSummaryByCategoryLargeClasses = calcSummaryByCategoryLargeClasses(summaryByCategories, targetYear, targetMonth, reSummaryByCategoryMiddleClasses);
  // 未定義(大項目)を集計する
  const reSummaryWithUndefinedLargeCategories = calcSummaryUndefinedLargeCategories(reSummaryByCategoryLargeClasses, targetYear, targetMonth);
  // 収支合計を集計する
  const reSummaryForAllCategories = calcSummaryForAllCategories(reSummaryWithUndefinedLargeCategories, targetYear, targetMonth);

  // 金額を検算する
  validateCalcMonthlySummaryResult(reSummaryForAllCategories, summaryByCategories);

  // 集計結果を出力対象シートに出力する
  exportMonthlyResultDetails(reSummaryForAllCategories, targetYear, targetMonth, getCalcMcImportSheetPrefix(), getCalcMcImportAddr());

  console.log("end monthly summary by category");
}

// 集計結果を検算する
function validateCalcMonthlySummaryResult(reSummaryForAllCategories, summaryByCategories){
  const calculatedTotalAmount = reSummaryForAllCategories.find(s => s.largeCategoryName == "全体").amount;
  const expectedTotalAmount = sumAmountFromDetails(summaryByCategories);
  if(calculatedTotalAmount != expectedTotalAmount){
    throw new PfmUnexpectedError(
      "failed to validate calculatedTotalAmount: expected, got: " + expectedTotalAmount + ", " + calculatedTotalAmount,
      true
    );
  }
  console.log("valid: calculatedTotalAmount");
}

// 収支合計を集計する
function calcSummaryForAllCategories(reSummaryWithUndefinedLargeCategories, targetYear, targetMonth){
  return omitBlankElemFromArray(
    getOrderedCategoryConfigs().map(
      c => {
        // 収支合計の場合、総合計項目 (収入総合計と支出総合計) を合算して、金額を求める
        if(c.middleCategoryName == "収支合計"){
          // 対象明細の金額を集計して返す
          return getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, c,
            sumAmountFromDetails(reSummaryWithUndefinedLargeCategories.filter(s => s.middleCategoryName == "総合計")));
        }
        // その他の項目の場合、項目別集計結果をそのままマッピングする
        return getMatchedSummaryByCategory(c, reSummaryWithUndefinedLargeCategories, targetYear, targetMonth);
      }
    )
  );
}

// 未定義(大項目)を集計 (config で未定義の大項目がある場合を考慮)
function calcSummaryUndefinedLargeCategories(reSummaryByCategoryLargeClasses, targetYear, targetMonth){
  return omitBlankElemFromArray(
    getOrderedCategoryConfigs().map(
      c => {
        // 集計対象外とする項目について skip
        if(["収支合計"].includes(c.middleCategoryName)){
          return null;
        }
        // 未定義(大項目)の場合、収入もしくは支出の総合計から定義済み大項目の合計値を差し引いて、金額を求める
        if(c.middleCategoryName == "未定義(大項目)"){
          // 収入/支出の総合計を取得
          const targetCategoryLargeClass = c.largeCategoryName == "未定義収入" ? "収入"
            : c.largeCategoryName == "未定義支出" ? "支出" : null;
          const sumAmountAtTargetCategoryLargeClass = reSummaryByCategoryLargeClasses
            .find(s => s.largeCategoryName == targetCategoryLargeClass && s.middleCategoryName == "総合計")
            .amount
          // 収入/支出における定義済み大項目の合計値を集計
          const configsLargeCategoryDefined = getOrderedCategoryConfigs()
            .filter(r => r.categoryLargeClassName == targetCategoryLargeClass && r.middleCategoryName != "未定義(大項目)");
          const sumAmountAtLargeCategoryDefined = sumAmountFromDetails(
            filterMatchedByCategoryFromConfigs(reSummaryByCategoryLargeClasses, configsLargeCategoryDefined));
          // 上記の差額から未定義(大項目)の金額を求める
          const sumAmountAtLargeCategoryUndefined = sumAmountAtTargetCategoryLargeClass - sumAmountAtLargeCategoryDefined;
          // 結果明細にマップして返す (0円の場合は skip する)
          if(sumAmountAtLargeCategoryUndefined == 0){
            return null;
          }
          return getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, c, sumAmountAtLargeCategoryUndefined);
        }
        // その他の項目の場合、項目別集計結果をそのままマッピングする
        return getMatchedSummaryByCategory(c, reSummaryByCategoryLargeClasses, targetYear, targetMonth);
      }
    )
  );
}

// 総合計 (収入合計、支出合計の粒度) を集計する
function calcSummaryByCategoryLargeClasses(summaryByCategories, targetYear, targetMonth, reSummaryByCategoryMiddleClasses){
  return omitBlankElemFromArray(
    getOrderedCategoryConfigs().map(
      c => {
        // 集計対象外とする項目について skip
        if(["収支合計", "未定義(収入)", "未定義(支出)"].includes(c.middleCategoryName)){
          return null;
        }
        // 総合計項目の場合、紐付く項目別明細を集計する
        if(c.middleCategoryName == "総合計"){
          // 集計対象明細を収集する
          const summariesAtTargetCategoryLargeClass = c.largeCategoryName == "収入" ? summaryByCategories.filter(s => s.amount > 0) // 収入の場合: 金額が正のものを収入とみなす
            : c.largeCategoryName == "支出" ? summaryByCategories.filter(s => s.amount < 0) : null; // 支出の場合: 金額が負のものを支出とみなす
          // 対象明細の金額を集計して返す
          return getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, c, sumAmountFromDetails(summariesAtTargetCategoryLargeClass));
        }
        // その他の項目の場合、項目別集計結果をそのままマッピングする
        return getMatchedSummaryByCategory(c, reSummaryByCategoryMiddleClasses, targetYear, targetMonth);
      }
    )
  );
}

// 合計 (変動費、固定費等の粒度) を集計する
function calcSummaryByCategoryMiddleClasses(reSummaryWithUndefinedMiddleCategories, targetYear, targetMonth){
  return omitBlankElemFromArray(
    getOrderedCategoryConfigs().map(
      c => {
        // 集計対象外とする項目について skip
        if(["収支合計", "総合計", "未定義(収入)", "未定義(支出)"].includes(c.middleCategoryName)){
          return null;
        }
        // 合計項目の場合、紐付く小計項目を集計する
        if(c.middleCategoryName == "合計"){
          // 集計対象の項目を収集する
          const targetMiddleClassConfigs = getOrderedCategoryConfigs().filter(r => r.categoryMiddleClassName == c.largeCategoryName);
          const sumAmountByCategoryMiddleClass = sumAmountFromDetails(
            filterMatchedByCategoryFromConfigs(reSummaryWithUndefinedMiddleCategories, targetMiddleClassConfigs));
          return getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, c, sumAmountByCategoryMiddleClass);
        }
        // その他の項目の場合、項目別集計結果をそのままマッピングする
        return getMatchedSummaryByCategory(c, reSummaryWithUndefinedMiddleCategories, targetYear, targetMonth);
      }
    )
  );
}

// 未定義(中項目) を集計 (大項目内に config で未定義の中項目がある場合を考慮)
function calcSummaryUndefinedMiddleCategories(reSummaryByLargeCategories, targetYear, targetMonth){
  return omitBlankElemFromArray(
    getOrderedCategoryConfigs().map(
      c => {
        // 集計対象外とする項目について skip
        if(["収支合計", "総合計", "合計", "未定義(収入)", "未定義(支出)"].includes(c.middleCategoryName)){
          return null;
        }
        // 未定義(中項目)項目の場合、小計から定義済み項目の合計値を差し引いて、金額を求める
        if(c.middleCategoryName == "未定義(中項目)"){
          // 該当大項目の小計を取得
          const summaryAtTargetLargeCategory = reSummaryByLargeCategories
            .find(s => s.largeCategoryName == c.largeCategoryName && s.middleCategoryName == "小計");
          // 小計が存在しないということは、該当大項目配下には1件も明細が存在しないということなので skip
          if(summaryAtTargetLargeCategory == null){
            return null;
          }
          const sumAmountByLargeCategory = summaryAtTargetLargeCategory.amount;
          // 該当大項目の定義済み中項目の合計値を求める
          const sumAmountAtMiddleCategoryDefined = sumAmountFromDetails(
            reSummaryByLargeCategories
              .filter(s => s.largeCategoryName == c.largeCategoryName && s.middleCategoryName != "小計"));
          // 上記の差額から未定義(中項目)の金額を求める
          const sumAmountAtMiddleCategoryUndefined = sumAmountByLargeCategory - sumAmountAtMiddleCategoryDefined;
          // 結果明細にマップして返す (0円の場合は skip する)
          if(sumAmountAtMiddleCategoryUndefined == 0){
            return null;
          }
          return getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, c, sumAmountAtMiddleCategoryUndefined);
        }
        // その他の項目の場合、項目別集計結果をそのままマッピングする
        return getMatchedSummaryByCategory(c, reSummaryByLargeCategories, targetYear, targetMonth);
      }
    )
  );
}

// 大項目別に小計を集計する
function calcSummaryByLargeCategories(summaryByCategories, targetYear, targetMonth){
  return omitBlankElemFromArray(
    getOrderedCategoryConfigs().map(
      c => {
        // 集計対象外とする項目について skip
        if(["収支合計", "総合計", "合計", "未定義(収入)", "未定義(支出)", "未定義(中項目)"].includes(c.middleCategoryName)){
          return null;
        }
        // 小計項目の場合、大項目別に集計する
        if(c.middleCategoryName == "小計"){
          const sumAmountByCategory = sumAmountFromDetails(summaryByCategories.filter(s => s.largeCategoryName == c.largeCategoryName));
          if(sumAmountByCategory == 0) {
            return null;
          }
          return getMonthlySummaryFromCategoryConfig(targetYear, targetMonth, c, sumAmountByCategory);
        }
        // その他の項目の場合、項目別集計結果をそのままマッピングする
        return getMatchedSummaryByCategory(c, summaryByCategories, targetYear, targetMonth);
      }
    )
  )
}

// 集計対象月の明細について、項目別に集計する
function calcTargetMonthSummaryByCategories(targetYear, targetMonth){
  // 集計対象月の明細データを取得する
  const targetDetails = fetchDailySummaryByCategoryDetails(targetYear, targetMonth);
  if(targetDetails.length == 0) return null;

  // 明細データから項目を収集する
  const targetCategories = getUniqueArrayFrom(targetDetails.map(d => d.category));
  // 項目別に明細を集計する
  return targetCategories.map(
    c => {
      const sumAmountByCategory = sumAmountFromDetails(targetDetails.filter(d => d.category == c));
      return getMonthlySummaryFromCategory(targetYear, targetMonth, c, sumAmountByCategory);
    }
  );
}

// 与えた configs と大項目と中項目が一致するものに filter する
function filterMatchedByCategoryFromConfigs(inputs, configs){
  const targetCategories = configs.map(c => c.largeCategoryName + ">>" + c.middleCategoryName);
  return inputs.filter(r => targetCategories.includes(r.largeCategoryName + ">>" + r.middleCategoryName));
}

// config の項目が一致する既存集計結果をマッピングして返す
function getMatchedSummaryByCategory(categoryConfig, summaryByCategories, targetYear, targetMonth){
  const matchedSummaryByCategory = summaryByCategories
    .find(s => s.largeCategoryName == categoryConfig.largeCategoryName && s.middleCategoryName == categoryConfig.middleCategoryName);
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

// 集計対象月の項目別集計元データ (日次項目別集計明細) を取得する
function fetchDailySummaryByCategoryDetails(targetYear, targetMonth){
  return fetchDetailsFromCell(
    getCalcDcExportSheetPrefix() + "_" + targetYear + targetMonth,
    getCalcDcExportAddr(),
    rowElems => {
      return {
        date:     rowElems[0], // 日付
        category: rowElems[1], // 項目 (大項目>>中項目)
        amount:   rowElems[2], // 金額
      };
    }
  );
}

// 月次集計の集計対象月を返す
function getMonthlyCalcTargetYearMonth(monthsAgo){
  const {someMonthsAgoYyyy, someMonthsAgoMm} = getYearMonthMonthsAgoFrom(getCurrentYyyy(), getCurrentMm(), monthsAgo);
  const targetYear = someMonthsAgoYyyy;
  const targetMonth = someMonthsAgoMm;
  return {targetYear, targetMonth};
}
