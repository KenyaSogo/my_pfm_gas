/** @OnlyCurrentDoc */

// TODO: delete
function testCalcWeeklySummary(){
  calcWeeklySummary(0);
  // calcWeeklySummary(1);
  // calcWeeklySummary(2);
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
  if(targetWeekDetails.length == 0){
    console.log("end weekly summary by category: skip summary and export: there is no data");
    return;
  }

  // それらを項目別に集計する TODO: calc monthly との項目別集計ロジック共通化
  const targetWeekStartDateStr = getYyyyMmDdFrom(targetDates[0]);
  const targetWeekSummaryByCategories = calcWeeklySummaryFromCategory(targetWeekDetails, targetWeekStartDateStr);

  // 項目を定められた順番に整えつつ再集計 (= 小計、合計、総合計の集計) する
  // 大項目別に小計を集計する
  const reSummaryByLargeCategories = calcWeeklySummaryByLargeCategories(targetWeekSummaryByCategories, targetWeekStartDateStr);
  // 未定義(中項目) を集計する
  const reSummaryWithUndefinedMiddleCategories = calcWeeklySummaryUndefinedMiddleCategories(reSummaryByLargeCategories, targetWeekStartDateStr);
  // 合計 (変動費、固定費等の粒度) を集計する
  const reSummaryByCategoryMiddleClasses = calcWeeklySummaryByCategoryMiddleClasses(reSummaryWithUndefinedMiddleCategories, targetWeekStartDateStr);
  // 総合計 (収入合計、支出合計の粒度) を集計する
  const reSummaryByCategoryLargeClasses = calcWeeklySummaryByCategoryLargeClasses(targetWeekSummaryByCategories, targetWeekStartDateStr, reSummaryByCategoryMiddleClasses);
  // 未定義(大項目)を集計する
  const reSummaryWithUndefinedLargeCategories = calcWeeklySummaryUndefinedLargeCategories(reSummaryByCategoryLargeClasses, targetWeekStartDateStr);
  // 収支合計を集計する
  const reSummaryForAllCategories = calcWeeklySummaryForAllCategories(reSummaryWithUndefinedLargeCategories, targetWeekStartDateStr);

  // 金額を検算する
  validateCalcWeeklySummaryResult(reSummaryForAllCategories, targetWeekSummaryByCategories);

  // 集計結果を出力対象シートに出力する
  exportWeeklyResultDetails(reSummaryForAllCategories, weeksAgo, getCalcWcImportSheetPrefix(), getCalcWcImportAddr());

  console.log("end weekly summary by category");
}

// 集計結果を検算する TODO: calc monthly との項目別集計ロジック共通化
function validateCalcWeeklySummaryResult(reSummaryForAllCategories, summaryByCategories){
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

// 収支合計を集計する TODO: calc monthly との項目別集計ロジック共通化
function calcWeeklySummaryForAllCategories(reSummaryWithUndefinedLargeCategories, targetWeekStartDateStr){
  return omitBlankElemFromArray(
    getOrderedCategoryConfigs().map(
      c => {
        // 収支合計の場合、総合計項目 (収入総合計と支出総合計) を合算して、金額を求める
        if(c.middleCategoryName == "収支合計"){
          // 対象明細の金額を集計して返す
          return getWeeklySummaryFromCategoryConfig(targetWeekStartDateStr, c,
            sumAmountFromDetails(reSummaryWithUndefinedLargeCategories.filter(s => s.middleCategoryName == "総合計")));
        }
        // その他の項目の場合、項目別集計結果をそのままマッピングする
        return getMatchedWeeklySummaryByCategory(c, reSummaryWithUndefinedLargeCategories, targetWeekStartDateStr);
      }
    )
  );
}

// 未定義(大項目)を集計 (config で未定義の大項目がある場合を考慮) TODO: calc monthly との項目別集計ロジック共通化
function calcWeeklySummaryUndefinedLargeCategories(reSummaryByCategoryLargeClasses, targetWeekStartDateStr){
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
          return getWeeklySummaryFromCategoryConfig(targetWeekStartDateStr, c, sumAmountAtLargeCategoryUndefined);
        }
        // その他の項目の場合、項目別集計結果をそのままマッピングする
        return getMatchedWeeklySummaryByCategory(c, reSummaryByCategoryLargeClasses, targetWeekStartDateStr);
      }
    )
  );
}

// 総合計 (収入合計、支出合計の粒度) を集計する TODO: calc monthly との項目別集計ロジック共通化
function calcWeeklySummaryByCategoryLargeClasses(summaryByCategories, targetWeekStartDateStr, reSummaryByCategoryMiddleClasses){
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
          return getWeeklySummaryFromCategoryConfig(targetWeekStartDateStr, c, sumAmountFromDetails(summariesAtTargetCategoryLargeClass));
        }
        // その他の項目の場合、項目別集計結果をそのままマッピングする
        return getMatchedWeeklySummaryByCategory(c, reSummaryByCategoryMiddleClasses, targetWeekStartDateStr);
      }
    )
  );
}

// 合計 (変動費、固定費等の粒度) を集計する TODO: calc monthly との項目別集計ロジック共通化
function calcWeeklySummaryByCategoryMiddleClasses(reSummaryWithUndefinedMiddleCategories, targetWeekStartDateStr){
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
          return getWeeklySummaryFromCategoryConfig(targetWeekStartDateStr, c, sumAmountByCategoryMiddleClass);
        }
        // その他の項目の場合、項目別集計結果をそのままマッピングする
        return getMatchedWeeklySummaryByCategory(c, reSummaryWithUndefinedMiddleCategories, targetWeekStartDateStr);
      }
    )
  );
}

// 未定義(中項目) を集計 (大項目内に config で未定義の中項目がある場合を考慮) TODO: calc monthly との項目別集計ロジック共通化
function calcWeeklySummaryUndefinedMiddleCategories(reSummaryByLargeCategories, targetWeekStartDateStr){
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
          return getWeeklySummaryFromCategoryConfig(targetWeekStartDateStr, c, sumAmountAtMiddleCategoryUndefined);
        }
        // その他の項目の場合、項目別集計結果をそのままマッピングする
        return getMatchedWeeklySummaryByCategory(c, reSummaryByLargeCategories, targetWeekStartDateStr);
      }
    )
  );
}

// 大項目別に小計を集計する TODO: calc monthly との項目別集計ロジック共通化
function calcWeeklySummaryByLargeCategories(summaryByCategories, targetWeekStartDateStr){
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
          return getWeeklySummaryFromCategoryConfig(targetWeekStartDateStr, c, sumAmountByCategory);
        }
        // その他の項目の場合、項目別集計結果をそのままマッピングする
        return getMatchedWeeklySummaryByCategory(c, summaryByCategories, targetWeekStartDateStr);
      }
    )
  )
}

// 集計対象週のデータを項目別に集計する TODO: calc monthly との項目別集計ロジック共通化
function calcWeeklySummaryFromCategory(targetWeekDetails, targetWeekStartDateStr){
  // 明細データから項目を収集する
  const targetCategories = getUniqueArrayFrom(targetWeekDetails.map(d => d.category));
  // 項目別に明細を集計する
  return targetCategories.map(
    c => {
      const sumAmountByCategory = sumAmountFromDetails(targetWeekDetails.filter(d => d.category == c));
      return getWeeklySummaryFromCategory(targetWeekStartDateStr, c, sumAmountByCategory);
    }
  );
}

// config の項目が一致する既存集計結果をマッピングして返す TODO: calc monthly との項目別集計ロジック共通化
function getMatchedWeeklySummaryByCategory(categoryConfig, summaryByCategories, weekStartDateStr){
  const matchedSummaryByCategory = summaryByCategories
    .find(s => s.largeCategoryName == categoryConfig.largeCategoryName && s.middleCategoryName == categoryConfig.middleCategoryName);
  if(matchedSummaryByCategory){
    return getWeeklySummaryFromCategoryConfig(weekStartDateStr, categoryConfig, matchedSummaryByCategory.amount);
  }
  return null;
}

// 項目定義 (CategoryConfig) を元に集計結果明細を作成し返す
function getWeeklySummaryFromCategoryConfig(weekStartDateStr, config, amount){
  return getWeeklySummary(
    weekStartDateStr,
    config.largeCategoryName,
    config.middleCategoryName,
    amount);
}

// 項目 (大項目>>中項目) を元に集計結果明細を作成し返す
function getWeeklySummaryFromCategory(weekStartDateStr, category, amount){
  return getWeeklySummary(
    weekStartDateStr,
    category.split(">>")[0],
    category.split(">>")[1],
    amount);
}

// 集計結果明細を返す
function getWeeklySummary(weekStartDateStr, largeCategoryName, middleCategoryName, amount){
  return {
    startDate:          weekStartDateStr,   // 対象週開始日
    largeCategoryName:  largeCategoryName,  // 大項目
    middleCategoryName: middleCategoryName, // 中項目
    amount:             amount,             // 金額
  };
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
