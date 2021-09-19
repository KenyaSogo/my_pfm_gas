/** @OnlyCurrentDoc */

let isTestEnv = false;
function setTestEnv(){
  isTestEnv = true;
}
function getIsTestEnv(){
  return isTestEnv;
}

function executeScenarioTest(){
  postConsoleAndSlackJobStart("execute scenario test");
  try {
    doScenarioTest();
  } catch(error){
    handleError(error);
  }
  postConsoleAndSlackJobEnd("execute scenario test");
}

function doScenarioTest(){
  // テスト環境フラグを立てる
  setTestEnv();

  // aggregate をテスト
  testAggregate();

  // calculate_daily をテスト
  testCalculateDaily();

  // calculate_monthly をテスト
  testCalculateMonthly();

  // report をテスト
  // TODO: 実装

  // check_maintenance, create_sheet, update_month のテストは skip
}

// calculate_monthly.js のテスト
function testCalculateMonthly(){
  console.log("start: testCalculateMonthly");

  // 検証対象の処理を実行
  calcMonthlySummary(0);

  // 結果検証
  [
    ["calc monthly", "category summary", getTestCalcMcResultExpected, getTestCalcMcResultGot],
  ].forEach(t => validateScenarioResult(...t));

  console.log("end: testCalculateMonthly");
}

// calculate_daily.js のテスト
function testCalculateDaily(){
  console.log("start: testCalculateDaily");

  // 検証対象の処理を実行
  calcDailySummary(0);

  // 結果検証
  [
    ["calc daily: cash balance", "earlier month", getTestCalcDcbResultExpectedEarlier, getTestCalcDcbResultGotEarlier],
    ["calc daily: cash balance", "later month", getTestCalcDcbResultExpectedLater, getTestCalcDcbResultGotLater],
    ["calc daily: asset balance", "earlier month", getTestCalcDabResultExpectedEarlier, getTestCalcDabResultGotEarlier],
    ["calc daily: asset balance", "later month", getTestCalcDabResultExpectedLater, getTestCalcDabResultGotLater],
    ["calc daily: category summary", "earlier month", getTestCalcDcResultExpectedEarlier, getTestCalcDcResultGotEarlier],
    ["calc daily: category summary", "later month", getTestCalcDcResultExpectedLater, getTestCalcDcResultGotLater],
  ].forEach(t => validateScenarioResult(...t));

  console.log("end: testCalculateDaily");
}

// aggregate.js のテスト
function testAggregate(){
  console.log("start: testAggregate");

  // 検証対象の処理を実行
  aggregateByMonth(0);

  // 結果検証
  [
    ["test aggregate", "pfm account 1", getTestAggreResultExpected1, getTestAggreResultGot1],
    ["test aggregate", "pfm account 2", getTestAggreResultExpected2, getTestAggreResultGot2],
  ].forEach(t => validateScenarioResult(...t));

  console.log("end: testAggregate");
}

// 結果の検証 (期待値と実際の結果の突き合わせ) を行う
function validateScenarioResult(testCaseCategoryName, testCaseName, expectedValueGetFunc, gotValueGetFunc){
  if(gotValueGetFunc() != expectedValueGetFunc()){
    throw new PfmUnexpectedError(
      testCaseCategoryName + " failed: " + testCaseName + ": expected, got: " + expectedValueGetFunc() + ", " + gotValueGetFunc(),
      true);
  }
}
