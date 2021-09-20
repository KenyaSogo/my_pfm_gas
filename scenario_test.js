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

  // update_graph をテスト
  testUpdateGraph();

  // report をテスト
  testReport();

  // check_maintenance, create_sheet, update_month のテストは skip
}

// report.js のテスト
function testReport(){
  console.log("start: testReport");

  // 検証対象の処理を実行
  doReport();

  // 結果検証
  [
    ["test report", "main scenario", getTestReportResultExpected, getTestReportResultGot],
  ].forEach(t => validateScenarioResult(...t));

  console.log("end: testReport");
}

// update_graph.js のテスト
function testUpdateGraph(){
  console.log("start: testUpdateGraph");

  // 検証対象の処理を実行
  doUpdateGraphSourceData();

  // 結果検証
  [
    ["test update graph: daily balance: cash", "main scenario", getTestUpdateGraphDcbResultExpected, getTestUpdateGraphDcbResultGot],
  ].forEach(t => validateScenarioResult(...t));

  console.log("end: testUpdateGraph");
}

// calculate_monthly.js のテスト
function testCalculateMonthly(){
  console.log("start: testCalculateMonthly");

  // 検証対象の処理を実行
  calcMonthlySummary(0);

  // 結果検証
  [
    ["test calc monthly: category summary", "main scenario", getTestCalcMcResultExpected, getTestCalcMcResultGot],
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
    ["test calc daily: cash balance", "earlier month", getTestCalcDcbResultExpectedEarlier, getTestCalcDcbResultGotEarlier],
    ["test calc daily: cash balance", "later month", getTestCalcDcbResultExpectedLater, getTestCalcDcbResultGotLater],
    ["test calc daily: asset balance", "earlier month", getTestCalcDabResultExpectedEarlier, getTestCalcDabResultGotEarlier],
    ["test calc daily: asset balance", "later month", getTestCalcDabResultExpectedLater, getTestCalcDabResultGotLater],
    ["test calc daily: category summary", "earlier month", getTestCalcDcResultExpectedEarlier, getTestCalcDcResultGotEarlier],
    ["test calc daily: category summary", "later month", getTestCalcDcResultExpectedLater, getTestCalcDcResultGotLater],
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
