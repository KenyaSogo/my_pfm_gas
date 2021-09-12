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
}

// aggregate.js のテスト
function testAggregate(){
  console.log("start: testAggregate");
  aggregateByMonth(0);
  console.log("end: testAggregate");
}
