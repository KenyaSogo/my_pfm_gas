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

  if(getTestAggreResultGot1() != getTestAggreResultExpected1()){
    throw new PfmUnexpectedError(
      "test aggregate failed: pfm account 1: expected, got: " + getTestAggreResultExpected1() + ", " + getTestAggreResultGot1(),
      true);
  }
  if(getTestAggreResultGot2() != getTestAggreResultExpected2()){
    throw new PfmUnexpectedError(
      "test aggregate failed: pfm account 2: expected, got: " + getTestAggreResultExpected1() + ", " + getTestAggreResultGot2(),
      true);
  }

  console.log("end: testAggregate");
}
