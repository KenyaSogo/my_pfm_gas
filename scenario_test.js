/** @OnlyCurrentDoc */

// TODO: テストの仕組み memo:
//   - env グローバル変数を設ける
//   - テストを起動した際には、これに "test" をセットする
//   - これを読み取って、外部IF回りのコードを分岐させる
//     - getSettingSheet
//     - UrlFetchApp ~ phantomJs kick
//   - jest 使いたかったが、export / import を既存コードに仕込むのが上手く動かず今回は断念

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
