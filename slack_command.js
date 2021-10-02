// slack command からの post リクエストを処理する
function doPost(e){
  console.log("start: do post");
  console.log("e: " + e);

  // TODO: token の検証
  if(e.parameter.command == "/test-pfm"){
    // executeScenarioTest の時間主導型トリガーを発火させる (3秒以内に返さないと slack コマンドがタイムアウトするため非同期実行とする)
    ScriptApp.newTrigger('executeScenarioTest')
      .timeBased()
      .after(1000)
      .create();

    console.log("end: do post");
    return getSlackCommandPlainResponse("my pfm: execute scenario test was kicked");
  }

  console.log("end: do post");
  return getSlackCommandPlainResponse("my pfm: command not found: " + e.command);
}

function getSlackCommandPlainResponse(message){
  const response = {text: message};
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}
