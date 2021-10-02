// slack command からの post リクエストを処理する
function doPost(e){
  if(e.command == "/test-pfm"){
    executeScenarioTest();
    return getSlackCommandPlainResponse("my pfm: execute scenario test was kicked");
  }

  return getSlackCommandPlainResponse("my pfm: command not found: " + e.command);
}

function getSlackCommandPlainResponse(message){
  const response = {text: message};
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}
