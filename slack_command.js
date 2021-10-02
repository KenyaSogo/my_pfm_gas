// slack command からの post リクエストを処理する
function doPost(e){
  if(e.command == "/test-pfm"){
    executeScenarioTest();
    const response = {text: "pfm: execute scenario test was kicked"};
    return ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
