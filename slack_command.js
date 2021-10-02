// slack command からの post リクエストを処理する
function doPost(e){
  console.log("start: do post");

  // token を検証する
  if(e.parameter.token != getMyPfmSlackAppVerificationToken()){
    postConsoleAndSlackWarning("doPost: invalid token", true);
    return getSlackCommandPlainResponse("my pfm: error: invalid token");
  }

  // コマンドに紐付く func のトリガーを発火させる
  const commandFromParam = e.parameter.command;
  const matchedCommandConfig = getSlackCommandConfigs().find(c => c.command == commandFromParam);
  if(matchedCommandConfig){
    // トリガー発火
    triggerJobExecutionAsynchronously(matchedCommandConfig.funcName);
    postConsoleAndSlackInfoMessage(`job execution was triggered: slack command, func: ${matchedCommandConfig.command}, ${matchedCommandConfig.funcName}`);

    console.log("end: do post");
    return getSlackCommandPlainResponse(`my pfm: ${matchedCommandConfig.funcName} was triggered`);
  }

  // 未知のコマンドだった場合は not found コメントを返して終わる
  console.log("end: do post");
  postConsoleAndSlackInfoMessage("job execution was not triggered: slack command: " + commandFromParam);
  return getSlackCommandPlainResponse("my pfm: error: command not found: " + commandFromParam);
}

// 関数名で指定された job execution の時間主導型トリガーを発火させる (3秒以内に返さないと slack コマンドがタイムアウトするため非同期実行とする)
function triggerJobExecutionAsynchronously(funcName){
  ScriptApp.newTrigger(funcName)
    .timeBased()
    .after(1000)
    .create();
}


// slack コマンドへのプレーンなレスポンスを取得する
function getSlackCommandPlainResponse(message){
  const response = {text: message};
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}
