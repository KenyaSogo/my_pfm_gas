/** @OnlyCurrentDoc */

function executeControl(){
  console.log("start executeControl");
  // exec group の trigger 判定用の現在時刻セルに値をセットする
  const currentTime = getCurrent_yyyy_mm_dd_hh_mm_ss();
  getCurrentTimeForExecTriggerRange().setValue(currentTime);
  // 各 exec group の activate trigger を処理する
  handleExecGroupTriggers(currentTime); // TODO: sheet io 系 memo 化による高速化

  // 実行中の処理がある場合は終了する
  if(getRunningExecExists()){
    console.log("end executeControl: running exec existed");
    return;
  }

  // 各 exec group を順に起動する
  let isExecOnceBooted = false; // 以下のループで複数の exec を起動しないように制御するためのフラグ
  getExecGroupNames().forEach(
    execGroupName => {
      console.log("start handle exec group: execGroupName: " + execGroupName);
      // exec が一度起動された場合は後続の exec を全て skip する
      if(isExecOnceBooted){
        console.log("skip exec group: isExecOnceBooted: execGroupName: " + execGroupName);
        return;
      }
      // 該当 exec group が activate されていない場合は skip する
      const execGrActivatedAtRange = getExecGrActivatedAtRangeFrom(execGroupName);
      if(!isActiveExecOrExecGroup(execGrActivatedAtRange)){
        console.log("skip exec group: is not activate: execGroupName: " + execGroupName);
        return;
      }
      // 該当 exec group 配下の exec を順に起動する
      const groupExecNames = getGroupExecNamesFrom(execGroupName);
      groupExecNames.forEach(
        execName => {
          console.log("start handle exec: execName: " + execName);
          // exec が一度起動された場合は後続の exec を全て skip する
          if(isExecOnceBooted){
            console.log("skip handle exec: isExecOnceBooted: execName: " + execName);
            return;
          }
          // 該当 exec が activate されていない場合は skip する
          const execActivatedAtRange = getExecActivatedAtRangeFrom(execName);
          if(!isActiveExecOrExecGroup(execActivatedAtRange)){
            console.log("skip handle exec: is not activate: execName: " + execName);
            return;
          }
          // exec 起動済みフラグを立てる
          isExecOnceBooted = true;
          try {
            // 該当 exec を起動する TODO: 各 execute~ でやっている handleError , start/end logging をここに移管する
            console.log("handle exec: boot exec: execName: " + execName);
            bootExec(execName);
          } catch(error){
            // 異常終了した場合
            console.log("handle exec: exec failed: execName: " + execName);
            handleExecFailed(execActivatedAtRange, execGrActivatedAtRange, error);
          }
          // 正常終了した場合
          handleExecSucceeded(execActivatedAtRange, groupExecNames, execName, execGrActivatedAtRange);
          console.log("end handle exec: execName: " + execName);
        }
      );
    }
  );
  console.log("end executeControl");
}

// exec group の activate trigger を処理する
function handleExecGroupTriggers(currentTime){
  // 各 exec group の trigger flag を取得し、on になっているものについて activated_at を立てる
  getExecGroupTriggerFlags().forEach(
    triggerFlag => {
      if(triggerFlag.isTriggerOn == 1){
        // exec group を activated する
        activateExecOrExecGroup(getExecGrActivatedAtRangeFrom(triggerFlag.execGroupName), currentTime);
        // exec group の先頭の exec を activated する
        activateExecOrExecGroup(getExecActivatedAtRangeFrom(getGroupExecNamesFrom(triggerFlag.execGroupName)[0]), currentTime);
      }
    }
  );
}

// exec を起動する
function bootExec(targetExecName){
  // 該当 exec 実行中の旨登録する
  registerExecRunning(targetExecName);
  // 該当 exec を実行する
  execFuncFrom(targetExecName);
}

// exec の正常終了を処理する TODO: aggre の更新無し時後続 skip 対応
function handleExecSucceeded(execActivatedAtRange, groupExecNames, currentExecName, execGrActivatedAtRange){
  // 該当 exec を deactivate する
  deactivateExecOrExecGroup(execActivatedAtRange);
  // exec group における最後の exec だった場合
  const currentExecOrderIndex = groupExecNames.indexOf(currentExecName);
  if(currentExecOrderIndex == groupExecNames.length - 1){
    // 該当 exec および exec group 完了の旨登録する
    registerRunningExecSucceededAndExecGroupFinished();
    // 該当 exec group を deactivate する
    deactivateExecOrExecGroup(execGrActivatedAtRange);
    return;
  }
  // 該当 exec 完了および次の exec 準備中の旨登録する
  registerRunningExecSucceededAndNextExecPreparing();
  // 次の exec を activate する
  getExecActivatedAtRangeFrom(groupExecNames[currentExecOrderIndex + 1]).setValue(getCurrent_yyyy_mm_dd_hh_mm_ss());
}

// exec の異常終了を処理する
function handleExecFailed(execActivatedAtRange, execGrActivatedAtRange, error){
  // 該当 exec 失敗の旨登録する
  registerRunningExecFailed();
  // 該当 exec を deactivate する
  deactivateExecOrExecGroup(execActivatedAtRange);
  // 該当 exec group を deactivate する
  deactivateExecOrExecGroup(execGrActivatedAtRange);
  // 異常終了させる
  throw error;
}

// 実行中の exec があるか否かを返す
function getRunningExecExists(){
  return getRunningExecStatusRange().getValue() == execStatusRunning;
}

// 指定した exec が実行中の旨を登録する
function registerExecRunning(execName){
  getRunningExecNameRange().setValue(execName);
  getRunningExecStatusRange().setValue(execStatusRunning);
}

// 実行中の exec が完了した旨を登録する
function registerRunningExecSucceededAndExecGroupFinished(){
  getRunningExecStatusRange().setValue(execStatusSucceededAndExecGroupFinished);
}

// 実行中の exec が完了し次の exec を準備している旨を登録する
function registerRunningExecSucceededAndNextExecPreparing(){
  getRunningExecStatusRange().setValue(execStatusSucceededAndNextExecPreparing);
}

// 実行中の exec が失敗した旨を登録する
function registerRunningExecFailed(){
  getRunningExecStatusRange().setValue(execStatusFailed);
}

// exec / exec group を activate する
function activateExecOrExecGroup(activatedAtRange, currentTime){
  activatedAtRange.setValue(currentTime);
}

// exec / exec group を deactivate する
function deactivateExecOrExecGroup(activatedAtRange){
  activatedAtRange.setValue("");
}

function isActiveExecOrExecGroup(activatedAtRange){
  return activatedAtRange.getValue() != "";
}
