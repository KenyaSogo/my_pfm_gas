/** @OnlyCurrentDoc */

// 現在の日付 (yyyy/MM/dd 形式) を返す
function getCurrent_yyyy_MM_dd(){
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
}

// 現在の時刻 (yyyy/MM/dd hh:mm:ss 形式) を返す
function getCurrent_yyyy_mm_dd_hh_mm_ss(){
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd hh:mm:ss");
}

// 現在の日付の年 (yyyy 形式) を返す
function getCurrentYyyy(){
  return getYyyyFrom(new Date());
}

// 現在の日付の月 (MM 形式) を返す
function getCurrentMm(){
  return getMmFrom(new Date());
}

// Date を yyyy 形式の文字列にして返す
function getYyyyFrom(date){
  return Utilities.formatDate(date, "Asia/Tokyo", "yyyy");
}

// Date を MM 形式の文字列にして返す
function getMmFrom(date){
  return Utilities.formatDate(date, "Asia/Tokyo", "MM");
}

// Date の n ヶ月前を返す
function getDateMonthsAgoFrom(date, agoNum){
  return new Date(date.getFullYear(), date.getMonth() - agoNum, 1);
}

// Date の前月を返す
function getPreviousMonthDate(date){
  return new Date(date.getFullYear(), date.getMonth() - 1, 1);
}

// Date の翌月を返す
function getNextMonthDate(date){
  return new Date(date.getFullYear(), date.getMonth() + 1, 1);
}

// 指定月日に対して n ヶ月前の月および日を返す
function getYearMonthMonthsAgoFrom(targetYear, targetMonth, agoNum){
  const currentMonthDate = new Date([targetYear, targetMonth, "01"].join("/"));
  const dateSomeMonthsAgo = getDateMonthsAgoFrom(currentMonthDate, agoNum);
  const someMonthsAgoYyyy = getYyyyFrom(dateSomeMonthsAgo);
  const someMonthsAgoMm = getMmFrom(dateSomeMonthsAgo);

  return {someMonthsAgoYyyy, someMonthsAgoMm};
}

// 指定月日に対して前月の月および日を返す
function getPreviousYearMonth(targetYear, targetMonth){
  const currentMonthDate = new Date([targetYear, targetMonth, "01"].join("/"));
  const previousMonthDate = getPreviousMonthDate(currentMonthDate);
  const previousMonthYyyy = getYyyyFrom(previousMonthDate);
  const previousMonthMm = getMmFrom(previousMonthDate);

  return {previousMonthYyyy, previousMonthMm};
}

// 指定月日に対して翌月の月および日を返す
function getNextYearMonth(targetYear, targetMonth){
  const currentMonthDate = new Date([targetYear, targetMonth, "01"].join("/"));
  const nextMonthDate = getNextMonthDate(currentMonthDate);
  const nextMonthYyyy = getYyyyFrom(nextMonthDate);
  const nextMonthMm = getMmFrom(nextMonthDate);

  return {nextMonthYyyy, nextMonthMm};
}

// 指定した列範囲に格納されている値から、ヘッダー行と空白行を除いた配列を返す
function getTrimmedColumnValues(sheet, columnRangeName){
  return sheet.getRange(columnRangeName).getValues()
    .map(r => r[0]) // 各行の値を取り出す
    .filter(v => v !== "") // 空白行を除外する (0 がセットされているセルを空白行扱いして除いてしまうため、厳密不等価演算子を使う)
    .slice(1); // ヘッダー行を除外する
}

// 0...elemNum の range (ex. elemNum=3 の場合 [0, 1, 2]) を返す
function getIntRangeFromZero(elemNum){
  return Array.from(Array(elemNum).keys());
}

// 配列の値を一意に集約して返す
function getUniqueArrayFrom(array){
  return Array.from(new Set(array));
}

// 各明細の金額を合計して返す ※details における金額項目の key は amount である必要がある
function sumAmountFromDetails(targetDetails){
  return targetDetails
    .map(d => Number(d.amount))
    .reduce((sum, amount) => sum + amount, 0);
}

// 数値文字列を表示用の文字列に変換して返す (99999 -> ¥99,999)
function formatNumStr(numStr){
  return "¥" + Number(numStr).toLocaleString();
}

// slack の my-pfm-report チャンネルにポストする
function postSlackReportingChannel(message){
  postSlackTo(getSlackReportingBotWebhookUrl(), convertSlackCodeFormatMessage(message));
}

// slack の my-pfm-dev-notif チャンネルにポストする
function postSlackLoggingChannel(message){
  postSlackTo(getSlackLoggingBotWebhookUrl(), convertSlackCodeFormatMessage(message));
}

// slack の my-pfm-dev-alert チャンネルにポストする
function postSlackAlertChannel(message, needsMention){
  postSlackTo(getSlackAlertBotWebhookUrl(), (needsMention ? "<!channel>\n" : "") + convertSlackCodeFormatMessage(message));
}

// slack の code フォーマット化されたメッセージを取得する
function convertSlackCodeFormatMessage(message){
  let formattedMessage = "```";
  formattedMessage += message;
  formattedMessage += "```";
  return formattedMessage;
}

// slack の webHookUrl に対してポストする
function postSlackTo(webHookUrl, message){
  const options = {
    method:      "post",
    contentType: "application/json",
    payload:     JSON.stringify({text: message}),
  };
  UrlFetchApp.fetch(webHookUrl, options);
}

// console と slack に job 開始メッセージをポストする
function postConsoleAndSlackJobStart(message){
  postConsoleAndSlackJobStartOrEnd("start", message);
}

// console と slack に job 終了メッセージをポストする
function postConsoleAndSlackJobEnd(message){
  postConsoleAndSlackJobStartOrEnd("end", message);
}

// console と slack に job 開始もしくは終了メッセージをポストする
function postConsoleAndSlackJobStartOrEnd(startOrEndString, message){
  const postMessageConsole = startOrEndString + ": " + message;
  const postMessageSlack = getCurrent_yyyy_mm_dd_hh_mm_ss() + ": " + postMessageConsole;
  console.log(postMessageConsole);
  postSlackLoggingChannel(postMessageSlack);
}

// ワーニングをポストする (その場で alert チャンネルに通知 / コンソールにエラー出力だけして処理を継続するもの)
function postConsoleAndSlackWarning(message, needsMention){
  const warningMessage = "warning: " + message;
  const warningMessageSlack = getCurrent_yyyy_mm_dd_hh_mm_ss() + ": " + warningMessage;
  console.error(warningMessage);
  postSlackAlertChannel(warningMessageSlack, needsMention);
}

// 想定エラー (処理中断するが正常終了扱いにするもの)
class PfmExpectedError extends Error {
  constructor(message, needsMention) {
    super(message);
    this.name = "PfmExpectedError";
    this.needsMention = needsMention;
  }
}

// 想定外エラー (処理中断して異常終了扱いにするもの)
class PfmUnexpectedError extends Error {
  constructor(message, needsMention) {
    super(message);
    this.name = "PfmUnexpectedError";
    this.needsMention = needsMention;
  }
}

// エラーをハンドリングする
function handleError(error){
  // 想定エラーの場合、メッセージをポストした後そのまま抜ける
  if(error instanceof PfmExpectedError){
    postConsoleAndSlackExpectedError(error, error.needsMention);
    return;
  }

  // 想定外エラーの場合、メッセージをポストした後 error を throw して異常終了させる
  if(error instanceof PfmUnexpectedError){
    postConsoleAndSlackUnexpectedError(error, error.needsMention);
  } else {
    postConsoleAndSlackUnexpectedError(error, true);
  }
  throw error;
}

function testAlert(){
  postConsoleAndSlackJobStart("test alert");
  try {
    // warning
    // postConsoleAndSlackWarning("test: warning", true);
    // expected error
    // throw new PfmExpectedError("test: expected error", true);
    // unexpected error
    // throw new PfmUnexpectedError("test: unexpected error", false);
    // exception
    const hoge = undefined[0];
  } catch(error){
    handleError(error);
  }
  postConsoleAndSlackJobEnd("test alert");
}

// 想定エラーをポストする
function postConsoleAndSlackExpectedError(error, needsMention){
  postConsoleAndSlackError(error, true, needsMention);
}

// 想定外エラーをポストする
function postConsoleAndSlackUnexpectedError(error, needsMention){
  postConsoleAndSlackError(error, false, needsMention);
}

// エラーをポストする
function postConsoleAndSlackError(error, isExpected, needsMention){
  const errorMessage = (isExpected ? "expected" : "unexpected") + " error occurred: \n" + getPrintErrorMessage(error);
  const errorMessageSlack = getCurrent_yyyy_mm_dd_hh_mm_ss() + ": " + errorMessage;
  console.error(errorMessage);
  postSlackAlertChannel(errorMessageSlack, needsMention);
}

// error 情報を出力フォーマットに整形して返す
function getPrintErrorMessage(error){
  return "[name]       " + error.name + "\n" +
         "[message]    " + error.message + "\n" +
         "[stacktrace] " + error.stack;
}

// cell から明細を取得して明細 object の配列に parse して返す
function fetchDetailsFromCell(targetSheetName, targetRangeName, parseDetailFunc){
  const rawDetailsValue = getThisSpreadSheet().getSheetByName(targetSheetName).getRange(targetRangeName).getValue();
  if(rawDetailsValue == ""){
    return null;
  }

  return rawDetailsValue.split("¥n").map(
    row => {
      const rowElems = row.split("#&#");
      return parseDetailFunc(rowElems);
    }
  );
}

// 表の各列の名前から明細を取得して明細 object の配列に parse して返す
// (各列において空白行を trim するので、終端までの行において空白行を含む列を許容しない)
function fetchDetailsFromColumns(columnRangeNames, parseDetailFunc){
  const colValueArrays = columnRangeNames.map(colName => getTrimmedColumnValues(getSettingSheet(), colName));
  const rowCount = colValueArrays[0].length;
  if(rowCount == 0){
    return null;
  }
  return getIntRangeFromZero(rowCount).map(i => parseDetailFunc(colValueArrays, i));
}
