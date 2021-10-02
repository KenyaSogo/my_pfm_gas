/** @OnlyCurrentDoc */

// 現在の日付 (yyyy/MM/dd 形式) を返す
function getCurrent_yyyy_MM_dd(){
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
}

// 現在の時刻 (yyyy/MM/dd HH:mm:ss 形式) を返す
function getCurrent_yyyy_mm_dd_hh_mm_ss(){
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
}

// 現在の日付の年 (yyyy 形式) を返す
function getCurrentYyyy(){
  return getYyyyFrom(getIsTestEnv() ? getTestCalenderDate() : new Date());
}

// 現在の日付の月 (MM 形式) を返す
function getCurrentMm(){
  return getMmFrom(getIsTestEnv() ? getTestCalenderDate() : new Date());
}

// 現在の日付を返す
function getCurrentDate(){
  return getIsTestEnv() ? getTestCalenderDate() : new Date();
}

// Date を yyyy 形式の文字列にして返す
function getYyyyFrom(date){
  return Utilities.formatDate(date, "Asia/Tokyo", "yyyy");
}

// Date を MM 形式の文字列にして返す
function getMmFrom(date){
  return Utilities.formatDate(date, "Asia/Tokyo", "MM");
}

// Date を yyyy/MM/dd 形式の文字列にして返す
function getYyyyMmDdFrom(date){
  return Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd");
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

// Date の当月末を返す
function getCurrentMonthEndDate(date){
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

// Date の n 日後を返す
function getDateLaterFrom(date, laterNum){
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + laterNum);
}

// 指定年月に対して n ヶ月前の年および月を返す
function getYearMonthMonthsAgoFrom(targetYear, targetMonth, agoNum){
  const currentMonthDate = new Date([targetYear, targetMonth, "01"].join("/"));
  const dateSomeMonthsAgo = getDateMonthsAgoFrom(currentMonthDate, agoNum);
  const someMonthsAgoYyyy = getYyyyFrom(dateSomeMonthsAgo);
  const someMonthsAgoMm = getMmFrom(dateSomeMonthsAgo);

  return {someMonthsAgoYyyy, someMonthsAgoMm};
}

// 指定年月に対して前月の年および月を返す
function getPreviousYearMonth(targetYear, targetMonth){
  const currentMonthDate = new Date([targetYear, targetMonth, "01"].join("/"));
  const previousMonthDate = getPreviousMonthDate(currentMonthDate);
  const previousMonthYyyy = getYyyyFrom(previousMonthDate);
  const previousMonthMm = getMmFrom(previousMonthDate);

  return {previousMonthYyyy, previousMonthMm};
}

// 指定年月に対して翌月の年および月を返す
function getNextYearMonth(targetYear, targetMonth){
  const currentMonthDate = new Date([targetYear, targetMonth, "01"].join("/"));
  const nextMonthDate = getNextMonthDate(currentMonthDate);
  const nextMonthYyyy = getYyyyFrom(nextMonthDate);
  const nextMonthMm = getMmFrom(nextMonthDate);

  return {nextMonthYyyy, nextMonthMm};
}

// 指定年月の月初から月末までの date 配列を返す
function getMonthlyDateArrayAt(targetYear, targetMonth){
  const currentMonthStartDate = new Date([targetYear, targetMonth, "01"].join("/"));
  const currentMonthEndDate = getCurrentMonthEndDate(currentMonthStartDate);
  const dateNum = currentMonthEndDate.getDate() - currentMonthStartDate.getDate() + 1;
  return getIntRangeFromZero(dateNum).map(i => getDateLaterFrom(currentMonthStartDate, i));
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

// 配列の最後の要素を取得して返す
function getLastElemFrom(array, defaultValue){
  return array.length == 0 ? defaultValue : array[array.length - 1];
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
function postSlackReportingChannel(message, imageUrls){
  if(getIsTestEnv()){
    getTestReportResultGotRange().setValue(message);
    return;
  }
  postSlackToWithImageUrlAttachments(getSlackReportingBotWebhookUrl(), convertSlackCodeFormatMessage(message), imageUrls);
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

// slack の webHookUrl に対して、画像URLを添付してポストする
function postSlackToWithImageUrlAttachments(webHookUrl, message, imageUrls){
  const payloadObj = {
    text: message,
    attachments: imageUrls.map(
      imageUrl => {
        return {
          fields: [
            {
              title: imageUrl.title,
              value: imageUrl.value,
            }
          ],
          image_url: imageUrl.url,
        };
      }
    ),
  };

  const options = {
    method:      "post",
    contentType: "application/json",
    payload:     JSON.stringify(payloadObj),
  };
  UrlFetchApp.fetch(webHookUrl, options);
}

// console と slack に job 開始メッセージをポストする
let jobProcessStartTime;
function postConsoleAndSlackJobStart(message){
  jobProcessStartTime = Date.now();
  postConsoleAndSlackJobStartOrEnd("start", message);
}

// console と slack に job 終了メッセージをポストする
function postConsoleAndSlackJobEnd(message){
  if(jobProcessStartTime){
    const processingTime = Date.now() - jobProcessStartTime;
    // 処理時間が90秒を超えた場合には warning を出しておく
    if(processingTime > 90000){
      postConsoleAndSlackWarning("job processing time exceeded 90 sec", true);
    }
    message += " (" + processingTime/1000 + " sec)";
  }
  postConsoleAndSlackJobStartOrEnd("end", message);
}

// console と slack に job 開始もしくは終了メッセージをポストする
function postConsoleAndSlackJobStartOrEnd(startOrEndString, message){
  const infoMessage = startOrEndString + ": " + message;
  postConsoleAndSlackInfoMessage(infoMessage);
}

// console と slack (logging channel) に info メッセージをポストする
function postConsoleAndSlackInfoMessage(message){
  console.log(message);
  const postMessageSlack = getCurrent_yyyy_mm_dd_hh_mm_ss() + ": " + message;
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

// 想定エラーをハンドリングする
function handleExpectedError(error){
  // 想定エラーの場合、メッセージをポストした後そのまま抜ける
  if(error instanceof PfmExpectedError){
    postConsoleAndSlackExpectedError(error, error.needsMention);
    return;
  }

  // 想定外エラーの場合、そのまま error を throw して異常終了させる (より上流のハンドラにハンドリングを委ねる)
  throw error;
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
    return [];
  }

  return rawDetailsValue.split("¥n").map(
    row => {
      const rowElems = row.split("#&#");
      return parseDetailFunc(rowElems);
    }
  );
}

// setting sheet における表 (含まれる各列の名前を指定) を明細 object の配列に parse して返す
function fetchSettingDetailsFromColumns(columnRangeNames, parseDetailFunc){
  return fetchDetailsFromColumns(getSettingSheet(), columnRangeNames, parseDetailFunc);
}

// 表の各列の名前から明細を取得して明細 object の配列に parse して返す
// (各列において空白行を trim するので、終端までの行において空白行を含む列を許容しない)
function fetchDetailsFromColumns(targetSheet, columnRangeNames, parseDetailFunc){
  const colValueArrays = columnRangeNames.map(colName => getTrimmedColumnValues(targetSheet, colName));
  const rowCount = colValueArrays[0].length;
  if(rowCount == 0){
    return null;
  }
  return getIntRangeFromZero(rowCount).map(i => parseDetailFunc(colValueArrays, i));
}

// 配列から空白項目を除外して返す
function omitBlankElemFromArray(array){
  return array.filter(r => r);
}

// 結果明細を export 対象シート (月次) に貼り付ける
function exportMonthlyResultDetails(resultDetails, targetYear, targetMonth, sheetPrefix, pasteAddr){
  console.log("exportMonthlyResultDetails: targetYear, targetMonth: " + [targetYear, targetMonth].join(", "));
  exportResultDetails(resultDetails, sheetPrefix + "_" + targetYear + targetMonth, pasteAddr)
}

// 結果明細を export 対象シートに貼り付ける
function exportResultDetails(resultDetails, sheetName, pasteAddr){
  console.log("exportResultDetails: start");

  // 集計結果の明細を区切り文字で結合し、貼り付け用に一つの文字列にする
  const mergedResultDetails = resultDetails.map(s => Object.values(s).join("#&#")).join("¥n");

  // 貼り付け対象のセルを取得
  const pasteTargetCell = getThisSpreadSheet().getSheetByName(sheetName).getRange(pasteAddr);

  // 対象データにつき、更新がなければ、貼り付けをスキップして終了
  const currentValue = pasteTargetCell.getValue();
  if(mergedResultDetails == currentValue){
    console.log("exportResultDetails: end: result details has no update: update skipped");
    return;
  }

  // データを対象セルに貼り付ける
  pasteTargetCell.setValue(mergedResultDetails);
  console.log("exportResultDetails: end: export target cell was updated");
}

// URL を叩いて rawData を取得する
let testAggreRawDataDetailGetCount = 0;
function fetchRawDataDetailFromUrl(urlDetail, options){
  if(getIsTestEnv()){
    testAggreRawDataDetailGetCount++;
    return testAggreRawDataDetailGetCount == 1 ? getTestAggreRawDataDetail1() : getTestAggreRawDataDetail2();
  }

  const responseDetail = JSON.parse(
    UrlFetchApp.fetch(urlDetail, options)
      .getContentText('UTF-8'));
  return JSON.stringify(responseDetail.content.data)
    .trim()
    .replace(/\\/g, '')
    .replace(/\" \"/g, '\"\\n\"');
}

// Range の Value を取得する
function getRangeValueFrom(targetSheet, TargetRangeName){
  return getRangeFrom(targetSheet, TargetRangeName).getValue();
}

// Range を取得する
function getRangeFrom(targetSheet, TargetRangeName){
  return targetSheet.getRange(getIsTestEnv() ? TargetRangeName + "_t" : TargetRangeName);
}

// 列の name の配列を取得する
function getColumnRangeNames(names){
  return names.map(n => getIsTestEnv() ? n + "_t" : n);
}

// exec name から func name を取得し、その関数を実行する
function execFuncFrom(execName){
  Function(getExecFuncNameFrom(execName) + "()")();
}
