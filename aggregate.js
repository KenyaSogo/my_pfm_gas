/** @OnlyCurrentDoc */

// 最新月分の明細を取得する
function executeAggregateCurrentMonth() {
  postConsoleAndSlackJobStart("execute aggregate current month");
  try {
    aggregateByMonth(0);
  } catch(error){
    handleError(error);
  }
  postConsoleAndSlackJobEnd("execute aggregate current month");
}

// 最新月の前月分の明細を取得する
function executeAggregatePreviousMonth() {
  postConsoleAndSlackJobStart("execute aggregate previous month");
  try {
    aggregateByMonth(1);
  } catch(error){
    handleError(error);
  }
  postConsoleAndSlackJobEnd("execute aggregate previous month");
}

// 指定月の入出金明細を取得し、対象シートに貼り付ける
function aggregateByMonth(monthsAgo) {
  // aggregate する対象の月日を取得する
  const {targetYear, targetMonth} = getAggregateTargetYearAndMonth(monthsAgo);
  console.log("targetYear, targetMonth: " + targetYear + ", " + targetMonth);
  
  // aggregate 対象 pfm アカウントの id / pass リストを取得
  const {aggregateIdStack, aggregatePassStack} = getAggregateIdPasses();
  const aggregateSheet = getThisSpreadSheet().getSheetByName(getAggreImportSheetPrefix() + "_" + targetYear + targetMonth);
  // 各 pfm アカウント毎に入出金明細をスクレイピング
  getIntRangeFromZero(aggregateIdStack.length).forEach(
    i => {
      // スクレイピング job をキックする
      const targetId = aggregateIdStack[i];
      const targetPass = aggregatePassStack[i];
      console.log("scrapeCashFlowDataDetail: start: targetId: " + targetId);
      const rawData = scrapeCashFlowDataDetail(targetId, targetPass, targetYear, targetMonth);
      console.log("scrapeCashFlowDataDetail: done");

      // rawData が想定通り取得出来ているかを確認する TODO: 月初の0件明細エラー対応
      if(!validateRawData(rawData)){
        postConsoleAndSlackWarning("failed to validateRawData: rawData: " + rawData, false);
        return;
      }
      console.log("validateRawData: done");

      // 月別に用意された貼り付け対象シートに raw データを貼り付ける
      const pasteTargetCell = aggregateSheet.getRange(getAggreImportAddrCol() + (i + 2));
      // raw データの更新有無を判定し、更新有りなら貼り付ける
      if(isUpdatedRawData(rawData, pasteTargetCell)){
        pasteTargetCell.setValue(rawData);
        console.log("rawData is updated");
      } else {
        console.log("rawData is not updated");
      }
    }
  );
}

// aggregate 対象の year (yyyy) おもよ month (MM) を取得する
function getAggregateTargetYearAndMonth(monthsAgo){
  const {aggregateYearStack, aggregateMonthStack} = fetchAggregateYearMonths();
  const targetYear = aggregateYearStack[aggregateMonthStack.length - 1 - monthsAgo];
  const targetMonth = aggregateMonthStack[aggregateMonthStack.length - 1 - monthsAgo];

  return {targetYear, targetMonth};
}

// rawData に更新があるかを判定する
function isUpdatedRawData(rawData, pasteTargetCell){
  const pastedValue = pasteTargetCell.getValue();
  // rawData が現在貼り付けられているデータと前方一致していない場合、もしくは現在値が空白の場合に、更新ありと判断する (pastedValue は全量取得できないので前方一致で判定する)
  return rawData.indexOf(pastedValue) != 0 || pastedValue == "";
}

// rawData が想定通りのパターンかを validate する
function validateRawData(rawData){
  // rawData の先頭文字列が rawDataHeadPattern に合致しているなら valid (= true)
  return rawData.indexOf(getRawDataHeadPattern()) == 0;
}

// aggregate する対象の id と pass を stack させた配列を返す TODO: config 移管、他との方式の平仄
function getAggregateIdPasses(){
  const aggregateIdStack = getTrimmedColumnValues(getSettingSheet(), "pfm_id");
  const aggregatePassStack = getTrimmedColumnValues(getSettingSheet(), "pfm_pass");

  return {aggregateIdStack, aggregatePassStack};
}

// aggregate する対象の year と month を stack させた配列を返す TODO: config 移管、他との方式の平仄
function fetchAggregateYearMonths(){
  const aggregateYearStack = getTrimmedColumnValues(getSettingSheet(), "aggregate_year");
  const aggregateMonthStack = getTrimmedColumnValues(getSettingSheet(), "aggregate_month");

  return {aggregateYearStack, aggregateMonthStack};
}

// 指定された id/pass に紐づく pfm アカウントの、指定月の明細を取得して返す
function scrapeCashFlowDataDetail(loginId, loginPassword, targetYear, targetMonth) {
  let scrapeUrl = "https://moneyforward.com/cf/csv?from=" + targetYear + "%2F" + targetMonth + "%2F25&month=" + targetMonth + "&year=" + targetYear;

  let payloadDetail = {
    "url": 'https://moneyforward.com/cf',
    "renderType": 'plainText',
    "outputAsJson": true,
    "overseerScript": "let _user='" + loginId + "';" +
                      "let _pass='" + loginPassword + "';" +
                      "await page.waitForSelector('a._2YH0UDm8.ssoLink');" +
                      "page.click('a._2YH0UDm8.ssoLink');" +
                      "await page.waitForSelector('input._2mGdHllU.inputItem');" +
                      "await page.type('input._2mGdHllU.inputItem', _user, {delay:100});" +
                      "page.click('input.zNNfb322.submitBtn.homeDomain');" +
                      "await page.waitForSelector('input._1vBc2gjI.inputItem');" +
                      "await page.type('input._1vBc2gjI.inputItem', _pass, {delay:100});" +
                      "page.click('input.zNNfb322.submitBtn.homeDomain');" +
                      "await page.waitForSelector('#page-transaction');" +
                      "await page.evaluate(" +
                        "()=>{" +
                          "var xhr = new XMLHttpRequest();" +
                          "xhr.open('GET', '" + scrapeUrl + "');" +
                          "xhr.overrideMimeType('text/plain; charset=Shift_JIS');" +
                          "xhr.send();" +
                          "xhr.onload = function(e) {" +
                            "$('#page-transaction').text(xhr.responseText);" +
                          "};" +
                        "}" +
                      ");" +
                      "page.done();",
  };
  let options = {
    "method": 'post',
    "payload": JSON.stringify(payloadDetail),
  };
  
  let urlDetail = 'https://PhantomJsCloud.com/api/browser/v2/'+ getPhantomJsKey() +'/';
  let responseDetail = JSON.parse(UrlFetchApp.fetch(urlDetail, options).getContentText('UTF-8'));
  
  let rawDataDetail = JSON.stringify(responseDetail.content.data).trim().replace(/\\/g, '').replace(/\" \"/g, '\"\\n\"');
  
  return rawDataDetail;
}
