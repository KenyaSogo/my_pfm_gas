/** @OnlyCurrentDoc */

// 最新月分の明細を取得する
function executeAggregateCurrentMonth() {
  aggregateByMonth(0);
}

// 最新月の前月分の明細を取得する
function executeAggregatePreviousMonth() {
  aggregateByMonth(1);
}

// 指定月の入出金明細を取得し、対象シートに貼り付ける
function aggregateByMonth(monthsAgo) {
  const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("settings");
  const phantomJsKey = settingSheet.getRange("phantom_js_key").getValue();

  // aggregate する対象の月日を取得する
  const {aggregateYearStack, aggregateMonthStack} = getAggregateYearMonths(settingSheet);
  const aggregateMonthLength = aggregateMonthStack.length;
  const targetYear = aggregateYearStack[aggregateMonthLength - 1 - monthsAgo];
  const targetMonth = aggregateMonthStack[aggregateMonthLength - 1 - monthsAgo];
  
  // aggregate 対象 pfm アカウントの id / pass リストを取得
  const {aggregateIdStack, aggregatePassStack} = getAggregateIdPasses(settingSheet);
  const aggregateAccountLength = aggregateIdStack.length;
  const pasteTargetSheetName = "ag_" + targetYear + targetMonth;
  // 各 pfm アカウント毎に入出金明細をスクレイピング
  for(let i = 0; i < aggregateAccountLength; i++){
    // スクレイピング job をキックする
    const targetId = aggregateIdStack[i];
    const targetPass = aggregatePassStack[i];
    console.log("scrapeCashFlowDataDetail: start");
    const rawData = scrapeCashFlowDataDetail(targetId, targetPass, phantomJsKey, targetYear, targetMonth);
    console.log("scrapeCashFlowDataDetail: done");

    // rawData が想定通り取得出来ているかを確認する
    if(!validateRawData(settingSheet, rawData)){
      console.error("failed to validateRawData: rawData: " + rawData);
      throw new Error("failed to validateRawData");
    }
    console.log("validateRawData: done");
    
    // 月別に用意された貼り付け対象シートに raw データを貼り付ける
    const aggregateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(pasteTargetSheetName);
    const pasteTargetCell = aggregateSheet.getRange("A" + (i + 2));
    // raw データの更新有無を判定し、更新有りなら貼り付ける
    if(isUpdatedRawData(rawData, pasteTargetCell)){
      pasteTargetCell.setValue(rawData);
      console.log("rawData is updated");
    } else {
      console.log("rawData is not updated");
    }
  }
}

// rawData に更新があるかを判定する
function isUpdatedRawData(rawData, pasteTargetCell){
  const pastedValue = pasteTargetCell.getValue();
  // rawData が現在貼り付けられているデータと前方一致していない場合、もしくは現在値が空白の場合に、更新ありと判断する (pastedValue は全量取得できないので前方一致で判定する)
  return rawData.indexOf(pastedValue) != 0 || pastedValue == "";
}

// rawData が想定通りのパターンかを validate する
function validateRawData(settingSheet, rawData){
  const rawDataHeadPattern = settingSheet.getRange("raw_data_head_pattern").getValue();
  // rawData の先頭文字列が rawDataHeadPattern に合致しているなら valid (= true)
  return rawData.indexOf(rawDataHeadPattern) == 0;
}

// aggregate する対象の id と pass を stack させた配列を返す
function getAggregateIdPasses(settingSheet){
  // 空白行以降の行を切り捨てた配列を作り、返す
  const aggregateIds = settingSheet.getRange("pfm_id").getValues();
  const aggregatePasses = settingSheet.getRange("pfm_pass").getValues();

  const aggregateIdStack = [];
  const aggregatePassStack = [];
  let i = 1;
  while(true){
    let currentId = aggregateIds[i][0];
    let currentPass = aggregatePasses[i][0];
    
    if(currentId == "" || currentPass == ""){
      break;
    }

    aggregateIdStack.push(currentId);
    aggregatePassStack.push(currentPass);
    i++;
  }

  return {aggregateIdStack, aggregatePassStack};
}

// aggregate する対象の year と month を stack させた配列を返す
function getAggregateYearMonths(settingSheet){
  // 空白行以降の行を切り捨てた配列を作り、返す
  const aggregateYears = settingSheet.getRange("aggregate_year").getValues();
  const aggregateMonths = settingSheet.getRange("aggregate_month").getValues();

  const aggregateYearStack = [];
  const aggregateMonthStack = [];
  let i = 1;
  while(true){
    let currentYear = aggregateYears[i][0];
    let currentMonth = aggregateMonths[i][0];
    
    if(currentYear == "" || currentMonth == ""){
      break;
    }

    aggregateYearStack.push(currentYear);
    aggregateMonthStack.push(currentMonth);
    i++;
  }

  return {aggregateYearStack, aggregateMonthStack};
}

// 指定された id/pass に紐づく pfm アカウントの、指定月の明細を取得して返す
function scrapeCashFlowDataDetail(loginId, loginPassword, phantomJsKey, targetYear, targetMonth) {
  var scrapeUrl = "https://moneyforward.com/cf/csv?from=" + targetYear + "%2F" + targetMonth + "%2F25&month=" + targetMonth + "&year=" + targetYear;

  var payloadDetail = {
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
  
  let urlDetail = 'https://PhantomJsCloud.com/api/browser/v2/'+ phantomJsKey +'/';
  let responseDetail = JSON.parse(UrlFetchApp.fetch(urlDetail, options).getContentText('UTF-8'));
  
  let rawDataDetail = JSON.stringify(responseDetail.content.data).trim().replace(/\\/g, '').replace(/\" \"/g, '\"\\n\"');
  
  return rawDataDetail;
}

// aggregate_year / aggregate_month のリストを更新する (これによって最新月/その前月として aggregation しに行く月の設定が切り替わる)
function executeUpdateAggregateYearMonthList(){
  const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("settings");

  // 現在の日付を設定
  const currentDate = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
  settingSheet.getRange("today_for_calender").setValue(currentDate);

  // 本日が月初めの日か否かを判定
  const isTodayMonthlyStart = settingSheet.getRange("is_today_monthly_start").getValue() == 1;
  if(!isTodayMonthlyStart){
    console.log("isTodayMonthlyStart: false: push skipped");
    return;
  }

  // 本日が月初めの場合: aggregate_year / aggregate_month のリストの末尾に本日の月日を追加する
  console.log("isTodayMonthlyStart: true");
  
  // リスト末尾のセル、およびその一つ後ろのセル (= push する対象のセル) を取得を取得しておく
  const {aggregateMonthStack} = getAggregateYearMonths(settingSheet);
  const stackedMonthLength = aggregateMonthStack.length;
  const lastCellYear = settingSheet.getRange("aggregate_year").getCell(stackedMonthLength + 1, 1);
  const lastCellMonth = settingSheet.getRange("aggregate_month").getCell(stackedMonthLength + 1, 1);
  const pushTargetCellYear = settingSheet.getRange("aggregate_year").getCell(stackedMonthLength + 2, 1);
  const pushTargetCellMonth = settingSheet.getRange("aggregate_month").getCell(stackedMonthLength + 2, 1);
  
  // 本日の月日が追加未済であるかを確認する
  const currentYear = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy");
  const currentMonth = Utilities.formatDate(new Date(), "Asia/Tokyo", "MM");
  const lastYearMonth = lastCellYear.getValue() + lastCellMonth.getValue();
  const currentYearMonth = currentYear + currentMonth;
  if(lastYearMonth == currentYearMonth){
    console.log("currentYearMonth is already pushed: push skipped");
    return;
  }

  // 追加未済の場合、本日の月日を追加
  console.log("currentYearMonth is not pushed yet");

  pushTargetCellYear.setValue(currentYear);
  pushTargetCellMonth.setValue(currentMonth);
  console.log("aggregate_year, aggregate_month: pushed to the list: year, month: " + currentYear + ", " + currentMonth);  
}

// 月別のシートを作成する
function executeCreateSheetByMonth(){
  const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = currentSpreadSheet.getSheetByName("settings");

  // 現在の日付を設定
  const currentDate = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
  settingSheet.getRange("today_for_calender").setValue(currentDate);

  // 本日が月初めの日の前日か否かを判定
  const isTomorrowMonthlyStart = settingSheet.getRange("is_tomorrow_monthly_start").getValue() == 1;
  if(!isTomorrowMonthlyStart){
    console.log("isTomorrowMonthlyStart: false: sheet create skipped");
    return;
  }

  // 本日が月初め日の前日の場合: 当月分のシートを tmp シートからコピーして追加
  console.log("isTomorrowMonthlyStart: true");

  // シートが既にコピー済みか否かを判定
  const currentYearMonth = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMM");
  const newSheetName = "ag_" + currentYearMonth;
  const existingNewSheet = currentSpreadSheet.getSheetByName(newSheetName);
  if(existingNewSheet){
    console.log("newSheet is already created: sheet create skipped: newSheetName: " + newSheetName);
    return;
  }

  // コピー未済の場合、当月分のシートをコピーして追加
  console.log("newSheet is not created yet");

  const tmpSheet = currentSpreadSheet.getSheetByName("ag_tmp");
  const copiedNewSheet = tmpSheet.copyTo(currentSpreadSheet);
  copiedNewSheet.setName(newSheetName);
  console.log("newSheet is created: sheet name: " + newSheetName);
}

// シートのメンテナンス状態をチェックする
function executeCheckSheetMaintenanceStatus(){
  const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("settings");
  
  // 休日カレンダーのメンテナンス要否を確認する
  // (現状数年分作っているが、時期が来て足さないといけなくなったタイミングで通知を飛ばす ※祝日のフラグを手入力する以外は関数の拡張のみ)
  const isCalenderNeedsMaitenance = settingSheet.getRange("is_calender_needs_maitenance").getValue() == 1;
  if(isCalenderNeedsMaitenance){
    console.log("isCalenderNeedsMaitenance: true: notified to slack");
    // TODO: slack 通知
  } else {
    console.log("isCalenderNeedsMaitenance: false");
  }

  // 月初めの日 (基本だが25日、その日が休日だと営業日ベースに前倒しさせた日) の設定が正しいか確認する
  // (現状、6連休の最後の日が25日、というパターンまでしか考慮しておらず、その限界を超えるケースの有無をチェックする)
  const isMonthlyStartDayInvalid = settingSheet.getRange("is_monthly_start_day_invalid").getValue() == 1;
  if(isMonthlyStartDayInvalid){
    console.log("isMonthlyStartDayInvalid: true: notified to slack");
    // TODO: slack 通知
  } else {
    console.log("isMonthlyStartDayInvalid: false");
  }
}

// 最新月分の日次集計を行う
function executeCalcDailySummaryCurrentMonth(){
  calcDailySummary(0);
}

// 最新月の前月分の日次集計を行う
function executeCalcDailySummaryPreviousMonth(){
  calcDailySummary(1);
}

// 指定月の日次集計を行う
function calcDailySummary(monthsAgo){
  const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("settings");

  // 集計対象月を取得する
  const {aggregateYearStack, aggregateMonthStack} = getAggregateYearMonths(settingSheet);
  const aggregateMonthLength = aggregateMonthStack.length;
  const targetYear = aggregateYearStack[aggregateMonthLength - 1 - monthsAgo];
  const targetMonth = aggregateMonthStack[aggregateMonthLength - 1 - monthsAgo];
  console.log("targetYear, targetMonth: " + targetYear + ", " + targetMonth);

  // 集計対象月の明細データを hash の配列に parse する
  const targetAggreSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ag_" + targetYear + targetMonth);
  const allRowsData = targetAggreSheet.getRange("V2").getValue();
  const aggregatedDetails = allRowsData.split("¥n").map(
    row => {
      const rowElems = row.split("#&#");
      return {
        isCalcTarget:       rowElems[0], // 計算対象
        date:               rowElems[1], // 日付
        contents:           rowElems[2], // 内容
        amount:             rowElems[3], // 金額(円)
        assetName:          rowElems[4], // 保有金融機関
        largeCategoryName:  rowElems[5], // 大項目
        middleCategoryName: rowElems[6], // 中項目
        memo:               rowElems[7], // メモ
        isTransfer:         rowElems[8], // 振替
        uuid:               rowElems[9], // ID
      };
    }
  );
  console.log("parse aggregatedDetails finished: length: " + aggregatedDetails.length);

  // 明細データのうち現預金の入出金を集計し、日次の残高推移を求める TODO: 集計系ロジックを function に切り出して整理
  // 現預金に該当する保有金融機関名を取得する
  const assetConfigs = getAssetConfigs(settingSheet);
  const cashAssetNames = assetConfigs.filter(a => a.isCash == 1).map(a => a.assetName);
  // 明細データを集計対象のものに絞り込む ※現預金の口座に限定して動きを見るため振替も含める
  const balanceCashTargetDetails = aggregatedDetails
                                     .filter(d => d.isCalcTarget == 1 || d.isTransfer == 1)
                                     .filter(d => cashAssetNames.includes(d.assetName));
  // 集計対象の日付を収集 (一度 Set にすることで重複を排除している)
  const balanceCashTargetDates = Array.from(new Set(balanceCashTargetDetails.map(d => d.date))); // TODO: 日が飛び飛びになる場合に間の残高を横置きで埋める考慮
  balanceCashTargetDates.sort(); // 昇順にソート
  // 前月末残高を取得する
  const {previousMonthYyyy, previousMonthMm} = getPreviousYearMonth(targetYear, targetMonth);
  const endBalances = getEndBalances(settingSheet);
  const previousMonthEndBalance = endBalances.find(b => b.calcYear == previousMonthYyyy && b.calcMonth == previousMonthMm);
  const previousMonthEndBalanceCash = Number(previousMonthEndBalance.cashEndBalance);
  // 日次で現預金残高を集計
  let accumulatedBalance = previousMonthEndBalanceCash;
  const cashBalanceByDates = balanceCashTargetDates.map(
    date => {
      // 対象日の明細を抽出
      const targetDateDetails = balanceCashTargetDetails.filter(d => d.date == date);
      // 残高を集計
      const sumAmountAtTargetDate = targetDateDetails.map(d => Number(d.amount)).reduce((sum, amount) => sum + amount, 0);
      accumulatedBalance += sumAmountAtTargetDate;
      return {
        date:    date,
        balance: accumulatedBalance,
      };
    }
  );
  console.log("calc cashBalanceByDates finished: length: " + cashBalanceByDates.length);

  // 対象月の末残高設定を更新する
  const cashBalanceAtMonthEnd = cashBalanceByDates.length == 0 ? previousMonthEndBalanceCash : cashBalanceByDates[cashBalanceByDates.length - 1].balance;
  const targetMonthIndexAtEndBalances = endBalances.map(b => b.calcYear + "/" + b.calcMonth).indexOf(targetYear + "/" + targetMonth);
  settingSheet.getRange("cash_end_balance").getCell(targetMonthIndexAtEndBalances + 2, 1).setValue(cashBalanceAtMonthEnd);
  console.log("cash end balance updated: cashBalanceAtMonthEnd: " + cashBalanceAtMonthEnd);

  // 集計した明細を、早い月分と遅い月分に振り分ける
  const earlierYearMonth = targetYear + "/" + targetMonth;
  const earlierMonthCashBalanceByDates = cashBalanceByDates.filter(s => s.date.indexOf(earlierYearMonth) == 0);
  const laterMonthCashBalanceByDates = cashBalanceByDates.filter(s => s.date.indexOf(earlierYearMonth) != 0);

  // 集計結果を結果シートに貼り付け
  // 早い月分
  if(earlierMonthCashBalanceByDates.length > 0){
    console.log("start to paste earlierMonthCashBalanceByDates: earlierMonthCashBalanceByDates.length: " + earlierMonthCashBalanceByDates.length);
    pasteDailyCalcResult(earlierMonthCashBalanceByDates, targetYear, targetMonth, true, "calc_dcb");
  } else {
    console.log("pasting earlierMonthCashBalanceByDates was skipped");
  }
  // 遅い月分
  if(laterMonthCashBalanceByDates.length > 0){
    console.log("start to paste laterMonthCashBalanceByDates: laterMonthCashBalanceByDates.length: " + laterMonthCashBalanceByDates.length);
    const earlierMonthDate = new Date(earlierYearMonth + "/01");
    const laterMonthDate = new Date(earlierMonthDate.getFullYear(), earlierMonthDate.getMonth() + 1, 1);
    const laterMonthYyyy = Utilities.formatDate(laterMonthDate, "Asia/Tokyo", "yyyy");
    const laterMonthMm = Utilities.formatDate(laterMonthDate, "Asia/Tokyo", "MM");
    pasteDailyCalcResult(laterMonthCashBalanceByDates, laterMonthYyyy, laterMonthMm, false, "calc_dcb");
  } else {
    console.log("pasting laterMonthCashBalanceByDates was skipped");
  }

  // 明細データのうち全金融機関分の入出金を集計し、日次の全資産分の残高推移を求める TODO: 集計系ロジックを function に切り出して整理
  // 明細データを集計対象のものに絞り込む
  //   ※振替は口座間で相殺されるので無くても良いが、念の為含めておく
  //     -> 相手勘定の無い振替(クレカ引落のカード側の明細が無い)が一部あり、やはり振替は含めない
  const balanceAssetTargetDetails = aggregatedDetails.filter(d => d.isCalcTarget == 1);
  // 集計対象の日付を収集 (一度 Set にすることで重複を排除している)
  const balanceAssetTargetDates = Array.from(new Set(balanceAssetTargetDetails.map(d => d.date))); // TODO: 日が飛び飛びになる場合に間の残高を横置きで埋める考慮
  balanceAssetTargetDates.sort(); // 昇順にソート
  // 前月末残高を取得する
  const previousMonthEndBalanceAsset = Number(previousMonthEndBalance.assetTotalEndBalance);
  // 日次で現預金残高を集計
  accumulatedBalance = previousMonthEndBalanceAsset;
  const assetBalanceByDates = balanceAssetTargetDates.map(
    date => {
      // 対象日の明細を抽出
      const targetDateDetails = balanceAssetTargetDetails.filter(d => d.date == date);
      // 残高を集計
      const sumAmountAtTargetDate = targetDateDetails.map(d => Number(d.amount)).reduce((sum, amount) => sum + amount, 0);
      accumulatedBalance += sumAmountAtTargetDate;
      return {
        date:    date,
        balance: accumulatedBalance,
      };
    }
  );
  console.log("calc assetBalanceByDates finished: length: " + assetBalanceByDates.length);

  // 対象月の末残高設定を更新する
  const assetBalanceAtMonthEnd = assetBalanceByDates.length == 0 ? previousMonthEndBalanceAsset : assetBalanceByDates[assetBalanceByDates.length - 1].balance;
  settingSheet.getRange("asset_total_end_balance").getCell(targetMonthIndexAtEndBalances + 2, 1).setValue(assetBalanceAtMonthEnd);
  console.log("asset total end balance updated: assetBalanceAtMonthEnd: " + assetBalanceAtMonthEnd);

  // 集計した明細を、早い月分と遅い月分に振り分ける
  const earlierMonthAssetBalanceByDates = assetBalanceByDates.filter(s => s.date.indexOf(earlierYearMonth) == 0);
  const laterMonthAssetBalanceByDates = assetBalanceByDates.filter(s => s.date.indexOf(earlierYearMonth) != 0);

  // 集計結果を結果シートに貼り付け
  // 早い月分
  if(earlierMonthAssetBalanceByDates.length > 0){
    console.log("start to paste earlierMonthAssetBalanceByDates: earlierMonthAssetBalanceByDates.length: " + earlierMonthAssetBalanceByDates.length);
    pasteDailyCalcResult(earlierMonthAssetBalanceByDates, targetYear, targetMonth, true, "calc_dab");
  } else {
    console.log("pasting earlierMonthAssetBalanceByDates was skipped");
  }
  // 遅い月分
  if(laterMonthAssetBalanceByDates.length > 0){
    console.log("start to paste laterMonthAssetBalanceByDates: laterMonthAssetBalanceByDates.length: " + laterMonthAssetBalanceByDates.length);
    const earlierMonthDate = new Date(earlierYearMonth + "/01");
    const laterMonthDate = new Date(earlierMonthDate.getFullYear(), earlierMonthDate.getMonth() + 1, 1);
    const laterMonthYyyy = Utilities.formatDate(laterMonthDate, "Asia/Tokyo", "yyyy");
    const laterMonthMm = Utilities.formatDate(laterMonthDate, "Asia/Tokyo", "MM");
    pasteDailyCalcResult(laterMonthAssetBalanceByDates, laterMonthYyyy, laterMonthMm, false, "calc_dab");
  } else {
    console.log("pasting laterMonthAssetBalanceByDates was skipped");
  }
  
  // 明細データを日次で大項目&中項目別に集計する
  // 集計対象明細に絞り込む
  const targetAggregateDetails = aggregatedDetails.filter(d => d.isCalcTarget == 1);
  // 集計対象の日付を収集 (一度 Set にすることで重複を排除している)
  const targetDates = Array.from(new Set(targetAggregateDetails.map(d => d.date)));
  targetDates.sort(); // 昇順にソート
  // 日次で大項目&中項目別に集計
  const summaryByDates = targetDates.map(
    date => {
      // 対象日の明細を抽出
      const targetDateDetails = targetAggregateDetails.filter(d => d.date == date);
      // 対象日における大項目&中項目を重複排除の上抽出
      const categories = Array.from(new Set(targetDateDetails.map(d => d.largeCategoryName + ">>" + d.middleCategoryName)));
      // 対象日の明細において、大項目&中項目別に集計
      return categories.map(
        c => {
          const targetDateAndCategoryDetails = targetDateDetails.filter(d => d.largeCategoryName + ">>" + d.middleCategoryName == c);
          const sumAmountByDateAndCategory = targetDateAndCategoryDetails.map(d => Number(d.amount)).reduce((sum, amount) => sum + amount, 0);
          return {
            date:         date,                       // 日付
            categoryName: c,                          // 項目 (大項目&中項目)
            amount:       sumAmountByDateAndCategory, // 金額
          };
        }
      );
    }
  ).flat(1);
  console.log("summaryByDate finished: length: " + summaryByDates.length);

  // 日次で集計した明細を、早い月分と遅い月分に振り分ける
  const earlierMonthSummaryByDates = summaryByDates.filter(s => s.date.indexOf(earlierYearMonth) == 0);
  const laterMonthSummaryByDates = summaryByDates.filter(s => s.date.indexOf(earlierYearMonth) != 0);

  // 集計結果を結果シートに貼り付け
  // 早い月分
  if(earlierMonthSummaryByDates.length > 0){
    console.log("start to paste earlierMonthSummaryByDates: earlierMonthSummaryByDates.length: " + earlierMonthSummaryByDates.length);
    pasteDailyCalcResult(earlierMonthSummaryByDates, targetYear, targetMonth, true, "calc_dc");
  } else {
    console.log("pasting earlierMonthSummaryByDates was skipped");
  }
  // 遅い月分
  if(laterMonthSummaryByDates.length > 0){
    console.log("start to paste laterMonthSummaryByDates: laterMonthSummaryByDates.length: " + laterMonthSummaryByDates.length);
    const earlierMonthDate = new Date(earlierYearMonth + "/01");
    const laterMonthDate = new Date(earlierMonthDate.getFullYear(), earlierMonthDate.getMonth() + 1, 1);
    const laterMonthYyyy = Utilities.formatDate(laterMonthDate, "Asia/Tokyo", "yyyy");
    const laterMonthMm = Utilities.formatDate(laterMonthDate, "Asia/Tokyo", "MM");
    pasteDailyCalcResult(laterMonthSummaryByDates, laterMonthYyyy, laterMonthMm, false, "calc_dc");
  } else {
    console.log("pasting laterMonthSummaryByDates was skipped");
  }
}

// 保有金融機関表を取得して返す
function getAssetConfigs(settingSheet){
  const assetNameValues = settingSheet.getRange("asset_name").getValues();
  const isCashValues = settingSheet.getRange("is_cash").getValues();

  return Array.from(Array(assetNameValues.filter(v => v != "").length - 1).keys()).map(
    i => {
      return {
        assetName: assetNameValues[i + 1][0],
        isCash:    isCashValues[i + 1][0],
      };
    }
  );
}

// 指定月日に対して前月の月および日を返す
function getPreviousYearMonth(targetYear, targetMonth){
  const currentCalcMonthDate = new Date([targetYear, targetMonth, "01"].join("/"));
  const previousCalcMonthDate = new Date(currentCalcMonthDate.getFullYear(), currentCalcMonthDate.getMonth() - 1, 1);
  const previousMonthYyyy = Utilities.formatDate(previousCalcMonthDate, "Asia/Tokyo", "yyyy");
  const previousMonthMm = Utilities.formatDate(previousCalcMonthDate, "Asia/Tokyo", "MM");

  return {previousMonthYyyy, previousMonthMm};
}

// 月末残高表を取得して返す
function getEndBalances(settingSheet){
  const calcYearValues = settingSheet.getRange("calc_year").getValues();
  const calcMonthValues = settingSheet.getRange("calc_month").getValues();
  const cashEndBalanceValues = settingSheet.getRange("cash_end_balance").getValues();
  const assetTotalEndBalanceValues = settingSheet.getRange("asset_total_end_balance").getValues();
  
  return Array.from(Array(calcYearValues.filter(v => v != "").length - 1).keys()).map(
    i => {
      return {
        calcYear:             calcYearValues[i + 1][0],
        calcMonth:            calcMonthValues[i + 1][0],
        cashEndBalance:       cashEndBalanceValues[i + 1][0],
        assetTotalEndBalance: assetTotalEndBalanceValues[i + 1][0],
      };
    }
  );
}

// 日次集計結果の calc シートへの貼り付けを行う
function pasteDailyCalcResult(summaryByDates, targetYear, targetMonth, isEarlier, sheetPrefix){
  console.log("pasteDailyCalcResult: start: targetYear, targetMonth, isEarlier: " + [targetYear, targetMonth, isEarlier].join(", "));

  // 集計結果の明細を区切り文字で結合し、貼り付け用に一つの文字列にする
  const mergedSummaryByDatesValue = summaryByDates.map(s => Object.values(s).join("#&#")).join("¥n");
  
  // 貼り付け対象のセルを取得
  const targetCalcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetPrefix + "_" + targetYear + targetMonth); // TODO: sheet 作成 job に追加
  const pasteTargetCell = targetCalcSheet.getRange(isEarlier ? "A3" : "A2");
  
  // 対象データにつき、更新がなければ、貼り付けをスキップして終了
  const currentValue = pasteTargetCell.getValue();
  if(mergedSummaryByDatesValue == currentValue){
    console.log("calc result has no update: update skipped");
    return;
  }

  // データを対象セルに貼り付ける
  pasteTargetCell.setValue(mergedSummaryByDatesValue);
  console.log("calc result was updated");
}
