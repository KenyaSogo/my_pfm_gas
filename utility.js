/** @OnlyCurrentDoc */

// 現在の日付 (yyyy/MM/dd 形式) を返す
function getCurrent_yyyy_MM_dd(){
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
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

// 0..elemNum の range (ex. elemNum=3 の場合 [0, 1, 2, 3]) を返す
function getIntRangeFromZero(elemNum){
  return Array.from(Array(elemNum).keys());
}

// 配列の値を一意に集約して返す
function getUniqueArrayFrom(array){
  return Array.from(new Set(array));
}
