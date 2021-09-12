/** @OnlyCurrentDoc */

let thisSpreadSheetHoge;
function getThisSpreadSheetHoge(){
  if(thisSpreadSheetHoge) return thisSpreadSheetHoge;
  thisSpreadSheetHoge = SpreadsheetApp.getActiveSpreadsheet();
  return thisSpreadSheetHoge;
}

let settingSheetHoge;
function getSettingSheetHoge(){
  if(settingSheetHoge) return settingSheetHoge;
  settingSheetHoge = getThisSpreadSheetHoge().getSheetByName("settings");
  return settingSheetHoge;
}

let phantomJsKeyHoge;
function getPhantomJsKeyHoge(){
  if(phantomJsKeyHoge) return phantomJsKeyHoge;
  phantomJsKeyHoge = getSettingSheetHoge().getRange("phantom_js_key").getValue();
  return phantomJsKeyHoge;
}

// module.exports = function testGetPhantomJsKeyHoge(){
//   return getPhantomJsKeyHoge();
// }

function consoleTestGetPhantomJsKeyHoge(){
  console.log(getPhantomJsKeyHoge());
}
