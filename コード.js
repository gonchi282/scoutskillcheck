// サイドバー
function showSidebar() {
//  var htmlOutput = HtmlService.createHtmlOutputFromFile("form");
//  SpreadsheetApp.getUi().showSidebar(htmlOutput);
  var htmlOutput = HtmlService.createTemplateFromFile("form2").evaluate().setTitle("選択");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// ダイアログ
function showDialog() {
  var htmlOutput = HtmlService.createTemplateFromFile("form2").evaluate().setTitle("abcde");
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Dialog");
}

// シート名取得
function getSheetsName() {
  var sheetlist = SpreadsheetApp.getActive().getSheets();
  var sheetNameList = [];
  for(var i = 0; i < sheetlist.length; i++) {
    var sheetName = { id: "sheet"+ String(i), name: sheetlist[i].getSheetName() };
    sheetNameList.push(sheetName);
  }
  return sheetNameList;
}

// doSomething
function doSomething(targetSheets) {
  var activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  targetSheets.forEach(sheetName => {
    var sheet = activeSpreadSheet.getSheetByName(sheetName);
    sheet.getRange(1, 1).setValue(1);
  });
}

// 野営章項目数
function getCampItemCnt() {
  return 7;
}

// 野営管理章項目数
function getMngCampItemCnt() {
  return 7;
}

// 救急章項目数
function getFirstAidItemCnt() {
  return 3;
}

// 野外炊飯章項目数
function getCookItemCnt() {
  return 7;
}

// 公民章項目数
function getCitizenItemCnt() {
  return 8;
}

// パイオニアリング章項目数
function getPioneerItemCnt() {
  return 7;
}

// リーダーシップ章項目数
function getLeaderItemCnt() {
  return 5;
}

// ハイキング章項目数
function getHikeItemCnt() {
  return 8;
}

// スカウトソング章項目数
function getSongItemCnt() {
  return 4;
}

// 通信章項目数
function getCommItemCnt() {
  return 5;
}

// 計測章項目数
function getMeasItemCnt() {
  return 8;
}

// 観察章項目数
function getObsrvItemCnt() {
  return 6;
}

// ログ
function SetLog(msg) {
  Logger.log(msg);
}

// メッセージ表示
function MsgBox(msg) {
  Browser.msgBox(msg);
}
