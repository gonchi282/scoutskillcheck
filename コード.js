function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("認定する", "showSidebar")
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

// サイドバー
function showSidebar() {
//  var htmlOutput = HtmlService.createHtmlOutputFromFile("form");
//  SpreadsheetApp.getUi().showSidebar(htmlOutput);
  var htmlOutput = HtmlService.createTemplateFromFile("form").evaluate().setTitle("選択");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// ダイアログ
function showDialog() {
  var htmlOutput = HtmlService.createTemplateFromFile("form").evaluate().setTitle("abcde");
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

// 対象者の対象のアイテムの「認定した人」、「日付」に入力する
function certScoutSkill(targetList, data, certpsn, date) {
  const certpsnCol = 3;  // 「認定した人」列
  const dateCol = 4;     // 「日付」列
  const activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const startRow = {"campList" : 4,
                    "mngCampList" : 15,
                    "fistAidList" : 25,
                    "cookList" : 31,
                    "citizenList" : 41,
                    "pioneerList" : 52,
                    "leaderList" : 62,
                    "hikeList" : 70,
                    "songList" : 81,
                    "commList" : 90,
                    "measList" : 98,
                    "obsrvList" : 109};
  
  // 対象者のシートに「認定した人」、「日付」を入力
  targetList.forEach(sheetName => {
    var sheet = activeSpreadSheet.getSheetByName(sheetName);
    Object.keys(startRow).forEach(key => {
      for(var row = startRow[key]; row < startRow[key] + data[key].length; row++) {
        var itemIndex = row - startRow[key];
        if (data[key][itemIndex]) {
          var cellpsn = sheet.getRange(row, certpsnCol, 1, 2);
          cellpsn.setValues([[certpsn, date]]);
        }
      }
    });
  });
  
}

// 野営章項目数
function getCampItemCnt() {
  return 8;
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
