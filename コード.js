// Copyright © 2020 Koki Imai
// 技能章チェック用スクリプト

// スプレッドシートオープン
// スプレッドシートを開いたらアドオンメニューに
// このスクリプトを呼び出す項目を追加する
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("認定する", "showSidebar")
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

// サイドバー表示
function showSidebar() {
//  var htmlOutput = HtmlService.createHtmlOutputFromFile("form");
//  SpreadsheetApp.getUi().showSidebar(htmlOutput);
  var htmlOutput = HtmlService.createTemplateFromFile("form").evaluate().setTitle("選択");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// ダイアログ表示
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

// 各項目スタート行取得
// 将来的に動的に取得できるようにしたい
function getStartRowList() {
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
  return startRow;
}

// 対象者の対象のアイテムの「認定した人」、「日付」に入力する
function certScoutSkill(targetList, data, certpsn, date) {
  const certpsnCol = 3;  // 「認定した人」列
  const dateCol = 4;     // 「日付」列
  const activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const startRow = getStartRowList();
  var notChangedDataObject = {};
  var notChangedItemObject;
  var notChangedItemIndexArray;
  
  // 対象者のシートに「認定した人」、「日付」を入力
  targetList.forEach(sheetName => {
    var sheet = activeSpreadSheet.getSheetByName(sheetName);
    // 大項目ごとの処理
    notChangedItemObject = {};
    Object.keys(startRow).forEach(key => {
      notChangedItemIndexArray = [];
      // チェックするアイテムごとの処理
      for(var row = startRow[key]; row < startRow[key] + data[key].length; row++) {
        var itemIndex = row - startRow[key];
        // チェックされたデータを更新
        if (data[key][itemIndex]) {
          var cellpsn = sheet.getRange(row, certpsnCol, 1, 2);
          var values = cellpsn.getValues();
          // 「認定した人」と「日付」が空白の場合のみ更新する
          if (!values[0][0] && !values[0][1]) {
            cellpsn.setValues([[certpsn, date]]);
          // すでに入力されている場合は変更しないデータとして蓄積する
          } else {
            notChangedItemIndexArray.push(itemIndex);
          }
        }
      }
      notChangedItemObject[key] = notChangedItemIndexArray;
      notChangedItemIndexArray = null;
    });
    notChangedDataObject[sheetName] = notChangedItemObject;
    notChangedItemObject = null;
  });

  return notChangedDataObject;
}

// ツールチップ用文字列を取得する
function getToolTipStr(itemId, itemIndex) {
  const descCol = 2;
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = getStartRowList();
  
  // セルからツールチップ用文字列を取得する
  var cell = activeSheet.getRange(startRow[itemId] + itemIndex, descCol);
  var tips = cell.getValue();
  
  return tips;
}

// ログ
function SetLog(msg) {
  Logger.log(msg);
}

// メッセージ表示
function MsgBox(msg) {
  Logger.log(msg);
  Browser.msgBox(msg);
}
