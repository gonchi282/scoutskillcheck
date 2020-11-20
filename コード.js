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
  var sheetsNameList = [];
  for(var i = 0; i < sheetlist.length; i++) {
    var sheetName = { id: "sheet"+ String(i), name: sheetlist[i].getSheetName() };
    sheetsNameList.push(sheetName);
  }
  return sheetsNameList;
}