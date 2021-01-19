// SQL_Statement
const url      = PropertiesService.getScriptProperties().getProperty('urlInfo');
const userName = PropertiesService.getScriptProperties().getProperty('userNameInfo');
const password = PropertiesService.getScriptProperties().getProperty('passwordInfo');
const conn     = Jdbc.getConnection(url, userName, password);
const stmt     = conn.createStatement();
// SS_info
const sheet = SpreadsheetApp.getActiveSheet();
const sheet_setRange = sheet.getRange(1, 1, sheet.getLastRow() + 1, sheet.getLastColumn() + 1)

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Google Cloud SQL").addItem("SQL実行", "openSidebar").addToUi();
}

function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("index").setTitle("Apps Scipt アプリケーション");
  SpreadsheetApp.getUi().showSidebar(html);
}

function sqlRunGas(value) {
  Logger.log(url);
  let studentsData = [];
  let selectID; 
  try {
    if ( value.match(/select/) ) {
      studentsData = stmt.executeQuery(value);
      // スプレッドシート出力
      ssWrite(studentsData);
    } 
    if ( value.match(/update/) ) {
      stmt.execute(value);
      selectID = value.split("=")[2];
      let sqlSelect = 'SELECT * FROM students where id = ' + selectID ;
      studentsData = stmt.executeQuery(sqlSelect);
      // スプレッドシート出力
      ssWrite(studentsData);  
    }
    if ( value.match(/insert/) ) {
      stmt.execute(value);
      // selectID取得
      let tmp1  = value.match(/\(.*\)/);
      let tmp2  = tmp1[0].split(",");
      selectID = tmp2[0]
      selectID = selectID.slice(1);
      let sqlSelect = 'SELECT * FROM students where id = ' + selectID ;
      studentsData = stmt.executeQuery(sqlSelect);
      // スプレッドシート出力
      ssWrite(studentsData);
    }
    if ( value.match(/delete/) ) {
      // delete処理前のselect発行とスプレッドシート出力
      selectID = value.split("=")[1];
      let sqlSelect = 'SELECT * FROM students where id = ' + selectID ;
      studentsData = stmt.executeQuery(sqlSelect);
      ssWrite(studentsData);
      stmt.execute(value);
    }
  } catch (e) {
    throw new Error('SQL実行エラー発生');
  } finally {
    if (conn !== undefined) conn.close();
  }
}

function ssWrite(studentsData) {
  sheet_setRange.clearContent();
  const numCols = studentsData.getMetaData().getColumnCount();
  while (studentsData.next()) {
    const arr = [];
    for (let i = 0; i < numCols; i++) {
      arr.push(studentsData.getString(i + 1));
    }
    sheet.appendRow(arr);
  }
  studentsData.close();
}
