// SQL_Statement
const url      = 'jdbc:mysql://35.188.25.67:3306/h_kusi';
const userName = 'root';
const password = 'root';
const conn     = Jdbc.getConnection(url, userName, password);
const stmt     = conn.createStatement();
// SS_info
const sheet = SpreadsheetApp.getActiveSheet();
const sheet_setRange = sheet.getRange(1, 1, sheet.getLastRow() + 1, sheet.getLastColumn() + 1)


function onOpen() {
  SpreadsheetApp.getUi().createMenu("Google Cloud SQL").addItem("SQL実行", "openSidebar").addToUi();
}


function openSidebar() {
  let html = HtmlService.createHtmlOutputFromFile("index").setTitle("Apps Scipt アプリケーション");
  SpreadsheetApp.getUi().showSidebar(html);
}


function sql_run_gs(value) {
  let students_data = [];
  let strSQL = value;
  let select_id;
  
    try {
      // select処理
      if ( strSQL.match(/select/) ) {
        students_data = stmt.executeQuery(strSQL);
        ss_write(students_data);
        
      // update処理
      } else if ( strSQL.match(/update/) ) {
        stmt.execute(strSQL);
        
        select_id = value.split("=")[2];
        let strSQL_update = 'SELECT * FROM students where id = ' + select_id ;
        students_data = stmt.executeQuery(strSQL_update);
        ss_write(students_data);
        
      // insert処理
      } else if  ( strSQL.match(/insert/) ) {
        stmt.execute(strSQL);
        
        // select_id取得        
        let tmp1  = value.match(/\(.*\)/);
        let tmp2  = tmp1[0].split(",");
        select_id = tmp2[0]
        select_id = select_id.slice(1);
        let strSQL_insert = 'SELECT * FROM students where id = ' + select_id ;
        students_data = stmt.executeQuery(strSQL_insert);
        ss_write(students_data);
              
      // delete処理
      } else {
        // delete処理前のselect発行とスプレッドシート出力
        select_id = value.split("=")[1];
        let strSQL_delete = 'SELECT * FROM students where id = ' + select_id ;
        students_data = stmt.executeQuery(strSQL_delete);
        ss_write(students_data);
        
        stmt.execute(strSQL);
      }
    } catch (e) {
      Logger.log(e);
      console.log(e);
      throw new Error('SQL実行エラー発生');
    } finally {
      if (conn !== undefined) conn.close();
    }
}


function ss_write(students_data) {
  sheet_setRange.clearContent();
  
  let numCols = students_data.getMetaData().getColumnCount();
  while (students_data.next()) {
    let arr = [];
    for (let i = 0; i < numCols; i++) {
      arr.push(students_data.getString(i + 1));
    }
    sheet.appendRow(arr);
  }
  students_data.close();
}

