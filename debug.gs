/************************************************
* DEBUG用関数群
*************************************************/

/************************************************
* DEBUG用定数
*************************************************/
const DEBUG = false; //Propertiesのログを出力しない場合は　false
const LOG_PUT = true; //trueの場合はログシートにログを書き出す
const LOG_SHEET_ID = "1y99Db9BY_rQBCOXiuIVBsAeIZQsLJGKF388pR941amg"; //ログシートのシートID

/************************************************
* 保存データの出力
* DEBUGがtrueの場合マイドライブに保存データを出力する
*************************************************/
function _savePropertiesToFile() {

  if ( DEBUG != true ) return;

  //const scriptProperties = PropertiesService.getScriptProperties();
  const scriptProperties = PropertiesService.getUserProperties();

  const allProperties = scriptProperties.getProperties();
  
  var fileContent = "Properties:\n";

  for (let key in allProperties) {
    if (allProperties.hasOwnProperty(key)) {
      fileContent += key + ": " + allProperties[key] + "\n";
    }
  }  
  const file = DriveApp.createFile("debugProperties.txt", fileContent);
  Logger.log("Properties saved to: " + file.getUrl());
}
/************************************************
* ログシートへの書き出し
* LOG_PUTがtrueの場合ログシートにログを書き出す
* デプロイ後にコンソールメッセージが取得できないため
*************************************************/
function _logSheetPut(message) {

  if(LOG_PUT != true) return;
  
  // スプレッドシートを開く
  const spreadsheet = SpreadsheetApp.openById(LOG_SHEET_ID);
  
  // 対象のシートを取得（ここでは最初のシートを取得していますが、シート名で指定も可能）
  const sheet = spreadsheet.getSheets()[0]; // または、sheet.getSheetByName('シート名');

  // 最終行の次の行を取得
  const lastRow = sheet.getLastRow();
  const nextRow = lastRow + 1;

  // データを指定の行に書き込む（ここでは1行目にdataを追加）
  sheet.getRange(nextRow, 1).setValue(message +  new Date() );
}

