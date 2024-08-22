const DEBUG = true; //Propertiesのログを出力しない場合は　false
const LOG_PUT = true;
const LOG_SHEET_ID = "1y99Db9BY_rQBCOXiuIVBsAeIZQsLJGKF388pR941amg";

//保存データを出力する　
//デバッグとして使用
function _savePropertiesToFile() {

  if ( DEBUG != true ) return;

  const scriptProperties = PropertiesService.getScriptProperties();
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
//property矯正削除　デバッグ用
function _delProperty()
{
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty(PROGRESS_PROPERTY);
  scriptProperties.deleteProperty(EXECUTION_FLAG_KEY);


}

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

function _getMyDriveId() {
  // マイドライブのルートフォルダを取得
  const rootFolder = DriveApp.getRootFolder();
  
  // ルートフォルダのIDをログに出力
  Logger.log('マイドライブのルートフォルダID: ' + rootFolder.getId());
}
