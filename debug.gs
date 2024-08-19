const DEBUG = false; //Propertiesのログを出力しない場合は　false

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
}
