//保存データを出力する　
//デバッグとして使用
function _savePropertiesToFile() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var allProperties = scriptProperties.getProperties();
  
  var fileContent = "Properties:\n";
  for (var key in allProperties) {
    if (allProperties.hasOwnProperty(key)) {
      fileContent += key + ": " + allProperties[key] + "\n";
    }
  }
  
  var file = DriveApp.createFile("neoProperties2.txt", fileContent);
  Logger.log("Properties saved to: " + file.getUrl());
}
//propery矯正削除　デバッグ用
function _delProperty()
{
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty(PROGRESS_PROPERTY);
}