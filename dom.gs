//--------------------------------------------------------------------------------------------------------------------------------------------
// DOM関連
//--------------------------------------------------------------------------------------------------------------------------------------------

// HTMLから呼び出すためのdoGet関数
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}
//ステータスメッセージの設定
function _setPutMess(message) {

  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(STATUS_MESSAGE, message);
  return message;

 // statusMessage = message;
 // return statusMessage
}
//ステータスメッセージの取得
function _getPutMess() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var statusMessage = scriptProperties.getProperty(STATUS_MESSAGE);
  return statusMessage;
}

function _setForm(formData) {
  
 //statusMessage = "実行中：おまちください";
 _setPutMess("実行中：おまちを");
 //mainformData = new _setFormData(formData.inputId,formData.mode)
//  statusMessage = "入力されたテキスト: " + mainformData.id + "\n選択されたラジオオプション: " + mainformData.mode;
 //_main(); 
  // メッセージを設定
  return formData;
}