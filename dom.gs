/************************************************
* DOM関連の処理 
*************************************************/
/************************************************
* HTMLから呼び出すためのdoGet関数 
*************************************************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}
/************************************************
* ステータスメッセージの設定 
*************************************************/
function _setPutMess(message) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(STATUS_MESSAGE, message);
  return message;
}
/************************************************
* ステータスメッセージの取得 
*************************************************/
function _getPutMess() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const statusMessage = scriptProperties.getProperty(STATUS_MESSAGE);
  return statusMessage;
}
/************************************************
* フォームの値を返す 
*************************************************/
function _setForm(formData) {
  return formData;
}