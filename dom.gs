//--------------------------------------------------------------------------------------------------------------------------------------------
// DOM関連
//--------------------------------------------------------------------------------------------------------------------------------------------

// HTMLから呼び出すためのdoGet関数
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}
//ステータスメッセージの設定
function _setPutMess(message) {
  statusMessage = message;
  return statusMessage
}
//ステータスメッセージの取得
function _getPutMess() {
  return statusMessage;
}

function _setForm(formData) {
  
 statusMessage = "実行中：おまちください";
 mainformData = new _setFormData(formData.inputId,formData.mode)
//  statusMessage = "入力されたテキスト: " + mainformData.id + "\n選択されたラジオオプション: " + mainformData.mode;

 _main(); 
  // メッセージを設定
  return statusMessage;


//_setPutMess ("入力されたテキスト: " + inputTextValue + "\n選択されたラジオオプション: " + radioOptionValue);
  // データを処理し、結果を返す
 // return "入力されたテキスト: " + inputTextValue + "\n選択されたラジオオプション: " + radioOptionValue;
}