/******************************
 *  Global
 ******************************/
 var mainformData;
/******************************
 *  Constant
 ******************************/
const FOLDER_ICON = "📂";
const START_COW = 1;
const START_ROW = 1;
const RESTART_TIME = 1 * 60 * 1000;
const TRIGGER_FUNC = '_main';
const MAX_EXECUTION_TIME = 10 * 60 * 1000; // Google有料版のタイムアウトが30分なので書き込み用のバッファを持って10分で強制終了
const WRITE_ROW_MAX = 1000;
const VERSION = "1.002";

//保存データ
const PROGRESS_PROPERTY = 'processProgress';  //保存データ
const EXECUTION_FLAG_KEY = 'isRunning'; //実行中かどうかを判別（多重実行を防ぐための処理）
const STATUS_MESSAGE = 'statusMessage';

/******************************
 *  Data_table
 ******************************/
//フォルダを色分けする色テーブル
const FOLDER_COLOR_TBL = [
  '#e6b8af',
  '#f4cccc',
  '#fce5cd',
  '#fff2cc',
  '#d9ead3',
  '#d0e0e3',
  '#c9daf8',
  '#d9d2e9',
  '#ead1dc',
  '#e6b8af',
  '#f4cccc',
  '#fce5cd',
  '#fff2cc',
  '#d9ead3',
  '#d0e0e3',
  '#c9daf8',
  '#d9d2e9',
  '#ead1dc',
]
/******************************
 *  MAIN
 ******************************/
function _main(formData)
{
  const scriptProperties = PropertiesService.getScriptProperties();
  const isRunning = scriptProperties.getProperty(EXECUTION_FLAG_KEY);
 
  //フォームのデータを格納
  mainformData = new _setFormData(formData.inputId,formData.mode)

  if (isRunning === 'true') {
    Logger.log('すでに処理が実行中です。');
    _logSheetPut("すでにmainは実行されています");
    return;
  }

  // 処理中フラグをセット
  scriptProperties.setProperty(EXECUTION_FLAG_KEY, 'true');

  try {
    _setPutMess("処理を設定しています。このままお待ちください");
    _runProcessing();
    Utilities.sleep(5000);

  } catch (e) {
    Logger.log('エラーが発生しました: ' + e.message);
    _sendErrorMail(e.message); // エラーメッセージをメールで送信
    throw e; // エラーを再スローしてログに記録
  }finally{
    scriptProperties.deleteProperty(EXECUTION_FLAG_KEY);
    scriptProperties.deleteProperty(STATUS_MESSAGE);
  }

}