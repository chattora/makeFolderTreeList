/******************************
 *  Global
 ******************************/
 var itemCnt;
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
const PROGRESS_PROPERTY = 'processProgress';  //保存データ
const VERSION = 0
 
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
]
/******************************
 *  MAIN
 ******************************/
function _main()
{
  itemCnt =0;

  try {
    _runProcessing();
  } catch (e) {
    Logger.log('エラーが発生しました: ' + e.message);
    _sendErrorMail(e.message); // エラーメッセージをメールで送信
    throw e; // エラーを再スローしてログに記録
  }
}