/************************************************
* coding by toshiyuki_maehashi@nnn.ac.jp
*************************************************/

/************************************************
* グローバル 
*************************************************/
var mainformData; //フォームデータ格納用
var sharedDrivesListArray = [];
var driveCount = 0;

/************************************************
* 定数 
*************************************************/
const FOLDER_ICON = "📂";
const START_COW = 1;
const START_ROW = 1;
const RESTART_TIME = 1 * 60 * 1000;
const TRIGGER_FUNC = '_main';
const MAX_EXECUTION_TIME = 10 * 60 * 1000; // Google有料版のタイムアウトが30分なので書き込み用のバッファを持って10分で強制終了
const WRITE_ROW_MAX = 50000; //書き込み上限5万行に設定 
const VERSION = "2.001";
const endMessStatus = {
  NONE:0,
  DEFAULT: 1,
  RESET: 2,
  SHARE_LIST:3,
  MAIN_ERR: -1,
};

// 保存データ　
const PROGRESS_PROPERTY = 'processProgress';  //保存データ
const EXECUTION_FLAG_KEY = 'isRunning'; //実行中かどうかを判別（多重実行を防ぐための処理）
const STATUS_MESSAGE = 'statusMessage'; //状況メッセージ格納

//共有ドライブリスト用
const SHARE_LIST_START_COW = 1;
const SHARE_LIST_START_ROW = 2;
const SHARE_LIST_BASE_URL = 'https://drive.google.com/drive/folders/';


/************************************************
* 定数 
*************************************************/
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
  '#e6b8af',
]
/************************************************
* MAIN 
*************************************************/
function _main(formData)
{
  //実行状況を保存フラグで管理
  const scriptProperties = PropertiesService.getUserProperties();
  const isRunning = scriptProperties.getProperty(EXECUTION_FLAG_KEY);
  const progress = scriptProperties.getProperty(PROGRESS_PROPERTY);

  //最初ならフォームのデータを格納
  if(progress === null)
  {
    if(formData.inputId === "マイドライブ"){
      mainformData = new _setFormData(_getMyDriveId(),formData.mode)
    }
    else{
      mainformData = new _setFormData(formData.inputId,formData.mode)
    }
  }
  
  //IDの内容によって処理を変更
  if(formData.inputId === "リセット"){
    _logSheetPut("リセットが完了");
  //  Utilities.sleep(5000);
    _resetData();
    return endMessStatus.RESET;
  } 
    
  //mainの多重処理禁止
  if (isRunning === 'true') {
    Logger.log('すでに処理が実行中です。');
    _logSheetPut("すでにmainは実行されています");
    return endMessStatus.MAIN_ERR;
  }

  // MAIN実行中はフラグをセット
  scriptProperties.setProperty(EXECUTION_FLAG_KEY, 'true');

  //共有ドライブリストの作成
  if(formData.mode === "mode3"){
    const sheetId = _createShareListRootSpreadSheet(formData.inputId);
    _setSharedDrives();
    _writeShareListSheetList(sheetId);
    _resetData();
    return endMessStatus.SHARE_LIST;
  }

  try {

    _setPutMess("処理を設定しています。このままお待ちください");

    //MAIN処理　書き込み上限を超えると-1を返す
    const result = _runProcessing(); 

    //書き込み上限の処理
    if( result === -1)
    {
      throw new Error(
        "リストへの書き込み上限"
         + WRITE_ROW_MAX 
         + "を超えました。\n"
         +"フォルダのみの階層（ファイル数が多いと予想される場合はこちら）を選択し、再度実行してください"
      );
    }
    //html側と同期をとるため5秒待つ　
   // Utilities.sleep(5000);

  } catch (e) {
    Logger.log('エラーが発生しました: ' + e.message);
    _sendErrorMail(e.message); // エラーメッセージをメールで送信
    _delProperty();
    throw e; // エラーを再スローしてログに記録
  }finally{
    //保存データの削除
    scriptProperties.deleteProperty(EXECUTION_FLAG_KEY);
    scriptProperties.deleteProperty(STATUS_MESSAGE);
    return endMessStatus.DEFAULT;
  }
}