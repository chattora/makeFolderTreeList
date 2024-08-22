//初期化
function _initProgress(progress)
{
  const scriptProperties = PropertiesService.getScriptProperties();

  _sendStartMail(progress); //開始メール送信
  _setConditional(progress.sheetId); //色付けルール設定

_logSheetPut("init");
_logSheetPut(mainformData.id);
_logSheetPut(mainformData.mode);

  progress.mode = mainformData.mode;
  progress.folderId = mainformData.id;
  scriptProperties.setProperty(PROGRESS_PROPERTY, JSON.stringify(progress)); //データ保存
  _clearTrigger(); //古いトリガーがあれば削除

  ScriptApp.newTrigger(TRIGGER_FUNC)
    .timeBased()
    .after(RESTART_TIME) // 1分後に再実行
    .create();
}
//メイン処理
function _runProcessing() {

  const scriptProperties = PropertiesService.getScriptProperties();
  const startTime = Date.now(); 
  const userMail = _getUserEmail();
  var progress = scriptProperties.getProperty(PROGRESS_PROPERTY);
  _setPutMess("実行中です。");

  //保存データがあるかないかで初期処理を行う
  if(progress === null)
  {
    const rootFolederInfo = _getRootFolderInfo();
    //rootFolderInfoがnullならエラー処理
    if(!rootFolederInfo)
    {
      _setPutMess("IDが適切ではありません");
      _logSheetPut("IDが適切ではありません");
      return;
    }
    else
    {
       const defaultData = {
        folderQueue: [{ id: rootFolederInfo.folderId, layer: 0 }],
        folderListArray: [],
        colorArray: [],
        startTime: startTime,
        sheetId: null,
        email: userMail,
        folderName: rootFolederInfo.folderName,
        itemCnt: 0,
        mode: 0,
        folderId: null,
       };

      defaultData.sheetId = _createRootSpreadSheet(rootFolederInfo.folderId,rootFolederInfo.folderName);
      _initProgress(defaultData);
      Logger.log('初期設定が完了しました');
      _logSheetPut ('初期設定が完了しました');
      _savePropertiesToFile(); //デバッグ用に保存データを書き出し
      _logSheetPut ('floderID' + rootFolederInfo.folderId);
      _setPutMess("設定が完了します。")


      return

    }
  }
  else
  {
     progress = JSON.parse(progress);
     _logSheetPut("2回目以降の処理");
  }
/*
  //保存用データがない場合は初期値
  progress = JSON.parse(scriptProperties.getProperty(PROGRESS_PROPERTY) || JSON.stringify({
    folderQueue: [{ id: rootFolederInfo.folderId, layer: 0 }],
      folderListArray:[],
      colorArray:[],
      startTime: startTime,
      sheetId:null,
      email:userMail,
      folderName:rootFolederInfo.folderName,
      itemCnt:0,
      mode:0,
      folderId:null,
  }));
  
  if(!progress.sheetId) //シートがなかったら作成する
  {
    progress.sheetId = _createRootSpreadSheet(rootFolederInfo.folderId,rootFolederInfo.folderName);
    _initProgress(progress);
    Logger.log('初期設定が完了しました');
    _logSheetPut ('初期設定が完了しました');
    _savePropertiesToFile(); //デバッグ用に保存データを書き出し
    return;
  }
*/
  //リストの階層化
  _folderList(progress);
  _clearTrigger(); //古いトリガーがあれば削除

  //処理の完了を判定
  if (progress.folderQueue.length === 0) {
      _logSheetPut('ファイナライズ');
      _setPutMess("ファイナライズ");

_savePropertiesToFile();
    // 全てのフォルダが処理された後にシートに書き出す
    _writeSheetList(progress);
    _sendEndMail(progress);
    scriptProperties.deleteProperty(PROGRESS_PROPERTY);
    Logger.log('フォルダ階層の取得が完了しました');

  } else {
    scriptProperties.setProperty(PROGRESS_PROPERTY, JSON.stringify(progress));
     ScriptApp.newTrigger(TRIGGER_FUNC)
             .timeBased()
             .after(RESTART_TIME) // 1分後に再実行
             .create();
  }
}
//階層リストの作成
function _folderList(progress) {
  const startTime = Date.now();

  while (progress.folderQueue.length > 0) {

    // タイムアウトチェック
    if ( ( (Date.now() - startTime) >= MAX_EXECUTION_TIME ) ) {

      // 処理時間が過ぎた場合は保存して終了
      // 階層データが膨大になることもあるので、タイムアウト時にデータを逐次的にシートに書き込みメモリを解放して保存
      if (progress.folderListArray.length > 0) {
        _writeSheetList(progress); 
        progress.folderListArray = [];
        progress.colorArray = [];
      }
      PropertiesService.getScriptProperties().setProperty(PROGRESS_PROPERTY, JSON.stringify(progress));
      _savePropertiesToFile(); //デバッグ用に保存データを書き出し
      Logger.log('タイムアウトが発生しました。処理を中断し、次回に続きます。');
      _logSheetPut('タイムアウトが発生しました。処理を中断し、次回に続きます。');
      return;
    }

    const { id, layer } = progress.folderQueue.shift();
    const folder = DriveApp.getFolderById(id);

    const folders = folder.getFolders();
    const files = folder.getFiles();

    // フォルダの処理
    while (folders.hasNext()) {
      const subFolder = folders.next();
      const folderId = subFolder.getId();
      const permission = _getPermissions(folderId);
      const owners = _getOrganizerEmails(permission);
      const writers = _getWriterEmails(permission);
      const readers = _getReaderEmails(permission);

      progress.folderListArray.push([FOLDER_ICON, subFolder.getName(), "フォルダ", subFolder.getUrl(), owners, writers, readers]);
      progress.colorArray.push(layer);
      progress.itemCnt++;
      console.log("フォルダ→"+subFolder.getName() + "itemCnt = " + progress.itemCnt);
      progress.folderQueue.push({ id: folderId, layer: layer + 1 });
    }
    _logSheetPut("mode=" + progress.mode);
    //mode2であればファイルも探索
    if(progress.mode == "mode2")
    {
      _logSheetPut("mode2なのでファイルも行う");
      // ファイルの処理
      while (files.hasNext()) {
        const file = files.next();
        const fileId = file.getId();
        const permission = _getPermissions(fileId);
        const owners = _getOrganizerEmails(permission);
        const writers = _getWriterEmails(permission);
        const readers = _getReaderEmails(permission);

        progress.folderListArray.push([layer, file.getName(), "ファイル", file.getUrl(), owners, writers, readers]);
        progress.colorArray.push(layer);   
        progress.itemCnt++;
        console.log("ファイル→"+file.getName() + "itemCnt = " + progress.itemCnt);
      }
    }
  }
}