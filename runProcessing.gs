
//メイン処理
function _runProcessing() {

  const scriptProperties = PropertiesService.getScriptProperties();
  const startTime = Date.now(); 
  const rootFolederInfo =  _getRootFolderInfo();
  const userMail = _getUserEmail();

  //保存用データがあるかどうかを判断
  let progress = JSON.parse(scriptProperties.getProperty(PROGRESS_PROPERTY) || JSON.stringify({
    folderQueue: [{ id: rootFolederInfo.folderId, layer: 0 }],
      folderListArray:[],
      colorArray:[],
      startTime: startTime,
      sheetId:null,
      email:userMail,
      folderName:rootFolederInfo.folderName,
      itemCnt:itemCnt,
  }));
  
  if(!progress.sheetId) //シートがなかったら作成する
  {
    progress.sheetId = _createRootSpreadSheet(rootFolederInfo.folderId,rootFolederInfo.folderName);
    _initProgress(progress);
    Logger.log('初期設定が完了しました');
    return;
  }

  //リストの階層化
  _folderList(progress);
  _clearTrigger(); //古いトリガーがあれば削除

  //処理の完了を判定
  if (progress.folderQueue.length === 0) {
    scriptProperties.deleteProperty(PROGRESS_PROPERTY);
    Logger.log('フォルダ階層の取得が完了しました');

    // 全てのフォルダが処理された後にシートに書き出す
    _writeSheetList(progress);
    _sendEndMail(progress);

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
      _savePropertiesToFile(); //デバッグ用に保存データを書き出し
      PropertiesService.getScriptProperties().setProperty(PROGRESS_PROPERTY, JSON.stringify(progress));
      Logger.log('タイムアウトが発生しました。処理を中断し、次回に続きます。');
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