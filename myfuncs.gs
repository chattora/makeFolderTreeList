/******************************
 *  myFuncs
 ******************************/
// 実行者のメールアドレスを取得
 function _getUserEmail() {
  try {
    const email = Session.getActiveUser().getEmail();
    
    // メールアドレスをログに表示
    Logger.log('Current user email: ' + email);
    return email

  } catch (e) {
    Logger.log('Error: ' + e.message);
  }
}
//開始メール送信
function _sendStartMail(progress) {
  try {
    // 実行者のメールアドレスを取得
    const email = progress.email;
    
    if (email) {
      // メールの件名と本文を設定
      const subject = progress.folderName + 'の階層リストを作成します';
      const body = '作業の終了まで数分から数時間かかる場合があります。\n'
      + '完了メールが送られるまでお待ちください。';

      // メールを送信
      MailApp.sendEmail(email, subject, body);
      
      Logger.log('Email sent to: ' + email);
    } else {
      Logger.log('No email address found for the current user.');
    }
  } catch (e) {
    Logger.log('Error: ' + e.message);
  }
}
//終了メール送信
function _sendEndMail(progress) {
  try {
    // 実行者のメールアドレスを取得
    const email = progress.email;
    const url = 'https://docs.google.com/spreadsheets/d/' + progress.sheetId;

    if (email) {
      // メールの件名と本文を設定
      const subject = progress.folderName + 'の階層リスト作成が完了しました';
      const body = 'リストが完了しました\n'
      +"url:" + url;

      // メールを送信
      MailApp.sendEmail(email, subject, body);
      
      Logger.log('Email sent to: ' + email);
    } else {
      Logger.log('No email address found for the current user.');
    }
  } catch (e) {
    Logger.log('Error: ' + e.message);
  }
}
//エラー通知用メール
function _sendErrorMail(errorMessage) {

  const recipient = Session.getActiveUser().getEmail(); // ユーザーのメアド取得
  const scriptName = Session.getScriptName(); // スクリプトファイル名を取得
  const subject = scriptName + ' のエラー通知';
  const body = 'スクリプトでエラーが発生しました:\n\n' +
               'エラーメッセージ: ' + errorMessage + '\n' +
               'スクリプト名: ' + scriptName + '\n' +
               '発生日時: ' + new Date();
  
  MailApp.sendEmail(recipient, subject, body);
}
//スクリプトのあるルート情報を取得
function _getRootFolderInfo(){

  const scriptId = ScriptApp.getScriptId();
  const file = DriveApp.getFileById(scriptId);
  const folders = file.getParents();

  while(folders.hasNext()){
    var folder = folders.next();
  }
  return{
    folderName:folder.getName(),
    folderId:folder.getId()
  };
}
//リストを出力するスプレッドシートを作成する
function _createRootSpreadSheet(folderId, sheetName){

  const spreadSheet = SpreadsheetApp.create(sheetName + "の階層リストシート_ver" + VERSION );
  const sheetId = spreadSheet.getId();
  const sheet = spreadSheet.getActiveSheet();

  //ヘッダーの書き込み
  sheet.appendRow(["階層", "フォルダ/ファイル名", "タイプ", "URL","管理者","編集者","閲覧者"]);

  if(folderId)
  {
    const sheet =  DriveApp.getFileById(sheetId);
    const folder = DriveApp.getFolderById(folderId);
    sheet.moveTo(folder);
  }
  return sheetId;
}

//初期化
function _initProgress(progress)
{
  const scriptProperties = PropertiesService.getScriptProperties();

  _sendStartMail(progress); //開始メール送信
  _setConditional(progress.sheetId); //色付けルール設定
  scriptProperties.setProperty(PROGRESS_PROPERTY, JSON.stringify(progress)); //データ保存
  _clearTrigger(); //古いトリガーがあれば削除

  ScriptApp.newTrigger(TRIGGER_FUNC)
    .timeBased()
    .after(RESTART_TIME) // 1分後に再実行
    .create();
}
//セル色付けのための条件を設定
function _setConditional(sheetId) {

  const sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
  const rules = sheet.getConditionalFormatRules();

  sheet.setConditionalFormatRules([]); // 空の配列をセットして全ての条件付き書式ルールを削除

  const ruleFolder = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A1="📂"') // A列の値がフォルダの場合
      .setBackground("#FFEEFF") // 背景色を薄紫に設定
      .setRanges([sheet.getRange("A:G")]) // A列からC列全体を指定
      .build();

  rules.push(ruleFolder);

  for(let i = 0; i < FOLDER_COLOR_TBL.length;i++)
  {
    const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($A1=${i}, NOT(ISBLANK($A1)))` )// A列の値が現在の数値と一致する　かつ　空白ではない場合
        .setBackground(FOLDER_COLOR_TBL[i]) // テーブルから色を設定
        .setRanges([sheet.getRange("A:G")]) // A列からC列全体を指定
        .build();
    
    rules.push(rule);
  }
  sheet.setConditionalFormatRules(rules);
}

//リストをシートに書き出す
 function _writeSheetList(progress)
 {
  const spreadSheet = SpreadsheetApp.openById(progress.sheetId);
  const sheet = spreadSheet.getActiveSheet();

  //_setHierarcheyColor(sheet,progress.colorArray); //処理が重いのでスキップ

  sheet.getRange(sheet.getLastRow() + START_ROW,START_COW,progress.folderListArray.length,progress.folderListArray[0].length).setValues(progress.folderListArray);

 }
//階層を色ごとに分ける
function _setHierarcheyColor(sheet,colorArray)
{
    for(let i=0; i< colorArray.length;i++)
  {
    if( isNaN(colorArray[i]) == false )
    {
      const range = sheet.getRange( (sheet.getLastRow()+ i ) +START_ROW, START_COW,1,3);
      range.setBackground(FOLDER_COLOR_TBL[ colorArray[i] % FOLDER_COLOR_TBL.length ]);
      range.setHorizontalAlignment("right"); // 右寄せに設定
    }
  }
}
//フォルダ・ファイルの権限を取得
function _getPermissions(fileId) {

  var permissionArray  = [];

  try {
    // Drive API を使用してファイルの権限情報を取得する
    let permissions = Drive.Permissions.list(fileId, {
      'supportsAllDrives': true,
      'includeItemsFromAllDrives': true
    });
    
    if (permissions && permissions.items) {
      let permissionList = permissions.items;

      for (let i = 0; i < permissionList.length; i++) {
        permissionArray.push( permissionList[i] );
      }
    
    } else {
      Logger.log('権限情報が見つかりませんでした。');
    }
  } catch (e) {
    Logger.log('エラー: ' + e.message);
  }
  return permissionArray;
}

//トリガーの削除
function _clearTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === TRIGGER_FUNC) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

//管理者のroleを持つ権限の email だけを取り出す
function _getOrganizerEmails(permissionsArray) {
  const emailsArray = permissionsArray
    .filter(function(permission) {
      return permission.role === 'owner'|| permission.role === 'organizer' || permission.role === 'fileOrganizer'; 
    })
    .map(function(permission) {
      return permission.emailAddress;
    });

  const emailsString = emailsArray.length === 0 ? '-': emailsArray.join(', ');
  return emailsString;
}
//編集者のroleを持つ権限の email だけを取り出す
function _getWriterEmails(permissionsArray) {
  const emailsArray = permissionsArray
    .filter(function(permission) {
      return permission.role === 'writer'; 
    })
    .map(function(permission) {
      return permission.emailAddress;
    });
  
  const emailsString = emailsArray.length === 0 ? '-': emailsArray.join(', ');
  
  return emailsString;
}
//閲覧者のroleを持つ権限の email だけを取り出す
function _getReaderEmails(permissionsArray) {
  const emailsArray = permissionsArray
    .filter(function(permission) {
      return permission.role === 'reader';
    })
    .map(function(permission) {
      return permission.emailAddress;
    });
  
  const emailsString = emailsArray.length === 0 ? '-': emailsArray.join(', ');
  
  return emailsString;
}

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

var hogein;
function _setForm(formData) {
  
  const inputTextValue = formData.inputText;
  hogein = formData.radioOption;

 


  statusMessage = "入力されたテキスト: " + inputTextValue + "\n選択されたラジオオプション: " + hogein;

 _main(); 
  // メッセージを設定
  return statusMessage;


//_setPutMess ("入力されたテキスト: " + inputTextValue + "\n選択されたラジオオプション: " + radioOptionValue);
  // データを処理し、結果を返す
 // return "入力されたテキスト: " + inputTextValue + "\n選択されたラジオオプション: " + radioOptionValue;
}