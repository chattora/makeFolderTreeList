# makeFolderTreeList

任意のフォルダをルートとしてGoogleドライブ内のフォルダ階層をリスト化するGAS。

## 使用方法
1.階層化したいルートにGASを移動  
2.GASを開き_mainを実行する  
3.開始メールが送信される  
4.処理が終わると終了メールが送信されシートが作成される  

## 仕様
・初回起動でTRIGGER_FUNCで指定された時間後にトリガーが設定され、開始メールを送信  
・最初のトリガー時に"processProgress"というファイル名の設定ファイルを作成する  
・その後はMAX_EXECUTION_TIMEで設定された時間を処理をし、処理が完了していない場合は完了まで再度トリガーを設定する。  
### データ
データそのものを保持をすると膨大なデータ量になるため、再帰的に処理用のフォルダIDを保持している。
  folderQueue: [{ id: rootFolederInfo.folderId, layer: 0 }],




## デバッグ仕様
・デバッグ用にJSON形式で保存したデータをdebugProperties.txtというファイルでマイドライブに出力している　  
・途中で処理を停止した場合、保存データが残っていると最初から処理が始まらないので、_delPropertyを実行してから再度行ってください。  
トリガー確認　→ https://script.google.com/home/

