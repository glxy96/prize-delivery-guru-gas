/**
 * 特典配布処理の開始関数
 * メニューから呼び出す最初のエントリーポイント
 */
function startPrizeDistribution() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 配布リストシートを取得
    const distSheet = spreadsheet.getSheetByName(SHEETS.DISTRIBUTION);
    if (!distSheet) {
      throw new Error('配布リストシートが見つかりません。ガチャ結果の解析を先に実行してください。');
    }
    
    // データ範囲を取得（ヘッダー行を除く）
    const dataRange = distSheet.getDataRange();
    const data = dataRange.getValues();
    if (data.length <= 1) {
      throw new Error('配布リストにデータがありません。');
    }
    
    // 実行日時フォルダを作成
    const timestamp = new Date();
    const formattedDate = Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    const rootFolderId = PropertiesService.getScriptProperties().getProperty(PROPERTIES.ROOT_FOLDER_ID);
    const rootFolder = DriveApp.getFolderById(rootFolderId);
    const outputFolder = rootFolder.createFolder(`${FOLDERS.OUTPUT}_${formattedDate}`);
    
    // 特典ファイルの準備確認
    const confirmMessage = `特典ファイルの配置を確認してください：\n\n` +
      `1. 特典ファイルは「prize-guru/prizes」フォルダに配置されていますか？\n` +
      `2. ファイル名は景品名と一致していますか？\n` +
      `   例：「アイコンリング1.jpg」「ヘッダー5.png」など\n\n` +
      `処理を ${data.length - 1} 人のリスナーに対して実行します。時間がかかる場合は自動的に分割処理されます。\n` +
      `処理を開始しますか？`;
    
    if (ui.alert('特典配布処理開始', confirmMessage, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
      return;
    }
    
    // **最初にユーザーに通知**
    ui.alert('処理を開始します', 
      '特典配布処理を開始します。処理は自動的に継続されます。\n' +
      '処理状況は「デバッグ情報」シートで確認できます。\n\n' +
      '処理中はスプレッドシートを開いたままにしておいてください。', 
      ui.ButtonSet.OK);
    
    // 処理状態を初期化
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('processingState', JSON.stringify({
      outputFolderId: outputFolder.getId(),
      startIndex: 1,
      currentIndex: 1,
      totalCount: data.length - 1,
      successCount: 0,
      errorCount: 0,
      startTime: timestamp.getTime(),
      lastProcessedTime: timestamp.getTime(),
      notificationShown: true  // 通知表示済みのフラグを追加
    }));
    
    // デバッグシートの初期化
    initDebugSheet(spreadsheet, timestamp);
    
    // 最初のバッチ処理を開始
    continuePrizeDistribution();
    
    // 通知は既に表示したので二重表示しない
    
  } catch (error) {
    ui.alert('エラー', '処理開始中にエラーが発生しました: ' + error.message, ui.ButtonSet.OK);
    console.error(error);
  }
}

/**
 * 特典配布処理を継続する関数
 * 時間制限に対応するため、処理を分割して実行
 */
function continuePrizeDistribution() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  try {
    // 処理状態を取得
    const stateJson = scriptProperties.getProperty('processingState');
    if (!stateJson) {
      console.log('処理状態が見つかりません。処理は完了しているか、初期化されていません。');
      return;
    }
    
    const state = JSON.parse(stateJson);
    const outputFolder = DriveApp.getFolderById(state.outputFolderId);
    
    // 配布リストシートを取得
    const distSheet = spreadsheet.getSheetByName(SHEETS.DISTRIBUTION);
    const data = distSheet.getDataRange().getValues();
    
    // 特典ファイルマップを取得または作成
    let prizeFilesMap = {};
    let normalizedPrizeFilesMap = {};
    
    // キャッシュされたマップがあれば使用
    const cachedMapJson = scriptProperties.getProperty('prizeFilesMap');
    if (cachedMapJson) {
      const cachedMap = JSON.parse(cachedMapJson);
      prizeFilesMap = cachedMap.original;
      normalizedPrizeFilesMap = cachedMap.normalized;
    } else {
      // 特典ファイルの一覧を作成
      const prizesFolderId = scriptProperties.getProperty(PROPERTIES.PRIZES_FOLDER_ID);
      const prizesFolder = DriveApp.getFolderById(prizesFolderId);
      const prizeFiles = prizesFolder.getFiles();
      
      const fileMap = {};
      const normalizedMap = {};
      
      // ファイルマップを構築
      while (prizeFiles.hasNext()) {
        const file = prizeFiles.next();
        const fileName = file.getName();
        const baseName = fileName.substring(0, fileName.lastIndexOf('.')) || fileName;
        const normalizedBaseName = baseName.normalize('NFC');
        
        fileMap[baseName] = {
          id: file.getId(),
          name: fileName
        };
        
        normalizedMap[normalizedBaseName] = {
          id: file.getId(),
          name: fileName
        };
      }
      
      prizeFilesMap = fileMap;
      normalizedPrizeFilesMap = normalizedMap;
      
      // キャッシュに保存
      scriptProperties.setProperty('prizeFilesMap', JSON.stringify({
        original: fileMap,
        normalized: normalizedMap
      }));
    }
    
    // デバッグシートを取得
    const debugSheet = spreadsheet.getSheetByName('デバッグ情報');
    
    // 処理の開始時間を記録
    const startTime = new Date().getTime();
    const maxExecutionTime = 5.5 * 60 * 1000; // 5.5分（余裕を持たせる）
    const batchSize = 10; // 一度のバッチで処理する最大リスナー数
    
    // 今回のバッチで処理する最大インデックス
    const endIndex = Math.min(state.currentIndex + batchSize, data.length);
    
    // このバッチでの成功・エラーカウント
    let batchSuccessCount = 0;
    let batchErrorCount = 0;
    
    // リスナーごとに処理
    for (let i = state.currentIndex; i < endIndex; i++) {
      // 経過時間をチェックし、制限に近づいたら処理を中断
      const currentTime = new Date().getTime();
      if (currentTime - startTime > maxExecutionTime) {
        console.log(`実行時間制限に近づいたため、インデックス ${i} で処理を中断します`);
        
        // 処理状態を更新
        state.currentIndex = i;
        state.successCount += batchSuccessCount;
        state.errorCount += batchErrorCount;
        state.lastProcessedTime = new Date().getTime();
        scriptProperties.setProperty('processingState', JSON.stringify(state));
        
        // 進捗状況をデバッグシートに追加
        updateProgressInDebugSheet(debugSheet, state);
        
        // 続きを処理するためのトリガーを設定
        deleteExistingTriggers('continuePrizeDistribution');
        ScriptApp.newTrigger('continuePrizeDistribution')
          .timeBased()
          .after(10000) // 10秒後に再開
          .create();
        
        return;
      }
      
      const userName = data[i][0];
      const prizeContent = data[i][2];
      const statusCell = distSheet.getRange(i + 1, 5);
      const urlCell = distSheet.getRange(i + 1, 4);
      const dateCell = distSheet.getRange(i + 1, 6);
      
      // 既に処理済みまたは配布済みの場合はスキップ
      const currentStatus = statusCell.getValue();
      if (currentStatus === '配布済み' || currentStatus === '準備完了') {
        continue;
      }
      
      try {
        // ユーザーごとのフォルダを作成
        const userFolder = outputFolder.createFolder(userName);
        
        // 特典内容から特典名を抽出
        let prizeNames = [];
        if (prizeContent.includes('※特典が多いため概要のみ表示')) {
          // 詳細シートから特典名を取得
          const detailSheet = spreadsheet.getSheetByName(`詳細_${userName}`) || 
                              spreadsheet.getSheetByName(`詳細_${userName.substring(0, 27)}...`);
          
          if (detailSheet) {
            const detailData = detailSheet.getDataRange().getValues();
            // ヘッダー行をスキップして特典名を取得
            for (let j = 3; j < detailData.length; j++) {
              if (detailData[j][2]) {
                prizeNames.push(detailData[j][2]);
              }
            }
          }
        } else {
          // 通常の特典内容から特典名を抽出
          const items = prizeContent.split(', ');
          items.forEach(item => {
            // 「SR アイコンリング1 ×3」または「SR アイコンリング1」形式から特典名部分を抽出
            const match = item.match(/^[A-Z]+ (.+?)(?:\s×\d+)?$/);
            if (match && match[1]) {
              prizeNames.push(match[1]);
            }
          });
        }
        
        if (prizeNames.length === 0) {
          throw new Error('特典名の抽出に失敗しました');
        }
        
        // 特典ファイルをコピー
        let filesCopied = 0;
        let missingFiles = [];
        
        // 重複を排除
        const uniquePrizeNames = [...new Set(prizeNames)];
        
        for (const prizeName of uniquePrizeNames) {
          // Unicode正規化を適用（NFC: 合成）
          const normalizedPrizeName = prizeName.normalize('NFC');
          
          // 正規化されたファイル名マップで検索
          let found = false;
          
          // 1. 正規化された名前で直接検索
          if (normalizedPrizeFilesMap[normalizedPrizeName]) {
            const fileInfo = normalizedPrizeFilesMap[normalizedPrizeName];
            const file = DriveApp.getFileById(fileInfo.id);
            file.makeCopy(fileInfo.name, userFolder);
            filesCopied++;
            found = true;
            debugSheet.appendRow([userName, prizeName, normalizedPrizeName, '一致（正規化後）', fileInfo.name]);
          } 
          // 2. 元の名前で検索（念のため）
          else if (prizeFilesMap[prizeName]) {
            const fileInfo = prizeFilesMap[prizeName];
            const file = DriveApp.getFileById(fileInfo.id);
            file.makeCopy(fileInfo.name, userFolder);
            filesCopied++;
            found = true;
            debugSheet.appendRow([userName, prizeName, normalizedPrizeName, '一致（元の名前）', fileInfo.name]);
          } 
          // 3. 不一致の場合
          else {
            missingFiles.push(prizeName);
            debugSheet.appendRow([userName, prizeName, normalizedPrizeName, '不一致', '対応するファイルが見つかりません']);
          }
        }
        
        // 共有設定と共有URLの取得
        userFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const shareUrl = userFolder.getUrl();
        
        // 配布リストを更新
        urlCell.setValue(shareUrl);
        dateCell.setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
        
        if (filesCopied === uniquePrizeNames.length) {
          statusCell.setValue('準備完了');
          batchSuccessCount++;
        } else {
          const status = `一部ファイル不足 (${filesCopied}/${uniquePrizeNames.length})`;
          statusCell.setValue(status);
          debugSheet.appendRow([userName, '不足ファイル', '', '', missingFiles.join(', ')]);
          batchErrorCount++;
        }
      } catch (error) {
        console.error(`Error processing user ${userName}: ${error.message}`);
        statusCell.setValue('エラー: ' + error.message);
        debugSheet.appendRow([userName, 'エラー', '', '', error.message]);
        batchErrorCount++;
      }
    }
    
    // 処理状態を更新
    state.currentIndex = endIndex;
    state.successCount += batchSuccessCount;
    state.errorCount += batchErrorCount;
    state.lastProcessedTime = new Date().getTime();
    
// 全ての処理が完了したかチェック
if (endIndex >= data.length) {
  // 処理完了
  scriptProperties.deleteProperty('processingState');
  scriptProperties.deleteProperty('prizeFilesMap');
  
  // 完了情報をデバッグシートに追加
  debugSheet.appendRow(['']);
  debugSheet.appendRow(['処理完了時刻', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')]);
  debugSheet.appendRow(['合計処理リスナー数', state.totalCount]);
  debugSheet.appendRow(['成功数', state.successCount]);
  debugSheet.appendRow(['エラー数', state.errorCount]);
  debugSheet.appendRow(['出力フォルダURL', `https://drive.google.com/drive/folders/${state.outputFolderId}`]);
  
  // UI経由で実行された場合は完了メッセージを表示
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '処理完了', 
      `特典ファイルのフォルダ化が完了しました。\n\n` +
      `成功: ${state.successCount}件\n` +
      `不足/エラー: ${state.errorCount}件\n\n` +
      `出力フォルダ: ${outputFolder.getName()}\n` +
      `フォルダURL: https://drive.google.com/drive/folders/${state.outputFolderId}\n\n` +
      `詳細なデバッグ情報は「デバッグ情報」シートを確認してください。`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    // UIがない場合は無視（時間トリガーからの実行時など）
    console.log('処理が完了しました');
  }
} else {
      // まだ処理すべきリスナーが残っている
      scriptProperties.setProperty('processingState', JSON.stringify(state));
      
      // 進捗状況をデバッグシートに追加
      updateProgressInDebugSheet(debugSheet, state);
      
      // 続きを処理するためのトリガーを設定
      deleteExistingTriggers('continuePrizeDistribution');
      ScriptApp.newTrigger('continuePrizeDistribution')
        .timeBased()
        .after(1000) // 1秒後に再開
        .create();
    }
    
  } catch (error) {
    console.error('処理継続中にエラーが発生しました: ' + error.message);
    console.error(error);
    
    // エラーが発生しても次回トリガーを設定して継続を試みる
    deleteExistingTriggers('continuePrizeDistribution');
    ScriptApp.newTrigger('continuePrizeDistribution')
      .timeBased()
      .after(30000) // 30秒後に再試行
      .create();
  }
}

/**
 * デバッグ情報シートを初期化する
 */
function initDebugSheet(spreadsheet, timestamp) {
  let debugSheet = spreadsheet.getSheetByName('デバッグ情報');
  if (debugSheet) {
    spreadsheet.deleteSheet(debugSheet);
  }
  debugSheet = spreadsheet.insertSheet('デバッグ情報');
  
  debugSheet.appendRow(['処理開始日時', Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')]);
  debugSheet.appendRow(['']);
  debugSheet.appendRow(['進捗状況', '']);
  debugSheet.appendRow(['現在の処理ステータス', '初期化中']);
  debugSheet.appendRow(['処理済みリスナー数', '0']);
  debugSheet.appendRow(['残りリスナー数', '計算中']);
  debugSheet.appendRow(['成功数', '0']);
  debugSheet.appendRow(['エラー数', '0']);
  debugSheet.appendRow(['最終更新時刻', Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')]);
  debugSheet.appendRow(['']);
  debugSheet.appendRow(['リスナー名', '特典名', '正規化された特典名', '一致状態', '備考']);
  
  // 列幅調整
  debugSheet.setColumnWidth(1, 150);
  debugSheet.setColumnWidth(2, 250);
  debugSheet.setColumnWidth(3, 250);
  debugSheet.setColumnWidth(4, 150);
  debugSheet.setColumnWidth(5, 350);
  
  return debugSheet;
}

/**
 * デバッグシートの進捗情報を更新する
 */
function updateProgressInDebugSheet(debugSheet, state) {
  const currentTime = new Date();
  const remainingCount = state.totalCount - (state.currentIndex - state.startIndex);
  const progressPercent = Math.round((state.currentIndex - state.startIndex) / state.totalCount * 100);
  
  debugSheet.getRange('B4').setValue('処理中');
  debugSheet.getRange('B5').setValue(`${state.currentIndex - state.startIndex} / ${state.totalCount} (${progressPercent}%)`);
  debugSheet.getRange('B6').setValue(remainingCount);
  debugSheet.getRange('B7').setValue(state.successCount);
  debugSheet.getRange('B8').setValue(state.errorCount);
  debugSheet.getRange('B9').setValue(Utilities.formatDate(currentTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
}

/**
 * 特定の関数名の既存トリガーをすべて削除
 */
function deleteExistingTriggers(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * 処理をキャンセルする関数
 * メニューから呼び出せるようにする
 */
function cancelPrizeDistribution() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  const response = ui.alert(
    '処理のキャンセル',
    '進行中の特典配布処理をキャンセルしますか？\n既に処理済みのリスナーはキャンセルされません。',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    // トリガーを削除
    deleteExistingTriggers('continuePrizeDistribution');
    
    // 処理状態を取得
    const stateJson = scriptProperties.getProperty('processingState');
    if (stateJson) {
      const state = JSON.parse(stateJson);
      
      // デバッグシートに中断情報を追記
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const debugSheet = spreadsheet.getSheetByName('デバッグ情報');
      if (debugSheet) {
        debugSheet.appendRow(['']);
        debugSheet.appendRow(['処理中断時刻', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')]);
        debugSheet.appendRow(['処理済みリスナー数', state.currentIndex - state.startIndex]);
        debugSheet.appendRow(['成功数', state.successCount]);
        debugSheet.appendRow(['エラー数', state.errorCount]);
        debugSheet.getRange('B4').setValue('ユーザーによって中断されました');
      }
      
      // 状態をクリア
      scriptProperties.deleteProperty('processingState');
      scriptProperties.deleteProperty('prizeFilesMap');
    }
    
    ui.alert('処理をキャンセルしました', '特典配布処理がキャンセルされました。', ui.ButtonSet.OK);
  }
}
