/**
 * 使い方シートの初期設定
 */
function setupInstructionSheet(spreadsheet) {
  let instructionSheet = spreadsheet.getSheetByName('使い方');
  if (!instructionSheet) {
    instructionSheet = spreadsheet.insertSheet('使い方');
  }
  
  instructionSheet.clear();
  
  // ヘッダー
  instructionSheet.getRange('A1').setValue('ガチャ特典配布管理ツール - 使い方');
  instructionSheet.getRange('A1').setFontWeight('bold').setFontSize(14);
  
  const instructions = [
    ['', ''],
    ['【初回設定】', ''],
    ['1.', 'メニューの「ガチャ特典配布管理」>「0:初期設定」を実行します'],
    ['2.', 'マイドライブ直下に「prize-guru」フォルダと「prizes」サブフォルダが作成されます（すでに作成済みの場合は何も変更はされません）'],
    ['3.', '「prizes」フォルダに特典ファイルをアップロードします（ガチャで「景品名」として設定した名前と同じ景品を用意してください）'],
    ['', '例：「アイコンリング1」を設定した場合、「アイコンリング1.png」など'],
    ['', ''],
    ['【ガチャ結果のインポート】', ''],
    ['1.', '「ガチャ結果入力」シートを開きます'],
    ['2.', 'スプレッドシートのメニューから「ファイル」>「インポート」を選択します'],
    ['3.', '「アップロード」タブで「なまずガチャ履歴吐き出し」のテキストファイルを選択します'],
    ['4.', 'インポート設定で以下を選択します：'],
    ['', '・「既存のシートの内容を置き換える」'],
    ['', '・区切り文字：「タブ」'],
    ['5.', '「インポート」ボタンをクリックします'],
    ['', ''],
    ['【ガチャ結果の解析とフォルダ作成】', ''],
    ['1.', 'メニューの「なまずガチャ特典配布管理」>「1:ガチャ結果を解析」をクリックします'],
    ['2.', '「配布リスト」シートが作成され、リスナーごとの特典が表示されます'],
    ['3.', 'メニューの「なまずガチャ特典配布管理」>「2:特典ファイルをフォルダ化」をクリックします'],
    ['4.', '各リスナーごとの特典フォルダが作成され、共有URLが配布リストに入力されます'],
    ['', ''],
    ['【Discord送信（未実装）】', ''],
    ['1.', '「設定」シートにDiscord Webhook URLを入力します'],
    ['2.', 'リスナー名と対応するDiscordチャンネルIDを設定します'],
    ['3.', 'メニューの「なまずガチャ特典配布管理」>「Discord送信」をクリックします'],
    ['', ''],
    ['【重要なポイント】', ''],
    ['・', '特典ファイルの名前は、ガチャの「景品名」と一致させてください（例：「アイコンリング1.jpg」）'],
    ['・', '特典ファイルは「prize-guru/prizes」フォルダに置いてください'],
    ['・', 'フォルダ化後、リスナーには個別フォルダへのリンクが共有されます'],
    ['・', 'リスナーは自分が当選した特典のみ閲覧・ダウンロードできます']
  ];
  
  // 説明文を設定
  instructionSheet.getRange(2, 1, instructions.length, 2).setValues(instructions);
  
  // 列幅調整
  instructionSheet.setColumnWidth(1, 30);
  instructionSheet.setColumnWidth(2, 500);
  
  // シートを一番前に移動
  spreadsheet.setActiveSheet(instructionSheet);
  spreadsheet.moveActiveSheet(1);
}/**
 * Prize Delivery Guru
 * Googleスプレッドシート & Google Apps Script を使用して
 * ガチャ特典の配布を効率化するツール
 */

// グローバル変数
const FOLDERS = {
  ROOT: 'prize-guru',
  PRIZES: 'prizes',
  OUTPUT: 'output'
};

const SHEETS = {
  INPUT: 'ガチャ結果入力',
  SETTINGS: '設定',
  DISTRIBUTION: '配布リスト'
};

// SpreadsheetとDriveのIDを保存するためのプロパティ
const PROPERTIES = {
  SPREADSHEET_ID: 'spreadsheetId',
  ROOT_FOLDER_ID: 'rootFolderId',
  PRIZES_FOLDER_ID: 'prizesFolderId'
};

/**
 * メニューをスプレッドシートに追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('なまずガチャ特典配布管理')
    .addItem('ガチャ結果を解析', 'parseAndCreateDistributionList')
    .addItem('特典ファイルをフォルダ化', 'createPrizeFolderForAllListeners')
    .addItem('Discord送信', 'sendDiscordMessages')
    .addItem('初期設定', 'setupTool')
    .addToUi();
}

/**
 * ツールの初期設定を行う
 * - 必要なフォルダの作成
 * - スプレッドシートの作成
 */
function setupTool() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '初期設定を開始します',
    'Google Driveに「prize-guru」フォルダと必要なシートを作成します。続けますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // スプレッドシートID取得
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = spreadsheet.getId();
    PropertiesService.getScriptProperties().setProperty(PROPERTIES.SPREADSHEET_ID, spreadsheetId);
    
    // 必要なフォルダ作成
    createFoldersIfNeeded();
    
    // 必要なシート作成
    createSheetsIfNeeded(spreadsheet);
    
    // 使い方シートを作成
    setupInstructionSheet(spreadsheet);
    
    ui.alert('初期設定完了', '必要なフォルダとシートの作成が完了しました。「使い方」シートの手順に従って進めてください。', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('エラー', '初期設定中にエラーが発生しました: ' + error.message, ui.ButtonSet.OK);
    console.error(error);
  }
}

/**
 * 必要なフォルダを作成する
 */
function createFoldersIfNeeded() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // ルートフォルダの作成または取得
  let rootFolder;
  const rootFolderId = scriptProperties.getProperty(PROPERTIES.ROOT_FOLDER_ID);
  
  if (!rootFolderId) {
    // ルートフォルダが存在しない場合は作成
    rootFolder = DriveApp.createFolder(FOLDERS.ROOT);
    scriptProperties.setProperty(PROPERTIES.ROOT_FOLDER_ID, rootFolder.getId());
  } else {
    try {
      rootFolder = DriveApp.getFolderById(rootFolderId);
    } catch (e) {
      // IDが無効になっていた場合は新規作成
      rootFolder = DriveApp.createFolder(FOLDERS.ROOT);
      scriptProperties.setProperty(PROPERTIES.ROOT_FOLDER_ID, rootFolder.getId());
    }
  }
  
  // prizesフォルダの作成または取得
  let prizesFolder;
  const prizesFolderId = scriptProperties.getProperty(PROPERTIES.PRIZES_FOLDER_ID);
  
  if (!prizesFolderId) {
    // prizesフォルダが存在しない場合は作成
    prizesFolder = rootFolder.createFolder(FOLDERS.PRIZES);
    scriptProperties.setProperty(PROPERTIES.PRIZES_FOLDER_ID, prizesFolder.getId());
  } else {
    try {
      prizesFolder = DriveApp.getFolderById(prizesFolderId);
    } catch (e) {
      // IDが無効になっていた場合は新規作成
      prizesFolder = rootFolder.createFolder(FOLDERS.PRIZES);
      scriptProperties.setProperty(PROPERTIES.PRIZES_FOLDER_ID, prizesFolder.getId());
    }
  }
}

/**
 * 必要なシートを作成する
 */
function createSheetsIfNeeded(spreadsheet) {
  // 入力シートの作成
  let inputSheet = spreadsheet.getSheetByName(SHEETS.INPUT);
  if (!inputSheet) {
    inputSheet = spreadsheet.insertSheet(SHEETS.INPUT);
    setupInputSheet(inputSheet);
  }
  
  // 設定シートの作成
  let settingsSheet = spreadsheet.getSheetByName(SHEETS.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = spreadsheet.insertSheet(SHEETS.SETTINGS);
    setupSettingsSheet(settingsSheet);
  }
  
  // 配布リストシートは解析時に自動作成するため、ここでは作成しない
}

/**
 * 入力シートの初期設定
 */
function setupInputSheet(sheet) {
  sheet.clear();
  
  // タイトルと説明
  sheet.getRange('A1').setValue('ガチャ結果入力シート');
  sheet.getRange('A1').setFontWeight('bold').setFontSize(14);
  
  sheet.getRange('A3').setValue('このシートはガチャ結果テキストをインポートするためのシートです。');
  sheet.getRange('A4').setValue('詳しい手順は「使い方」シートを参照してください。');
  
  // 指示文
  sheet.getRange('A6').setValue('【操作手順】');
  sheet.getRange('A6').setFontWeight('bold');
  
  sheet.getRange('A7').setValue('1. スプレッドシートのメニューから「ファイル」>「インポート」を選択');
  sheet.getRange('A8').setValue('2. 「なまずガチャ履歴吐き出し」のテキストファイルをアップロード');
  sheet.getRange('A9').setValue('3. インポート設定で「既存のシートの内容を置き換える」と「タブ区切り」を選択');
  sheet.getRange('A10').setValue('4. インポート後、メニューの「なまずガチャ特典配布管理」>「ガチャ結果を解析」を実行');
  
  // 空白行
  sheet.getRange('A12').setValue('インポート後はこのテキストは上書きされますが問題ありません。');
}

/**
 * 設定シートの初期設定
 */
function setupSettingsSheet(sheet) {
  sheet.clear();
  
  const headers = [
    ['Discord設定', ''],
    ['Webhook URL', ''],
    ['メッセージテンプレート', '{username}さん、ガチャ特典の配布URLです: {url}'],
    ['', ''],
    ['チャンネルID設定', ''],
    ['リスナー名', 'チャンネルID'],
  ];
  
  sheet.getRange(1, 1, headers.length, 2).setValues(headers);
  sheet.getRange('A1:A5').setFontWeight('bold');
  sheet.getRange('A6:B6').setFontWeight('bold').setBackground('#f3f3f3');
  
  // 列幅調整
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 300);
}

/**
 * ガチャ結果を解析し、配布リストを作成する
 */
function parseAndCreateDistributionList() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 入力シートの確認
    const inputSheet = spreadsheet.getSheetByName(SHEETS.INPUT);
    if (!inputSheet) {
      throw new Error('入力シートが見つかりません。初期設定を実行してください。');
    }
    
    // 入力シートにデータがあるか確認
    const dataRange = inputSheet.getDataRange();
    if (dataRange.getNumRows() <= 1) {
      throw new Error('ガチャ結果データが入力されていません。指示に従ってテキストファイルをインポートしてください。');
    }
    
    // ガチャ結果を解析（モデル変更：直接スプレッドシートのデータを使用）
    const parseResult = parseGachaResults();
    
    // 集計結果が空でないか確認
    if (Object.keys(parseResult.userPrizes).length === 0) {
      throw new Error('有効なガチャ結果が見つかりませんでした。データを確認してください。');
    }
    
    // 配布リストシートの作成/更新
    createDistributionSheet(spreadsheet, parseResult);
    
    ui.alert('解析完了', '配布リストの作成が完了しました。', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('エラー', '解析中にエラーが発生しました: ' + error.message, ui.ButtonSet.OK);
    console.error(error);
  }
}

/**
 * 景品一覧を解析する
 * @param {string} prizeSection 景品一覧テキスト
 * @return {Object[]} 景品一覧
 */
function parsePrizeList(prizeSection) {
  const lines = prizeSection.split('\n').filter(line => line.trim() !== '');
  const prizes = [];
  
  // 最初の行はヘッダーなのでスキップ
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    
    const parts = line.split('\t');
    if (parts.length >= 4) {
      prizes.push({
        number: parseInt(parts[0], 10),
        rarity: parts[1],
        rate: parts[2],
        name: parts[3]
      });
    }
  }
  
  return prizes;
}

/**
 * ガチャ履歴を解析する
 * @param {string} historySection ガチャ履歴テキスト
 * @return {Object[]} ガチャ履歴
 */
function parseGachaHistory(historySection) {
  const lines = historySection.split('\n').filter(line => line.trim() !== '');
  const history = [];
  
  // 最初の行はヘッダーなのでスキップ
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    
    const parts = line.split('\t');
    if (parts.length >= 6) {
      history.push({
        gachaNo: parseInt(parts[0], 10),
        userName: parts[1],
        prizeNo: parseInt(parts[2], 10),
        rarity: parts[3],
        prizeName: parts[4],
        count: parseInt(parts[5], 10)
      });
    }
  }
  
  return history;
}

/**
 * リスナーごとの特典を集計する
 * @param {Object[]} history ガチャ履歴
 * @param {Object[]} prizes 景品一覧
 * @return {Object} リスナーごとの特典
 */
function aggregateUserPrizes(history, prizes) {
  const userPrizes = {};

    // ユーザー名でグループ化
    if (!userPrizes[entry.userName]) {
      userPrizes[entry.userName] = [];
    }
    
    // 対応する景品を見つける
    const prize = prizes.find(p => p.number === entry.prizeNo);
    if (prize) {
      // 個数分追加
      for (let i = 0; i < entry.count; i++) {
        userPrizes[entry.userName].push({
          number: prize.number,
          rarity: prize.rarity,
          name: prize.name
        });
      }
    }
  return userPrizes;
}

/**
 * 特典の詳細リストを別シートに出力
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet スプレッドシート
 * @param {string} userName リスナー名
 * @param {Object[]} prizes 特典リスト
 */
function createDetailedPrizeSheet(spreadsheet, userName, prizes) {
  // シート名（リスナー名が長い場合は短縮）
  let sheetName = `詳細_${userName}`;
  if (sheetName.length > 30) {
    sheetName = sheetName.substring(0, 27) + '...';
  }
  
  // 既存のシートを検索
  let detailSheet = spreadsheet.getSheetByName(sheetName);
  if (detailSheet) {
    // 既存シートがある場合は削除して再作成
    spreadsheet.deleteSheet(detailSheet);
  }
  detailSheet = spreadsheet.insertSheet(sheetName);
  
  // ヘッダー設定
  detailSheet.getRange('A1').setValue(`${userName}の特典詳細リスト`);
  detailSheet.getRange('A1').setFontWeight('bold').setFontSize(14);
  
  detailSheet.getRange('A3:C3').setValues([['No.', 'レアリティ', '特典名']]);
  detailSheet.getRange('A3:C3').setFontWeight('bold').setBackground('#f3f3f3');
  
  // データ行を設定
  const data = prizes.map((prize, index) => [
    index + 1,
    prize.rarity,
    prize.name
  ]);
  
  if (data.length > 0) {
    detailSheet.getRange(4, 1, data.length, 3).setValues(data);
  }
  
  // 列幅調整
  detailSheet.setColumnWidth(1, 60);   // No.
  detailSheet.setColumnWidth(2, 80);   // レアリティ
  detailSheet.setColumnWidth(3, 300);  // 特典名
  
  // データの範囲に罫線を設定
  detailSheet.getRange(3, 1, data.length + 1, 3).setBorder(true, true, true, true, true, true);
}

/**
 * 全リスナーの特典ファイルをフォルダ化する
 */
function createPrizeFolderForAllListeners() {
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
    
    // 特典ファイルが格納されているフォルダを取得
    const prizesFolderId = PropertiesService.getScriptProperties().getProperty(PROPERTIES.PRIZES_FOLDER_ID);
    if (!prizesFolderId) {
      throw new Error('特典フォルダが見つかりません。初期設定を実行してください。');
    }
    const prizesFolder = DriveApp.getFolderById(prizesFolderId);
    
    // 特典ファイルの検出前に確認ダイアログを表示
    const confirmMessage = `特典ファイルの配置を確認してください：\n\n` +
      `1. 特典ファイルは「prize-guru/prizes」フォルダに配置されていますか？\n` +
      `2. ファイル名は景品名と一致していますか？\n` +
      `   例：「アイコンリング1.jpg」「ヘッダー5.png」など\n\n` +
      `特典ファイルの検出を開始しますか？`;
    
    if (ui.alert('特典ファイルの確認', confirmMessage, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
      return;
    }
    
    // デバッグシートを作成
    let debugSheet = spreadsheet.getSheetByName('デバッグ情報');
    if (debugSheet) {
      spreadsheet.deleteSheet(debugSheet);
    }
    debugSheet = spreadsheet.insertSheet('デバッグ情報');
    debugSheet.appendRow(['処理日時', Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')]);
    debugSheet.appendRow(['']);
    debugSheet.appendRow(['検出されたファイル（拡張子なし）', '元のファイル名', '正規化されたファイル名']);
    
    // 特典ファイルの一覧を取得
    const prizeFiles = prizesFolder.getFiles();
    const prizeFilesMap = {};
    const normalizedPrizeFilesMap = {};
    let fileCount = 0;
    let fileNameExamples = [];
    
    while (prizeFiles.hasNext()) {
      const file = prizeFiles.next();
      const fileName = file.getName();
      fileCount++;
      
      // 例として最初の5つのファイル名を記録
      if (fileNameExamples.length < 5) {
        fileNameExamples.push(fileName);
      }
      
      // ファイル名から拡張子を除いたベース名を取得
      const baseName = fileName.substring(0, fileName.lastIndexOf('.')) || fileName;
      
      // Unicode正規化を適用（NFC: 合成）
      const normalizedBaseName = baseName.normalize('NFC');
      
      // 元のベース名と正規化されたベース名の両方を登録
      prizeFilesMap[baseName] = file;
      normalizedPrizeFilesMap[normalizedBaseName] = file;
      
      // デバッグ情報を記録
      debugSheet.appendRow([baseName, fileName, normalizedBaseName]);
    }
    
    // 特典ファイルの検出結果を表示
    if (fileCount === 0) {
      ui.alert('警告', '特典フォルダにファイルが見つかりませんでした。特典ファイルをアップロードしてください。', ui.ButtonSet.OK);
      return;
    } else {
      const fileMessage = `特典フォルダ内に ${fileCount} 個のファイルを検出しました。\n\n` +
        `検出されたファイル名の例：\n` +
        fileNameExamples.join('\n') + 
        '\n\n詳細はスプレッドシートの「デバッグ情報」シートを確認してください。';
      
      ui.alert('特典ファイル検出', fileMessage, ui.ButtonSet.OK);
    }
    
    // デバッグ情報に空行と次のセクションを追加
    debugSheet.appendRow(['']);
    debugSheet.appendRow(['リスナー名', '特典名', '正規化された特典名', '一致状態', '備考']);
    
    // リスナーごとに処理
    let successCount = 0;
    let errorCount = 0;
    
    for (let i = 1; i < data.length; i++) {
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
            normalizedPrizeFilesMap[normalizedPrizeName].makeCopy(normalizedPrizeFilesMap[normalizedPrizeName].getName(), userFolder);
            filesCopied++;
            found = true;
            debugSheet.appendRow([userName, prizeName, normalizedPrizeName, '一致（正規化後）', normalizedPrizeFilesMap[normalizedPrizeName].getName()]);
          } 
          // 2. 不一致の場合
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
        dateCell.setValue(Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
        
        if (filesCopied === uniquePrizeNames.length) {
          statusCell.setValue('準備完了');
          successCount++;
        } else {
          const status = `一部ファイル不足 (${filesCopied}/${uniquePrizeNames.length})`;
          statusCell.setValue(status);
          debugSheet.appendRow([userName, '不足ファイル', '', '', missingFiles.join(', ')]);
          errorCount++;
        }
      } catch (error) {
        console.error(`Error processing user ${userName}: ${error.message}`);
        statusCell.setValue('エラー: ' + error.message);
        debugSheet.appendRow([userName, 'エラー', '', '', error.message]);
        errorCount++;
      }
    }
    
    // 列幅調整（デバッグシート）
    debugSheet.setColumnWidth(1, 150);
    debugSheet.setColumnWidth(2, 250);
    debugSheet.setColumnWidth(3, 250);
    debugSheet.setColumnWidth(4, 150);
    debugSheet.setColumnWidth(5, 350);
    
    // ドライブのURLを取得
    const driveUrl = "https://drive.google.com/drive/folders/" + outputFolder.getId();
    
    ui.alert(
      '処理完了', 
      `特典ファイルのフォルダ化が完了しました。\n\n` +
      `成功: ${successCount}件\n` +
      `不足/エラー: ${errorCount}件\n\n` +
      `出力フォルダ: ${outputFolder.getName()}\n` +
      `フォルダURL: ${driveUrl}\n\n` +
      `詳細なデバッグ情報は「デバッグ情報」シートを確認してください。`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('エラー', 'フォルダ作成処理中にエラーが発生しました: ' + error.message, ui.ButtonSet.OK);
    console.error(error);
  }
}


/**
 * Discordにメッセージを送信する
 */
function sendDiscordMessages() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 配布リストシートを取得
    const distSheet = spreadsheet.getSheetByName(SHEETS.DISTRIBUTION);
    if (!distSheet) {
      throw new Error('配布リストシートが見つかりません。ガチャ結果の解析を先に実行してください。');
    }
    
    // 設定シートを取得
    const settingsSheet = spreadsheet.getSheetByName(SHEETS.SETTINGS);
    if (!settingsSheet) {
      throw new Error('設定シートが見つかりません。初期設定を実行してください。');
    }
    
    // Webhook URLとメッセージテンプレートを取得
    const webhookUrl = settingsSheet.getRange('B2').getValue();
    if (!webhookUrl) {
      throw new Error('Discord Webhook URLが設定されていません。設定シートで設定してください。');
    }
    
    const messageTemplate = settingsSheet.getRange('B3').getValue() || '{username}さん、ガチャ特典の配布URLです: {url}';
    
    // チャンネルID設定を取得
    const settingsData = settingsSheet.getDataRange().getValues();
    const channelSettings = {};
    for (let i = 6; i < settingsData.length; i++) {
      if (settingsData[i][0] && settingsData[i][1]) {
        channelSettings[settingsData[i][0]] = settingsData[i][1];
      }
    }
    
    // 配布リストデータを取得
    const dataRange = distSheet.getDataRange();
    const data = dataRange.getValues();
    if (data.length <= 1) {
      throw new Error('配布リストにデータがありません。');
    }
    
    // 送信確認
    const sendAll = ui.alert(
      'Discord送信確認',
      '全リスナーに特典URLを送信しますか？',
      ui.ButtonSet.YES_NO
    ) === ui.Button.YES;
    
    if (!sendAll) {
      return;
    }
    
    // リスナーごとに処理
    let sentCount = 0;
    let errorCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const userName = data[i][0];
      const shareUrl = data[i][3];
      const status = data[i][4];
      const statusCell = distSheet.getRange(i + 1, 5);
      
      // 共有URLがない、または既に配布済みの場合はスキップ
      if (!shareUrl || status === '配布済み') {
        continue;
      }
      
      // チャンネルIDを取得
      const channelId = channelSettings[userName];
      if (!channelId) {
        // チャンネルIDが設定されていない場合は警告
        statusCell.setValue('チャンネルID未設定');
        continue;
      }
      
      try {
        // メッセージを作成
        const message = messageTemplate
          .replace('{username}', userName)
          .replace('{url}', shareUrl);
        
        // Webhook URLにチャンネルIDを追加
        const webhookWithChannel = `${webhookUrl}?thread_id=${channelId}`;
        
        // Discordに送信
        const response = UrlFetchApp.fetch(webhookWithChannel, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({
            content: message
          })
        });
        
        if (response.getResponseCode() === 204) {
          statusCell.setValue('配布済み');
          sentCount++;
        } else {
          statusCell.setValue('送信エラー');
          errorCount++;
        }
      } catch (error) {
        console.error(`Error sending message to ${userName}: ${error.message}`);
        statusCell.setValue('送信エラー');
        errorCount++;
      }
      
      // Discord APIレート制限を回避するために少し待機
      Utilities.sleep(1000);
    }
    
    const resultMessage = `送信完了: ${sentCount}件\nエラー: ${errorCount}件`;
    ui.alert('送信結果', resultMessage, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('エラー', 'Discord送信中にエラーが発生しました: ' + error.message, ui.ButtonSet.OK);
    console.error(error);
  }
}
