/**
 * ガチャ結果のテキストを解析する
 * @return {Object} 解析結果
 */
function parseGachaResults() {
  console.log("解析開始");
  
  // スプレッドシートから直接データを読み取る方式に変更
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.INPUT);
  const values = sheet.getDataRange().getValues();
  
  // 景品一覧とガチャ履歴の位置を特定
  let prizeListStartRow = -1;
  let historyStartRow = -1;
  
  for (let i = 0; i < values.length; i++) {
    const cell = values[i][0]?.toString().trim() || '';
    if (cell.includes('景品一覧')) {
      prizeListStartRow = i;
    } else if (cell.includes('ガチャ履歴')) {
      historyStartRow = i;
      break;
    }
  }
  
  console.log(`景品一覧開始行: ${prizeListStartRow}, ガチャ履歴開始行: ${historyStartRow}`);
  
  if (prizeListStartRow < 0 || historyStartRow < 0) {
    throw new Error('景品一覧またはガチャ履歴セクションが見つかりません。データ形式を確認してください。');
  }
  
  // 景品一覧の解析
  const prizes = [];
  for (let i = prizeListStartRow + 2; i < historyStartRow; i++) {
    if (!values[i][0]) continue; // 空行はスキップ
    
    const row = values[i];
    // 番号、レアリティ、出現率、景品名の各列があるか確認
    if (row[0] && row[1] && row[3]) {
      prizes.push({
        number: parseInt(row[0], 10) || 0,
        rarity: row[1],
        rate: row[2] || '',
        name: row[3]
      });
    }
  }
  
  console.log(`解析された景品数: ${prizes.length}`);
  
  // 景品をマップ化して高速アクセス
  const prizeMap = {};
  prizes.forEach(prize => {
    prizeMap[prize.number] = prize;
  });
  
  // ガチャ履歴の解析
  // リスナーごとの特典を直接集計
  const userPrizes = {};
  
  // ガチャ履歴の見出し行（3行）をスキップ
  for (let i = historyStartRow + 3; i < values.length; i++) {
    if (!values[i][0]) continue; // 空行はスキップ
    
    const row = values[i];
    // ガチャNo.、名前、景品No.、レアリティ、景品名、当たり数の各列があるか確認
    if (row[0] && row[1] && row[2] && row[3] && row[4] && row[5]) {
      const gachaNo = parseInt(row[0], 10) || 0;
      const userName = row[1].toString().trim();
      const prizeNo = parseInt(row[2], 10) || 0;
      const count = parseInt(row[5], 10) || 0;
      
      // 対応する景品情報を取得
      const prize = prizeMap[prizeNo];
      
      if (prize && userName && count > 0) {
        // ユーザー名でグループ化
        if (!userPrizes[userName]) {
          userPrizes[userName] = [];
        }
        
        // 個数分特典を追加
        for (let j = 0; j < count; j++) {
          userPrizes[userName].push({
            number: prize.number,
            rarity: prize.rarity,
            name: prize.name
          });
        }
      }
    }
  }
  
  console.log(`集計されたリスナー数: ${Object.keys(userPrizes).length}`);
  
  return {
    prizes: prizes,
    userPrizes: userPrizes
  };
}

/**
 * 配布リストシートを作成/更新する
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet スプレッドシート
 * @param {Object} parseResult 解析結果
 */
function createDistributionSheet(spreadsheet, parseResult) {
  // 配布リストシートの取得または作成
  let distSheet = spreadsheet.getSheetByName(SHEETS.DISTRIBUTION);
  if (!distSheet) {
    distSheet = spreadsheet.insertSheet(SHEETS.DISTRIBUTION);
  }
  
  // シートをクリア
  distSheet.clear();
  
  // ヘッダー行を設定
  const headers = ['リスナー名', '特典数', '特典内容', '共有URL', '配布状況', '処理日時'];
  distSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  distSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
  
  // リスナーごとのデータを設定
  const userNames = Object.keys(parseResult.userPrizes);
  const data = userNames.map(userName => {
    const prizes = parseResult.userPrizes[userName];
    
    // 特典内容の文字数制限対策
    // レアリティと名前の組み合わせでグループ化してカウント
    const prizeGroups = {};
    prizes.forEach(p => {
      const key = `${p.rarity} ${p.name}`;
      if (!prizeGroups[key]) {
        prizeGroups[key] = 0;
      }
      prizeGroups[key]++;
    });
    
    // 表示用の文字列を構築
    const prizeContent = [];
    Object.keys(prizeGroups).forEach(key => {
      const count = prizeGroups[key];
      if (count > 1) {
        prizeContent.push(`${key} ×${count}`);
      } else {
        prizeContent.push(key);
      }
    });
    
    // 特典数が多い場合は省略表示
    let prizeText = prizeContent.join(', ');
    if (prizeText.length > 40000) { // 余裕を持って制限
      // レアリティごとの集計のみ表示
      const rarityCount = {};
      prizes.forEach(p => {
        if (!rarityCount[p.rarity]) {
          rarityCount[p.rarity] = 0;
        }
        rarityCount[p.rarity]++;
      });
      
      const raritySummary = Object.keys(rarityCount).map(r => 
        `${r}: ${rarityCount[r]}個`
      ).join(', ');
      
      prizeText = `※特典が多いため概要のみ表示: ${raritySummary}`;
      
      // 詳細リストを別シートに出力
      createDetailedPrizeSheet(spreadsheet, userName, prizes);
    }
    
    return [
      userName,
      prizes.length,
      prizeText,
      '', // 共有URLは後で設定
      '未配布',
      ''  // 処理日時は後で設定
    ];
  });
  
  if (data.length > 0) {
    distSheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
  
  // 列幅調整
  distSheet.setColumnWidth(1, 150);  // リスナー名
  distSheet.setColumnWidth(2, 80);   // 特典数
  distSheet.setColumnWidth(3, 400);  // 特典内容
  distSheet.setColumnWidth(4, 250);  // 共有URL
  distSheet.setColumnWidth(5, 100);  // 配布状況
  distSheet.setColumnWidth(6, 150);  // 処理日時
  
  // データの範囲に罫線を設定
  if (data.length > 0) {
    distSheet.getRange(1, 1, data.length + 1, headers.length).setBorder(true, true, true, true, true, true);
  }
}