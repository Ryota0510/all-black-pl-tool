/**
 * 転記処理.gs
 * LINE売上報告からマスターシートへの自動転記機能
 * 
 * 【重要】このファイルは Google Apps Script プロジェクトに新規作成してください
 * ファイル名: 転記処理.gs
 */

/**
 * LINE売上報告からマスターシートへの自動転記メイン関数（一括処理版）
 * processSheet関数の完了後に自動実行される
 */
function transferToMasterSheetBatch() {
  try {
    console.log('=== 転記処理開始（一括モード） ===');
    
    // 1. スプレッドシートとソースシートを取得
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = spreadsheet.getSheetByName('コピペ用');
    
    if (!sourceSheet) {
      showError('エラー: 「コピペ用」シートが見つかりません。シート名を確認してください。');
      return;
    }
    
    // 2. C列の全データを取得
    var lastRow = sourceSheet.getLastRow();
    if (lastRow < 1) {
      showError('エラー: 「コピペ用」シートにデータが存在しません。');
      return;
    }
    
    var sourceData = sourceSheet.getRange('C1:C' + lastRow).getValues();
    var flatData = [];
    for (var i = 0; i < sourceData.length; i++) {
      flatData.push(sourceData[i][0]);
    }
    console.log('取得データ件数:', flatData.length);
    
    // 3. データを店舗ごとの報告ブロックに構造化
    var reportBlocks = parseReportData(flatData);
    console.log('解析された報告ブロック数:', reportBlocks.length);
    
    if (reportBlocks.length === 0) {
      showError('エラー: 有効な報告データが見つかりません。データ形式を確認してください。');
      return;
    }
    
    // 4. 各報告ブロックを処理（一括モード）
    var processedCount = 0;
    var errorCount = 0;
    
    for (var i = 0; i < reportBlocks.length; i++) {
      var block = reportBlocks[i];
      console.log('処理中: ' + block.store + ' - ' + block.date);
      
      // データが空の場合はスキップ
      if (Object.keys(block.data).length === 0) {
        console.log('⚠ データが空のためスキップ: ' + block.store + ' - ' + block.date);
        continue;
      }
      
      try {
        // マスターシート名を生成して取得
        var masterSheet = getMasterSheet(spreadsheet, block.date);
        if (!masterSheet) {
          console.log('⚠ マスターシートが見つからないため、この店舗はスキップします');
          errorCount++;
          continue;
        }
        
        console.log('✓ マスターシート取得成功:', masterSheet.getName());
        
        // エラーチェック1: 日付重複チェック（一括モードでは自動承認）
        var duplicateCheckResult = checkDateDuplicateBatch(masterSheet, block);
        if (duplicateCheckResult === 'abort') {
          console.log('日付重複チェックで処理中断');
          return;
        }
        
        // エラーチェック2: 前日データとの完全重複チェック
        if (checkPreviousDayDuplicate(masterSheet, block)) {
          console.log('前日重複チェックで処理中断');
          return;
        }
        
        // データ書き込み実行
        writeDataToMasterSheet(masterSheet, block);
        processedCount++;
        console.log('✓ 店舗 ' + block.store + ' の処理完了');
        
      } catch (error) {
        console.error('✗ 店舗 ' + block.store + ' の処理でエラー:', error);
        errorCount++;
        console.log('エラーが発生しましたが、処理を続行します');
      }
    }
    
    // 5. 成功通知
    var resultMessage = '';
    if (processedCount > 0) {
      resultMessage = processedCount + '件の店舗データの転記が完了しました。';
      if (errorCount > 0) {
        resultMessage += '\n' + errorCount + '件でエラーが発生しました。';
      }
      showSuccess(resultMessage);
    } else {
      showInfo('転記対象のデータが見つかりませんでした。\n' +
               'C列のデータ形式を確認してください。');
    }
    
  } catch (error) {
    console.error('転記処理全体でエラー:', error);
    showError('システムエラーが発生しました。\n詳細: ' + error.message);
  }
}

/**
 * 日付重複チェック（一括処理版 - 自動で上書き）
 * @param {Sheet} masterSheet - マスターシート
 * @param {Object} block - 報告ブロック
 * @returns {string} 'proceed'（続行）, 'abort'（中断）
 */
function checkDateDuplicateBatch(masterSheet, block) {
  try {
    console.log('日付重複チェック開始（一括モード）: ' + block.store + ' - ' + block.date);
    
    // 店舗行と日付列を特定
    var storeRowInfo = findStoreRows(masterSheet, block.store);
    var dateCol = findDateColumn(masterSheet, block.date);
    
    if (!storeRowInfo || !dateCol) {
      console.log('店舗または日付が見つからないため、重複チェックをスキップ');
      return 'proceed';
    }
    
    // 主要項目（売上、仕入）の既存データをチェック
    var checkItems = ['売上', '仕入'];
    var hasExistingData = false;
    
    for (var j = 0; j < checkItems.length; j++) {
      var itemName = checkItems[j];
      var mappedItemName = mapItemName(itemName);
      var itemRow = storeRowInfo[mappedItemName];
      if (itemRow) {
        var cellValue = masterSheet.getRange(itemRow, dateCol).getValue();
        if (cellValue && cellValue !== 0) {
          hasExistingData = true;
          break;
        }
      }
    }
    
    if (hasExistingData) {
      console.log('既存データを自動上書き（一括モード）: ' + block.store + ' - ' + block.date);
    }
    
    return 'proceed';
    
  } catch (error) {
    console.error('日付重複チェックエラー:', error);
    return 'abort';
  }
}

/**
 * LINE売上報告からマスターシートへの自動転記メイン関数（確認モード）
 * 1店舗・1日ずつ確認ダイアログを表示
 */
function transferToMasterSheet() {
  try {
    console.log('=== 転記処理開始 ===');
    
    // 1. スプレッドシートとソースシートを取得
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = spreadsheet.getSheetByName('コピペ用');
    
    if (!sourceSheet) {
      showError('エラー: 「コピペ用」シートが見つかりません。シート名を確認してください。');
      return;
    }
    
    // 2. C列の全データを取得
    var lastRow = sourceSheet.getLastRow();
    if (lastRow < 1) {
      showError('エラー: 「コピペ用」シートにデータが存在しません。');
      return;
    }
    
    var sourceData = sourceSheet.getRange('C1:C' + lastRow).getValues();
    var flatData = [];
    for (var i = 0; i < sourceData.length; i++) {
      flatData.push(sourceData[i][0]);
    }
    console.log('取得データ件数:', flatData.length);
    
    // デバッグ: C列の最初の20行を表示
    console.log('=== C列データサンプル（最初の20行）===');
    for (var i = 0; i < Math.min(flatData.length, 20); i++) {
      console.log('行' + (i+1) + ':', String(flatData[i]));
    }
    
    // 3. データを店舗ごとの報告ブロックに構造化
    var reportBlocks = parseReportData(flatData);
    console.log('解析された報告ブロック数:', reportBlocks.length);
    
    if (reportBlocks.length === 0) {
      showError('エラー: 有効な報告データが見つかりません。データ形式を確認してください。');
      return;
    }
    
    // 4. 各報告ブロックを処理
    var processedCount = 0;
    
    for (var i = 0; i < reportBlocks.length; i++) {
      var block = reportBlocks[i];
      console.log('=== 店舗処理確認: ' + (i + 1) + '/' + reportBlocks.length + ' ===');
      console.log('店舗:', block.store);
      console.log('日付:', block.date);
      console.log('データ項目数:', Object.keys(block.data).length);
      
      // データが空の場合はスキップ
      if (Object.keys(block.data).length === 0) {
        console.log('⚠ データが空のためスキップ: ' + block.store + ' - ' + block.date);
        continue;
      }
      
      // 1店舗・1日ずつ確認ダイアログを表示
      var ui = SpreadsheetApp.getUi();
      var normalizedStoreName = normalizeStoreName(block.store);
      var formattedDate = formatDateForDisplay(block.date);
      
      var confirmMessage = '店舗データの転記確認\n\n' +
                          '店舗: ' + block.store + '\n' +
                          '正規化後: ' + normalizedStoreName + '\n' +
                          '日付: ' + formattedDate + '\n' +
                          'データ項目数: ' + Object.keys(block.data).length + '件\n\n' +
                          'この店舗のデータを転記しますか？\n\n' +
                          '【データ詳細】\n';
      
      // データ項目の詳細を追加
      for (var itemName in block.data) {
        var value = block.data[itemName];
        var formattedValue = (typeof value === 'number') ? value.toLocaleString() + '円' : String(value);
        confirmMessage += '・' + itemName + ': ' + formattedValue + '\n';
      }
      
      var response = ui.alert(
        '転記確認 (' + (i + 1) + '/' + reportBlocks.length + ')',
        confirmMessage,
        ui.ButtonSet.YES_NO_CANCEL
      );
      
      if (response === ui.Button.CANCEL) {
        console.log('ユーザーが処理をキャンセルしました');
        showInfo('処理をキャンセルしました。');
        return;
      } else if (response === ui.Button.NO) {
        console.log('この店舗をスキップ: ' + block.store + ' - ' + block.date);
        continue;
      }
      
      console.log('ユーザーが転記を承認: ' + block.store + ' - ' + block.date);
      
      try {
        // マスターシート名を生成して取得
        var masterSheet = getMasterSheet(spreadsheet, block.date);
        if (!masterSheet) {
          console.log('⚠ マスターシートが見つからないため、この店舗はスキップします');
          continue; // エラーは getMasterSheet 内で表示済み
        }
        
        console.log('✓ マスターシート取得成功:', masterSheet.getName());
        
        // エラーチェック1: 日付重複チェック
        var duplicateCheckResult = checkDateDuplicate(masterSheet, block);
        if (duplicateCheckResult === 'abort') {
          console.log('日付重複チェックで処理中断');
          return; // 処理中断
        }
        
        // エラーチェック2: 前日データとの完全重複チェック
        if (checkPreviousDayDuplicate(masterSheet, block)) {
          console.log('前日重複チェックで処理中断');
          return; // 処理中断（エラー表示は関数内で実施）
        }
        
        // データ書き込み実行
        writeDataToMasterSheet(masterSheet, block);
        processedCount++;
        console.log('✓ 店舗 ' + block.store + ' の処理完了');
        
      } catch (error) {
        console.error('✗ 店舗 ' + block.store + ' の処理でエラー:', error);
        console.error('エラースタック:', error.stack);
        
        // エラー時の詳細ダイアログ
        var errorResponse = ui.alert(
          'エラー発生',
          '店舗: ' + block.store + '\n' +
          '日付: ' + block.date + '\n' +
          'エラー: ' + error.message + '\n\n' +
          '処理を続行しますか？\n' +
          '「はい」: 次の店舗に進む\n' +
          '「いいえ」: 処理を中断',
          ui.ButtonSet.YES_NO
        );
        
        if (errorResponse === ui.Button.NO) {
          console.log('ユーザーがエラー後の処理継続を拒否');
          return;
        }
        
        console.log('エラーが発生しましたが、処理を続行します');
      }
    }
    
    // 5. 成功通知
    if (processedCount > 0) {
      showSuccess(processedCount + '件の店舗データの転記が正常に完了しました。');
    } else {
      showInfo('転記対象のデータが見つかりませんでした。\n' +
               'C列のデータ形式を確認してください。\n\n' +
               '詳細はGASエディターのログを確認してください。');
    }
    
  } catch (error) {
    console.error('転記処理全体でエラー:', error);
    showError('システムエラーが発生しました。\n詳細: ' + error.message);
  }
}

/**
 * ソースデータを店舗ごとの報告ブロックに解析
 * 店舗名のスペースを考慮した解析
 * @param {Array} sourceData - C列の全データ
 * @returns {Array} 解析された報告ブロックの配列
 */
function parseReportData(sourceData) {
  var blocks = [];
  var currentBlock = null;
  
  for (var i = 0; i < sourceData.length; i++) {
    var cellValue = String(sourceData[i]).trim();
    
    // 空行はスキップ
    if (!cellValue) continue;
    
    // 日付行を検出（新しいブロックの開始）
    if (cellValue.indexOf('日付') > -1 || cellValue.indexOf('日時') > -1 || cellValue.indexOf('【日付】') > -1) {
      // 前のブロックを保存
      if (currentBlock && currentBlock.date && currentBlock.store) {
        blocks.push(currentBlock);
      }
      
      // 新しいブロックを開始
      currentBlock = {
        store: '',
        date: '',
        data: {}
      };
      
      // 日付を抽出（複数パターンに対応）
      // 括弧内の曜日情報を削除（全角・半角括弧に対応）
      var cleanedValue = cellValue.replace(/[\(（][^）\)]*[\)）]/g, '');
      
      var dateMatch = cleanedValue.match(/(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})/);
      if (dateMatch) {
        currentBlock.date = dateMatch[1].replace(/-/g, '/');
      } else {
        // 「7月29日」「7月29」形式の場合（日がない場合も対応）
        var monthDayMatch = cleanedValue.match(/(\d{1,2})月(\d{1,2})日?/);
        if (monthDayMatch) {
          var currentYear = new Date().getFullYear();
          var month = monthDayMatch[1].padStart(2, '0');
          var day = monthDayMatch[2].padStart(2, '0');
          currentBlock.date = currentYear + '/' + month + '/' + day;
        }
      }
      
      console.log('新しいブロック開始 - 日付:', currentBlock.date, '（元データ:', cellValue, '）');
      continue;
    }
    
    // 現在のブロックが存在しない場合はスキップ
    if (!currentBlock) continue;
    
    // 店舗名を検出（スペースを考慮）
    var normalizedCellValue = cellValue.replace(/\s+/g, '');
    if ((normalizedCellValue.indexOf('店舗') > -1 || normalizedCellValue.indexOf('【店舗名】') > -1) && !currentBlock.store) {
      currentBlock.store = cellValue;
      console.log('店舗名検出: "' + cellValue + '"');
      continue;
    }
    
    // データ項目を解析（様々な形式に対応）
    // パターン1: 【項目名】  数値円
    // パターン2: 項目名  数値円 
    // パターン3: 【項目名】項目詳細  数値円
    
    var itemMatch = null;
    var itemName = '';
    var value = 0;
    
    // 「円」を含む数値を検索（スペースも考慮）
    var yenMatch = cellValue.match(/([0-9,]+)\s*円/);
    if (yenMatch) {
      var valueStr = yenMatch[1].replace(/,/g, '');
      value = parseInt(valueStr, 10);
      
      if (!isNaN(value) || value === 0) {  // 0円も含める
        // 項目名を抽出
        if (cellValue.indexOf('【売上】') > -1) {
          itemName = '売上';
        } else if (cellValue.indexOf('【人件費】') > -1 && cellValue.indexOf('P/A') > -1) {
          itemName = 'P/A';
        } else if (cellValue.indexOf('P/A') > -1 && /[0-9,]+\s*円/.test(cellValue)) {
          itemName = 'P/A';
        } else if (cellValue.indexOf('【人件費】') > -1 && cellValue.indexOf('社員') > -1) {
          itemName = '社員';
        } else if (cellValue.indexOf('社員') > -1 && /[0-9,]+\s*円/.test(cellValue)) {
          itemName = '社員';
        } else if (cellValue.indexOf('【人件費】') > -1) {
          itemName = '人件費';
        } else if (cellValue.indexOf('【仕入費】') > -1 || cellValue.indexOf('仕入') > -1) {
          itemName = '仕入';
        } else {
          // その他の項目名を抽出（【】内または最初の単語）
          var bracketMatch = cellValue.match(/【(.+?)】/);
          if (bracketMatch) {
            itemName = bracketMatch[1];
          } else {
            // 項目名を推測（数値の前の文字列）
            var beforeNumber = cellValue.substring(0, cellValue.search(/[0-9,]+/)).trim();
            if (beforeNumber) {
              itemName = beforeNumber;
            }
          }
        }
        
        if (itemName) {
          console.log('データ項目検出:', itemName, '=', value, '（元データ:', cellValue, '）');
          currentBlock.data[itemName] = value;
        } else {
          console.log('項目名が特定できない行:', cellValue);
        }
      }
    } else {
      // 「円」がない場合の従来の処理
      var dataMatch = cellValue.match(/^(.+?)[:：\s]*([0-9,]+)$/);
      if (dataMatch) {
        itemName = dataMatch[1].trim();
        var valueStr = dataMatch[2].replace(/,/g, '');
        value = parseInt(valueStr, 10);
        
        if (!isNaN(value)) {
          console.log('データ項目検出（従来形式）:', itemName, '=', value);
          currentBlock.data[itemName] = value;
        }
      } else {
        console.log('解析できない行:', cellValue);
      }
    }
  }
  
  // 最後のブロックを保存
  if (currentBlock && currentBlock.date && currentBlock.store) {
    blocks.push(currentBlock);
  }
  
  // デバッグ: 解析されたブロック情報を表示
  console.log('=== 解析結果 ===');
  for (var i = 0; i < blocks.length; i++) {
    var block = blocks[i];
    console.log('ブロック' + (i+1) + ':');
    console.log('  店舗:', block.store);
    console.log('  日付:', block.date);
    console.log('  データ項目数:', Object.keys(block.data).length);
    for (var key in block.data) {
      console.log('    ' + key + ':', block.data[key]);
    }
  }
  
  return blocks;
}

/**
 * 日付からマスターシート名を生成してシートを取得
 * @param {Spreadsheet} spreadsheet - スプレッドシートオブジェクト
 * @param {string} dateStr - 日付文字列（YYYY/MM/DD形式）
 * @returns {Sheet|null} マスターシートオブジェクト
 */
function getMasterSheet(spreadsheet, dateStr) {
  try {
    console.log('=== マスターシート取得開始 ===');
    console.log('対象日付:', dateStr);
    
    var date = new Date(dateStr);
    var year = date.getFullYear();
    var month = date.getMonth() + 1;
    
    console.log('年:', year, '月:', month);
    
    // マスターシート名を生成（例: 2507月_売上）
    var yearSuffix = String(year).slice(-2);
    var monthFormatted = month < 10 ? '0' + month : String(month);
    var sheetName = yearSuffix + monthFormatted + '月_売上';
    
    console.log('生成されたシート名:', sheetName);
    
    // 利用可能なシート一覧を取得
    var allSheets = spreadsheet.getSheets();
    console.log('スプレッドシート内の全シート:');
    for (var i = 0; i < allSheets.length; i++) {
      var existingSheetName = allSheets[i].getName();
      console.log('  "' + existingSheetName + '"' + (existingSheetName === sheetName ? ' ← 一致！' : ''));
    }
    
    var masterSheet = spreadsheet.getSheetByName(sheetName);
    if (!masterSheet) {
      console.log('✗ マスターシートが見つかりません');
      
      // 類似するシート名を探索
      console.log('類似するシート名を検索中...');
      for (var i = 0; i < allSheets.length; i++) {
        var existingSheetName = allSheets[i].getName();
        if (existingSheetName.indexOf(yearSuffix) > -1 || 
            existingSheetName.indexOf(monthFormatted + '月') > -1 ||
            existingSheetName.indexOf('売上') > -1) {
          console.log('  類似候補: "' + existingSheetName + '"');
        }
      }
      
      showError('エラー: 書き込み先のマスターシート（' + sheetName + '）が見つかりません。\n\n' +
               'スプレッドシート内の利用可能なシートを確認してください。\n' +
               '詳細はGASエディターのログを確認してください。');
      return null;
    }
    
    console.log('✓ マスターシート取得成功:', sheetName);
    return masterSheet;
    
  } catch (error) {
    console.error('マスターシート取得でエラー:', error);
    showError('エラー: 日付「' + dateStr + '」からマスターシート名を生成できませんでした。\n詳細: ' + error.message);
    return null;
  }
}

/**
 * 日付重複チェック（エラーチェック1）
 * @param {Sheet} masterSheet - マスターシート
 * @param {Object} block - 報告ブロック
 * @returns {string} 'proceed'（続行）, 'abort'（中断）
 */
function checkDateDuplicate(masterSheet, block) {
  try {
    console.log('日付重複チェック開始: ' + block.store + ' - ' + block.date);
    
    // 店舗行と日付列を特定
    var storeRowInfo = findStoreRows(masterSheet, block.store);
    var dateCol = findDateColumn(masterSheet, block.date);
    
    if (!storeRowInfo || !dateCol) {
      console.log('店舗または日付が見つからないため、重複チェックをスキップ');
      return 'proceed';
    }
    
    // 主要項目（売上、仕入）の既存データをチェック
    var checkItems = ['売上', '仕入'];
    var hasExistingData = false;
    
    for (var j = 0; j < checkItems.length; j++) {
      var itemName = checkItems[j];
      var mappedItemName = mapItemName(itemName);
      var itemRow = storeRowInfo[mappedItemName];
      if (itemRow) {
        var cellValue = masterSheet.getRange(itemRow, dateCol).getValue();
        if (cellValue && cellValue !== 0) {
          hasExistingData = true;
          break;
        }
      }
    }
    
    if (hasExistingData) {
      // 確認ダイアログを表示
      var ui = SpreadsheetApp.getUi();
      var formattedDate = formatDateForDisplay(block.date);
      
      var response = ui.alert(
        '日付重複の確認',
        '「' + block.store + '」の「' + formattedDate + '」のデータは既に入力されています。\n\n' +
        'これは報告内容の修正ですか？\n' +
        '「はい」を選択すると、既存のデータを新しいデータで上書きします。\n' +
        '「いいえ」を選択すると、処理を中断します。',
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        console.log('ユーザーが上書きを承認');
        return 'proceed';
      } else {
        showInfo('処理は中断されました。データは更新されていません。');
        return 'abort';
      }
    }
    
    return 'proceed';
    
  } catch (error) {
    console.error('日付重複チェックエラー:', error);
    showError('エラー: 日付重複チェック中に問題が発生しました。\n詳細: ' + error.message);
    return 'abort';
  }
}

/**
 * 前日データとの完全重複チェック（エラーチェック2）
 * @param {Sheet} masterSheet - マスターシート
 * @param {Object} block - 報告ブロック
 * @returns {boolean} true: エラー検知（処理中断）, false: 正常
 */
function checkPreviousDayDuplicate(masterSheet, block) {
  try {
    console.log('前日重複チェック開始: ' + block.store + ' - ' + block.date);
    
    // 店舗行と日付列を特定
    var storeRowInfo = findStoreRows(masterSheet, block.store);
    var todayCol = findDateColumn(masterSheet, block.date);
    
    if (!storeRowInfo || !todayCol) {
      console.log('店舗または日付が見つからないため、前日重複チェックをスキップ');
      return false;
    }
    
    // 前日の列を取得
    var yesterdayCol = todayCol - 1;
    if (yesterdayCol < 1) {
      console.log('前日の列が存在しないため、前日重複チェックをスキップ');
      return false;
    }
    
    // チェック対象項目
    var checkItems = ['売上', '仕入'];
    var allMatching = true;
    var hasDataToCompare = false;
    
    for (var j = 0; j < checkItems.length; j++) {
      var itemName = checkItems[j];
      var mappedItemName = mapItemName(itemName);
      var itemRow = storeRowInfo[mappedItemName];
      if (itemRow && block.data[itemName] !== undefined) {
        var todayValue = block.data[itemName];
        var yesterdayValue = masterSheet.getRange(itemRow, yesterdayCol).getValue();
        
        if (yesterdayValue && yesterdayValue !== 0) {
          hasDataToCompare = true;
          if (todayValue !== yesterdayValue) {
            allMatching = false;
            break;
          }
        } else {
          allMatching = false;
          break;
        }
      }
    }
    
    if (hasDataToCompare && allMatching) {
      // エラーを検知
      var formattedDate = formatDateForDisplay(block.date);
      var yesterdayDate = new Date(block.date);
      yesterdayDate.setDate(yesterdayDate.getDate() - 1);
      var formattedYesterday = formatDateForDisplay(yesterdayDate.getFullYear() + '/' + (yesterdayDate.getMonth() + 1) + '/' + yesterdayDate.getDate());
      
      showError(
        '【エラー箇所】\n' +
        '店舗: ' + block.store + '\n' +
        '日付: ' + formattedDate + '\n\n' +
        '【エラー内容】\n' +
        '「売上」と「仕入」の金額が、前日（' + formattedYesterday + '）のデータと全く同じです。\n' +
        '報告内容が正しいか確認してください。\n\n' +
        '処理を中断しました。'
      );
      
      return true; // エラー検知
    }
    
    return false; // 正常
    
  } catch (error) {
    console.error('前日重複チェックエラー:', error);
    showError('エラー: 前日データとの重複チェック中に問題が発生しました。\n詳細: ' + error.message);
    return true; // エラー扱い
  }
}

/**
 * 店舗名を正規化（「【店舗名】○○店」→「○○」に変換）
 * スペースが含まれる場合も考慮
 * @param {string} storeName - 元の店舗名
 * @returns {string} 正規化された店舗名
 */
function normalizeStoreName(storeName) {
  var normalized = storeName
    .replace(/【店舗名】/g, '')       // 【店舗名】を削除
    .replace(/\s+/g, '')              // すべてのスペース（全角・半角）を削除
    .replace(/店$/, '')               // 末尾の「店」を削除
    .trim();                          // 前後の空白を削除
  
  // 特別な店舗名のマッピング
  var storeMapping = {
    '野木': 'マルタツ野木',
    '小山': 'マルタツ小山',
    '結城': 'マルタツ結城',
    '藤岡': 'マルタツ藤岡',
    '真岡': 'マルタツ真岡',
    '羽川': 'マルタツ羽川',
    '高崎': 'マルタツ高崎',
    'クロリ': 'クロリ小山',
    'クロリ小山店': 'クロリ小山',
    'クロリ小山工場佐野': 'クロリ小山',
    '晴れパン': 'ハレパン小山野木真岡',
    'ハレパン': 'ハレパン小山野木真岡',
    '寅ジロー': '寅ジロー小山'
  };
  
  // マッピングがある場合は適用
  if (storeMapping[normalized]) {
    var mapped = storeMapping[normalized];
    console.log('店舗名マッピング適用: ' + normalized + ' → ' + mapped);
    normalized = mapped;
  }
  
  console.log('店舗名正規化: "' + storeName + '" → "' + normalized + '"');
  return normalized;
}

/**
 * 項目名をマスターシート形式にマッピング
 * @param {string} itemName - 転記データの項目名
 * @returns {string} マスターシートの項目名
 */
function mapItemName(itemName) {
  var mapping = {
    '売上': '当日売上',
    '仕入': '当日仕入費',
    '仕入費': '当日仕入費',
    '人件費': '当日人件費',
    'P/A': 'P/A',
    '社員': '社員'
  };
  
  var mapped = mapping[itemName] || itemName;
  if (mapped !== itemName) {
    console.log('項目名マッピング: ' + itemName + ' → ' + mapped);
  }
  return mapped;
}

/**
 * マスターシートから店舗の各項目行を検索
 * @param {Sheet} masterSheet - マスターシート
 * @param {string} storeName - 店舗名
 * @returns {Object|null} 項目名をキー、行番号を値とするオブジェクト
 */
function findStoreRows(masterSheet, storeName) {
  try {
    console.log('=== 店舗行検索開始 ===');
    console.log('元の店舗名:', storeName);
    
    // 店舗名を正規化
    var normalizedStoreName = normalizeStoreName(storeName);
    console.log('正規化後の店舗名:', normalizedStoreName);
    
    var lastRow = masterSheet.getLastRow();
    console.log('マスターシートの最終行:', lastRow);
    
    var aColumn = masterSheet.getRange('A1:A' + lastRow).getValues();
    var bColumn = masterSheet.getRange('B1:B' + lastRow).getValues();
    
    var storeRows = {};
    var foundAnyMatch = false;
    
    console.log('A列とB列をスキャン中...');
    for (var i = 0; i < aColumn.length; i++) {
      var aValue = String(aColumn[i][0]).trim();
      var bValue = String(bColumn[i][0]).trim();
      
      // デバッグ: 最初の10行と一致する行をログ出力
      if (i < 10 || aValue === normalizedStoreName) {
        console.log('  行' + (i+1) + ': A列="' + aValue + '", B列="' + bValue + '"' + 
                   (aValue === normalizedStoreName ? ' ← 一致！' : ''));
      }
      
      // A列が正規化された店舗名と一致する行を検索
      if (aValue === normalizedStoreName) {
        storeRows[bValue] = i + 1; // 行番号（1ベース）
        foundAnyMatch = true;
      }
    }
    
    if (foundAnyMatch) {
      console.log('✓ 店舗「' + normalizedStoreName + '」の行情報:', storeRows);
      console.log('見つかった項目数:', Object.keys(storeRows).length);
      return storeRows;
    } else {
      console.log('✗ 店舗「' + normalizedStoreName + '」に一致する行が見つかりませんでした');
      
      // 類似する店舗名を探索（部分一致）
      console.log('類似する店舗名を検索中...');
      var similarStores = [];
      
      for (var i = 0; i < aColumn.length; i++) {
        var aValue = String(aColumn[i][0]).trim();
        if (aValue && aValue.indexOf(normalizedStoreName) > -1) {
          console.log('  部分一致候補: "' + aValue + '" (行' + (i+1) + ')');
          similarStores.push(aValue);
        }
      }
      
      // 部分一致する店舗が1つだけ見つかった場合は自動採用
      if (similarStores.length === 1) {
        var autoMatchStore = similarStores[0];
        console.log('🔄 自動マッチング: ' + normalizedStoreName + ' → ' + autoMatchStore);
        
        // 再検索
        var autoStoreRows = {};
        for (var i = 0; i < aColumn.length; i++) {
          var aValue = String(aColumn[i][0]).trim();
          var bValue = String(bColumn[i][0]).trim();
          
          if (aValue === autoMatchStore) {
            autoStoreRows[bValue] = i + 1;
          }
        }
        
        if (Object.keys(autoStoreRows).length > 0) {
          console.log('✓ 自動マッチング成功:', autoStoreRows);
          return autoStoreRows;
        }
      }
      
      return null;
    }
    
  } catch (error) {
    console.error('店舗行検索でエラー:', error);
    return null;
  }
}

/**
 * マスターシートから日付列を検索
 * @param {Sheet} masterSheet - マスターシート
 * @param {string} dateStr - 日付文字列
 * @returns {number|null} 列番号（1ベース）
 */
function findDateColumn(masterSheet, dateStr) {
  try {
    console.log('=== 日付列検索開始 ===');
    console.log('検索対象日付文字列:', dateStr);
    
    var targetDate = new Date(dateStr);
    // 時刻を00:00:00にリセット
    targetDate.setHours(0, 0, 0, 0);
    
    console.log('検索対象日付（リセット後）:', targetDate);
    console.log('年:', targetDate.getFullYear(), '月:', targetDate.getMonth() + 1, '日:', targetDate.getDate());
    
    var lastCol = masterSheet.getLastColumn();
    console.log('マスターシートの最終列:', lastCol);
    
    // C1(列3)からAG1(列33)までを検索範囲とする
    var searchStartCol = 3; // C列
    var searchEndCol = Math.min(33, lastCol); // AG列または最終列のいずれか小さい方
    var headerRow = masterSheet.getRange(1, searchStartCol, 1, searchEndCol - searchStartCol + 1).getValues()[0];
    console.log('1行目のデータ取得完了（C1～AG1の範囲）');
    
    console.log('各列をスキャン中...');
    for (var i = 0; i < headerRow.length; i++) {
      var cellValue = headerRow[i];
      var cellType = typeof cellValue;
      var actualColumn = i + searchStartCol; // 実際の列番号を計算
      
      // 詳細ログ（最初の20列まで）
      if (i < 20) {
        if (cellValue instanceof Date) {
          console.log('  列' + actualColumn + ': ' + cellValue + ' (Dateオブジェクト)');
        } else if (cellType === 'number') {
          var serialDate = new Date((cellValue - 25569) * 86400 * 1000);
          console.log('  列' + actualColumn + ': ' + cellValue + ' (数値, 日付変換: ' + serialDate + ')');
        } else {
          console.log('  列' + actualColumn + ': "' + cellValue + '" (' + cellType + ')');
        }
      }
      
      // 日付マッチング処理
      if (cellValue instanceof Date) {
        var cellDate = new Date(cellValue);
        cellDate.setHours(0, 0, 0, 0);
        
        if (cellDate.getTime() === targetDate.getTime()) {
          console.log('✓ 日付が一致（Dateオブジェクト）: 列' + actualColumn);
          return actualColumn;
        }
      } else if (typeof cellValue === 'number') {
        // Excelのシリアル値を日付に変換
        var serialDate = new Date((cellValue - 25569) * 86400 * 1000);
        serialDate.setHours(0, 0, 0, 0);
        
        if (serialDate.getFullYear() === targetDate.getFullYear() &&
            serialDate.getMonth() === targetDate.getMonth() &&
            serialDate.getDate() === targetDate.getDate()) {
          console.log('✓ 日付が一致（シリアル値）: 列' + actualColumn);
          return actualColumn;
        }
      } else if (typeof cellValue === 'string' && cellValue) {
        try {
          var parsedDate = new Date(cellValue);
          parsedDate.setHours(0, 0, 0, 0);
          
          if (!isNaN(parsedDate.getTime()) &&
              parsedDate.getTime() === targetDate.getTime()) {
            console.log('✓ 日付が一致（文字列）: 列' + actualColumn);
            return actualColumn;
          }
        } catch (e) {
          // 日付として解析できない場合は無視
        }
      }
    }
    
    console.log('✗ 日付「' + dateStr + '」に対応する列が見つかりませんでした');
    
    // AF列（129列目）付近を重点的にチェック
    if (lastCol >= 129) {
      console.log('=== AF列付近の詳細チェック ===');
      for (var col = 127; col <= 131 && col <= lastCol; col++) {
        var cellValue = headerRow[col - 1];
        console.log('列' + col + ':', cellValue, '(型:', typeof cellValue, ')');
        
        if (typeof cellValue === 'number') {
          var serialDate = new Date((cellValue - 25569) * 86400 * 1000);
          console.log('  → 日付変換:', serialDate);
        }
      }
    }
    
    return null;
    
  } catch (error) {
    console.error('日付列検索でエラー:', error);
    return null;
  }
}

/**
 * マスターシートにデータを書き込み
 * @param {Sheet} masterSheet - マスターシート
 * @param {Object} block - 報告ブロック
 */
function writeDataToMasterSheet(masterSheet, block) {
  try {
    console.log('=== データ書き込み開始 ===');
    console.log('対象店舗:', block.store);
    console.log('対象日付:', block.date);
    console.log('転記データ:', block.data);
    
    // 店舗行情報を取得
    console.log('--- 店舗行検索開始 ---');
    var storeRowInfo = findStoreRows(masterSheet, block.store);
    
    if (!storeRowInfo) {
      // 詳細エラー情報を出力
      logMasterSheetStructure(masterSheet, block.store);
      throw new Error('店舗「' + block.store + '」の行が見つかりません。マスターシートのA列を確認してください。');
    }
    
    console.log('店舗行検索成功:', storeRowInfo);
    
    // 日付列を取得
    console.log('--- 日付列検索開始 ---');
    var dateCol = findDateColumn(masterSheet, block.date);
    
    if (!dateCol) {
      // 詳細エラー情報を出力
      logDateHeaderStructure(masterSheet, block.date);
      throw new Error('日付「' + block.date + '」の列が見つかりません。マスターシートの1行目を確認してください。');
    }
    
    console.log('日付列検索成功: 列' + dateCol);
    
    var writtenCount = 0;
    
    // 転記データが空の場合のチェック
    var dataItemCount = Object.keys(block.data).length;
    if (dataItemCount === 0) {
      console.log('⚠ 警告: 転記可能なデータが存在しません。データ解析を確認してください。');
      console.log('店舗名:', block.store);
      console.log('日付:', block.date);
      return; // 書き込みをスキップ
    }
    
    console.log('転記予定データ項目数:', dataItemCount);
    
    // 各データ項目を書き込み
    console.log('--- データ項目書き込み開始 ---');
    for (var itemName in block.data) {
      var value = block.data[itemName];
      
      // 項目名をマスターシート形式にマッピング
      var mappedItemName = mapItemName(itemName);
      var itemRow = storeRowInfo[mappedItemName];
      
      if (itemRow) {
        masterSheet.getRange(itemRow, dateCol).setValue(value);
        console.log('✓ 書き込み成功: ' + itemName + '(' + mappedItemName + ') = ' + value + ' (行' + itemRow + ', 列' + dateCol + ')');
        writtenCount++;
      } else {
        console.log('⚠ 警告: 項目「' + itemName + '」(マッピング後:「' + mappedItemName + '」)に対応する行が見つかりませんでした');
        console.log('  利用可能な項目:', Object.keys(storeRowInfo));
      }
    }
    
    if (writtenCount === 0) {
      console.log('⚠ 警告: 書き込み可能なデータ項目が1つも見つかりませんでした');
      console.log('原因: マスターシートの項目名と転記データの項目名が一致していない可能性があります');
    }
    
    console.log('=== データ書き込み完了 ===');
    console.log(block.store + 'のデータ書き込み完了: ' + writtenCount + '項目');
    
  } catch (error) {
    console.error('=== データ書き込みエラー ===');
    console.error('エラー詳細:', error.message);
    throw error;
  }
}

/**
 * マスターシートの構造をログ出力（デバッグ用）
 * @param {Sheet} masterSheet - マスターシート
 * @param {string} targetStore - 検索対象の店舗名
 */
function logMasterSheetStructure(masterSheet, targetStore) {
  try {
    console.log('=== マスターシート構造分析 ===');
    console.log('シート名:', masterSheet.getName());
    
    var normalizedTarget = normalizeStoreName(targetStore);
    console.log('検索対象店舗（正規化後）:', normalizedTarget);
    
    var lastRow = Math.min(masterSheet.getLastRow(), 50); // 最大50行まで
    var aColumn = masterSheet.getRange('A1:A' + lastRow).getValues();
    var bColumn = masterSheet.getRange('B1:B' + lastRow).getValues();
    
    console.log('A列の内容（最初の50行）:');
    var uniqueStores = {};
    for (var i = 0; i < aColumn.length; i++) {
      var aValue = String(aColumn[i][0]).trim();
      var bValue = String(bColumn[i][0]).trim();
      
      if (aValue) {
        uniqueStores[aValue] = true;
        console.log('  行' + (i+1) + ': A列="' + aValue + '", B列="' + bValue + '"');
        
        // 部分一致チェック
        if (aValue.indexOf(normalizedTarget) > -1 || normalizedTarget.indexOf(aValue) > -1) {
          console.log('    → 部分一致の可能性あり！');
        }
      }
    }
    
    console.log('A列で見つかった店舗名一覧:');
    for (var store in uniqueStores) {
      console.log('  "' + store + '"');
    }
    
  } catch (error) {
    console.error('マスターシート構造分析でエラー:', error);
  }
}

/**
 * 日付ヘッダーの構造をログ出力（デバッグ用）
 * @param {Sheet} masterSheet - マスターシート
 * @param {string} targetDate - 検索対象の日付
 */
function logDateHeaderStructure(masterSheet, targetDate) {
  try {
    console.log('=== 日付ヘッダー構造分析 ===');
    console.log('検索対象日付:', targetDate);
    console.log('検索対象日付（Dateオブジェクト）:', new Date(targetDate));
    
    // C1(列3)からAG1(列33)までの範囲で最大20列まで
    var searchStartCol = 3; // C列
    var lastCol = Math.min(masterSheet.getLastColumn(), 33); // AG列または最終列のいずれか小さい方
    var searchEndCol = Math.min(searchStartCol + 19, lastCol); // 最大20列まで
    var headerRow = masterSheet.getRange(1, searchStartCol, 1, searchEndCol - searchStartCol + 1).getValues()[0];
    
    console.log('1行目の内容（C列から最大20列）:');
    for (var i = 0; i < headerRow.length; i++) {
      var cellValue = headerRow[i];
      var cellType = typeof cellValue;
      var actualColumn = i + searchStartCol; // 実際の列番号を計算
      
      if (cellValue instanceof Date) {
        console.log('  列' + actualColumn + ': ' + cellValue + ' (Dateオブジェクト)');
      } else if (cellType === 'number') {
        // シリアル値の可能性
        var serialDate = new Date((cellValue - 25569) * 86400 * 1000);
        console.log('  列' + actualColumn + ': ' + cellValue + ' (数値, 日付変換: ' + serialDate + ')');
      } else {
        console.log('  列' + actualColumn + ': "' + cellValue + '" (' + cellType + ')');
      }
    }
    
  } catch (error) {
    console.error('日付ヘッダー構造分析でエラー:', error);
  }
}

/**
 * 日付を表示用にフォーマット
 * @param {string} dateStr - 日付文字列
 * @returns {string} フォーマット済み日付
 */
function formatDateForDisplay(dateStr) {
  try {
    var date = new Date(dateStr);
    var month = date.getMonth() + 1;
    var day = date.getDate();
    return month + '月' + day + '日';
  } catch (error) {
    return dateStr;
  }
}

/**
 * エラーメッセージを表示
 * @param {string} message - エラーメッセージ
 */
function showError(message) {
  console.error('エラー:', message);
  SpreadsheetApp.getUi().alert('エラー', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 成功メッセージを表示
 * @param {string} message - 成功メッセージ
 */
function showSuccess(message) {
  console.log('成功:', message);
  SpreadsheetApp.getUi().alert('処理完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 情報メッセージを表示
 * @param {string} message - 情報メッセージ
 */
function showInfo(message) {
  console.log('情報:', message);
  SpreadsheetApp.getUi().alert('お知らせ', message, SpreadsheetApp.getUi().ButtonSet.OK);
}