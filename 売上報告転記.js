/**
 * 売上報告転記.gs
 * 売上報告シートのB列からマスターシートへの転記
 */

/**
 * 売上報告シートからマスターシートへ転記
 * @param {boolean} batchMode - true: 一括転記（確認なし）、false: 確認あり
 */
function transferReportToMaster(batchMode) {
  try {
    console.log('=== 売上報告転記処理開始 ===');
    console.log('モード:', batchMode ? '一括転記' : '確認あり');
    
    // 1. スプレッドシートとソースシートを取得
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = spreadsheet.getSheetByName('売上報告');
    
    if (!sourceSheet) {
      showError('エラー: 「売上報告」シートが見つかりません。');
      return;
    }
    
    // 2. B列の全データを取得（B2から開始）
    var lastRow = sourceSheet.getLastRow();
    if (lastRow < 2) {
      showError('エラー: 売上報告シートにデータが存在しません。');
      return;
    }
    
    var sourceData = sourceSheet.getRange('B2:B' + lastRow).getValues();
    var flatData = [];
    for (var i = 0; i < sourceData.length; i++) {
      flatData.push(sourceData[i][0]);
    }
    console.log('取得データ件数:', flatData.length);
    
    // 3. データを店舗ごとの報告ブロックに構造化
    var reportBlocks = parseReportDataFromB(flatData);
    console.log('解析された報告ブロック数:', reportBlocks.length);
    
    if (reportBlocks.length === 0) {
      showError('エラー: 有効な報告データが見つかりません。\nB列のデータ形式を確認してください。');
      return;
    }
    
    // 4. 各報告ブロックを処理
    var processedCount = 0;
    var errorCount = 0;
    
    for (var i = 0; i < reportBlocks.length; i++) {
      var block = reportBlocks[i];
      console.log('=== 処理中: ' + block.store + ' - ' + block.date + ' ===');
      
      // データが空の場合はスキップ
      if (Object.keys(block.data).length === 0) {
        console.log('⚠ データが空のためスキップ');
        continue;
      }
      
      // 確認モードの場合、ダイアログを表示
      if (!batchMode) {
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
          console.log('この店舗をスキップ');
          continue;
        }
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
        
        // 日付重複チェック（一括モードでは自動承認）
        var duplicateCheckResult = batchMode ? 
          checkDateDuplicateBatch(masterSheet, block) : 
          checkDateDuplicate(masterSheet, block);
        
        if (duplicateCheckResult === 'abort') {
          console.log('日付重複チェックで処理中断');
          return;
        }
        
        // 前日データとの完全重複チェック
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
        
        if (!batchMode) {
          var errorResponse = ui.alert(
            'エラー発生',
            '店舗: ' + block.store + '\n' +
            '日付: ' + block.date + '\n' +
            'エラー: ' + error.message + '\n\n' +
            '処理を続行しますか？',
            ui.ButtonSet.YES_NO
          );
          
          if (errorResponse === ui.Button.NO) {
            console.log('ユーザーがエラー後の処理継続を拒否');
            return;
          }
        }
      }
    }
    
    // 5. 成功通知
    if (processedCount > 0) {
      var resultMessage = processedCount + '件の店舗データの転記が完了しました。';
      if (errorCount > 0) {
        resultMessage += '\n' + errorCount + '件でエラーが発生しました。';
      }
      showSuccess(resultMessage);
    } else {
      showInfo('転記対象のデータが見つかりませんでした。');
    }
    
  } catch (error) {
    console.error('転記処理全体でエラー:', error);
    showError('システムエラーが発生しました。\n詳細: ' + error.message);
  }
}

/**
 * B列のデータを店舗ごとの報告ブロックに解析
 * @param {Array} sourceData - B列の全データ
 * @returns {Array} 解析された報告ブロックの配列
 */
function parseReportDataFromB(sourceData) {
  var blocks = [];
  var currentBlock = null;
  
  for (var i = 0; i < sourceData.length; i++) {
    var cellValue = String(sourceData[i]).trim();
    
    // 空行はブロックの区切り
    if (!cellValue) {
      if (currentBlock && currentBlock.date && currentBlock.store) {
        blocks.push(currentBlock);
        currentBlock = null;
      }
      continue;
    }
    
    // 日付行を検出（新しいブロックの開始）
    if (cellValue.match(/^\d{2}:\d{2}\s+/) || 
        cellValue.indexOf('日付') > -1 || 
        cellValue.indexOf('日時') > -1) {
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
      
      // 日付を抽出
      currentBlock.date = extractDateFromLine(cellValue);
      console.log('新しいブロック開始 - 日付:', currentBlock.date);
      continue;
    }
    
    // 現在のブロックが存在しない場合はスキップ
    if (!currentBlock) continue;
    
    // 店舗名を検出
    if (cellValue.indexOf('店舗') > -1 && !currentBlock.store) {
      currentBlock.store = cellValue;
      console.log('店舗名検出:', currentBlock.store);
      continue;
    }
    
    // データ項目を解析
    var itemData = extractItemData(cellValue);
    if (itemData) {
      console.log('データ項目検出:', itemData.name, '=', itemData.value);
      currentBlock.data[itemData.name] = itemData.value;
    }
  }
  
  // 最後のブロックを保存
  if (currentBlock && currentBlock.date && currentBlock.store) {
    blocks.push(currentBlock);
  }
  
  // デバッグ出力
  console.log('=== 解析結果 ===');
  blocks.forEach(function(block, index) {
    console.log('ブロック' + (index + 1) + ':');
    console.log('  店舗:', block.store);
    console.log('  日付:', block.date);
    console.log('  データ:', block.data);
  });
  
  return blocks;
}

/**
 * 行から日付を抽出
 * @param {string} line - 日付を含む行
 * @returns {string} YYYY/MM/DD形式の日付文字列
 */
function extractDateFromLine(line) {
  // 時刻とユーザー名を削除
  line = line.replace(/^\d{2}:\d{2}\s+[^\s]+\s+/, '');
  
  // 括弧内の曜日を削除
  line = line.replace(/[\(（][^）\)]*[\)）]/g, '');
  
  // YYYY/MM/DD形式
  var fullMatch = line.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (fullMatch) {
    return fullMatch[0];
  }
  
  // 月日形式（現在年を使用）- スペースも考慮
  var mdMatch = line.match(/(\d{1,2})月\s*(\d{1,2})日?/);
  if (mdMatch) {
    var year = new Date().getFullYear();
    var month = mdMatch[1].padStart(2, '0');
    var day = mdMatch[2].padStart(2, '0');
    return year + '/' + month + '/' + day;
  }
  
  return '';
}

/**
 * 行からデータ項目を抽出
 * @param {string} line - データ項目を含む行
 * @returns {Object|null} {name: 項目名, value: 値}
 */
function extractItemData(line) {
  // 円を含む数値を検索
  var yenMatch = line.match(/([0-9,]+)\s*円/);
  if (!yenMatch) return null;
  
  var value = parseInt(yenMatch[1].replace(/,/g, ''), 10);
  if (isNaN(value)) value = 0;
  
  // 項目名を判定
  var itemName = '';
  if (line.indexOf('売上') > -1) {
    itemName = '売上';
  } else if (line.indexOf('P/A') > -1) {
    itemName = 'P/A';
  } else if (line.indexOf('社員') > -1) {
    itemName = '社員';
  } else if (line.indexOf('仕入') > -1) {
    itemName = '仕入';
  } else if (line.indexOf('人件費') > -1) {
    // 人件費の場合、P/Aか社員かを判定
    if (line.indexOf('P/A') > -1) {
      itemName = 'P/A';
    } else if (line.indexOf('社員') > -1) {
      itemName = '社員';
    } else {
      itemName = '人件費';
    }
  }
  
  if (itemName) {
    return { name: itemName, value: value };
  }
  
  return null;
}