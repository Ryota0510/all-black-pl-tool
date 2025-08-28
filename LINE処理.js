/**
 * LINE処理.gs
 * LINEメッセージの解析と整形処理
 */

/**
 * LINEメッセージをB列に展開・整形
 * @param {boolean} silent - 成功メッセージを表示しない場合はtrue
 */
function processLineMessage(silent) {
  var ui = SpreadsheetApp.getUi();
  
  try {
    console.log('=== LINE処理開始 ===');
    
    // 売上報告シートを取得
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var reportSheet = ss.getSheetByName('売上報告');
    
    if (!reportSheet) {
      ui.alert('⚠️ エラー', 'シート「売上報告」が見つかりません。', ui.ButtonSet.OK);
      return false;
    }
    
    // A1セルの内容を取得
    var lineMessage = reportSheet.getRange('A1').getValue();
    
    if (!lineMessage || lineMessage.toString().trim() === '') {
      ui.alert(
        '⚠️ 入力エラー',
        '売上報告シートのA1セルにLINEメッセージを貼り付けてください。',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    console.log('LINEメッセージ取得完了:', lineMessage.toString().substring(0, 100) + '...');
    
    // 改行で分割
    var lines = lineMessage.toString().split('\n');
    console.log('行数:', lines.length);
    
    // データをフィルタリング・整形
    var processedData = [];
    var excludePatterns = /天気|天候|単価|達成|弁当|食堂|予算|サービス|運営|小山売上|野木売上|真岡売上|佐野売上|本数|最高|気温|月間|ラスク|揚|問題|客数|組数|コメント|現金|新規|過不足/;
    var includePatterns = /【|】|日付|日時|店舗|担当|売上|仕入|人件|費|P\/A|社員/;
    
    // 現在のセット情報
    var currentSet = [];
    var hasDate = false;
    var hasStore = false;
    
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      
      // 空行は無視
      if (!line) continue;
      
      // 時刻情報から始まる行は新しいセットの開始
      if (line.match(/^\d{2}:\d{2}\s+/)) {
        // 前のセットを保存
        if (currentSet.length > 0 && hasDate && hasStore) {
          processedData = processedData.concat(sortSetData(currentSet));
          processedData.push(['']); // セット間に空行
        }
        currentSet = [];
        hasDate = false;
        hasStore = false;
      }
      
      // 除外パターンにマッチする行はスキップ
      if (excludePatterns.test(line)) {
        console.log('除外:', line);
        continue;
      }
      
      // 含むべきパターンにマッチする行のみ処理
      if (includePatterns.test(line)) {
        var processedLine = formatLineData(line);
        currentSet.push(processedLine);
        
        // 日付と店舗の存在チェック
        if (line.indexOf('日付') > -1 || line.indexOf('日時') > -1) hasDate = true;
        if (line.indexOf('店舗') > -1) hasStore = true;
        
        console.log('追加:', processedLine);
      }
    }
    
    // 最後のセットを保存
    if (currentSet.length > 0 && hasDate && hasStore) {
      processedData = processedData.concat(sortSetData(currentSet));
    }
    
    console.log('処理済みデータ数:', processedData.length);
    
    if (processedData.length === 0) {
      ui.alert(
        '⚠️ データなし',
        '処理可能なデータが見つかりませんでした。\n' +
        'LINEメッセージの形式を確認してください。',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // B列をクリア（B1は関数なのでB2以降）
    var lastRow = reportSheet.getLastRow();
    if (lastRow > 1) {
      reportSheet.getRange(2, 2, lastRow - 1, 1).clearContent();
    }
    
    // B2以降にデータを書き込み
    var outputData = [];
    for (var i = 0; i < processedData.length; i++) {
      outputData.push([processedData[i]]);
    }
    reportSheet.getRange(2, 2, outputData.length, 1).setValues(outputData);
    
    // 成功メッセージ（silentがfalseの場合のみ）
    if (!silent) {
      ui.alert(
        '✅ 完了',
        'LINEメッセージの処理が完了しました。\n' +
        'B列に' + outputData.length + '行のデータを出力しました。',
        ui.ButtonSet.OK
      );
    }
    
    console.log('=== LINE処理完了 ===');
    return true;
    
  } catch (error) {
    console.error('LINE処理エラー:', error);
    ui.alert(
      '⚠️ エラー',
      'LINEメッセージの処理中にエラーが発生しました。\n\n' +
      'エラー詳細: ' + error.message,
      ui.ButtonSet.OK
    );
    return false;
  }
}

/**
 * 行データを整形する関数
 * @param {string} line - 処理対象の行
 * @returns {string} 整形後の行
 */
function formatLineData(line) {
  var formatted = line;
  
  // 日付を○月○日形式に変換
  formatted = convertDateFormat(formatted);
  
  // 年のカンマを削除（2,025 → 2025）
  formatted = formatted.replace(/(\d),(\d{3})/g, '$1$2');
  
  // P/Aと社員の後にスペースを追加
  formatted = formatted.replace(/(P\/A|社員)(\d)/g, '$1 $2');
  
  // 数値のカンマを一旦削除
  formatted = formatted.replace(/(\d),(\d)/g, '$1$2');
  
  // 数値を適切にフォーマット（カンマ区切り）
  // 7桁
  formatted = formatted.replace(/\b(\d)(\d{3})(\d{3})\b/g, '$1,$2,$3');
  // 6桁
  formatted = formatted.replace(/\b(\d{3})(\d{3})\b/g, '$1,$2');
  // 5桁
  formatted = formatted.replace(/\b(\d{2})(\d{3})\b/g, '$1,$2');
  // 4桁
  formatted = formatted.replace(/\b(\d)(\d{3})\b/g, '$1,$2');
  
  return formatted;
}

/**
 * 日付を○月○日形式に変換
 * @param {string} text - 変換対象のテキスト
 * @returns {string} 変換後のテキスト
 */
function convertDateFormat(text) {
  // まずカンマを削除（2,025 → 2025）
  text = text.replace(/(\d),(\d)/g, '$1$2');
  
  // YYYY/MM/DD形式を○月○日に変換
  text = text.replace(/(\d{4})\/(\d{1,2})\/(\d{1,2})/g, function(match, year, month, day) {
    return parseInt(month, 10) + '月' + parseInt(day, 10) + '日';
  });
  
  // YYYY-MM-DD形式を○月○日に変換
  text = text.replace(/(\d{4})-(\d{1,2})-(\d{1,2})/g, function(match, year, month, day) {
    return parseInt(month, 10) + '月' + parseInt(day, 10) + '日';
  });
  
  // YYYY年MM月DD日形式を○月○日に変換
  text = text.replace(/(\d{4})年(\d{1,2})月(\d{1,2})日/g, function(match, year, month, day) {
    return parseInt(month, 10) + '月' + parseInt(day, 10) + '日';
  });
  
  return text;
}

/**
 * セット内のデータを並び替える
 * @param {Array} setData - セット内のデータ配列
 * @returns {Array} 並び替え済みのデータ配列
 */
function sortSetData(setData) {
  var desiredOrder = ["日時", "日付", "店舗", "担当者", "売上", "人件費", "P/A", "社員", "仕入費", "仕入"];
  var sorted = [];
  var processed = [];
  
  // 指定順序で並び替え
  desiredOrder.forEach(function(keyword) {
    for (var i = 0; i < setData.length; i++) {
      if (processed.indexOf(i) > -1) continue;
      
      var line = setData[i];
      var normalizedLine = line.replace(/\s+/g, '');
      
      if (normalizedLine.indexOf(keyword) > -1) {
        sorted.push(line);
        processed.push(i);
        
        // 人件費の次の行もチェック（P/Aや社員が別行の場合）
        if (keyword === "人件費" && i + 1 < setData.length) {
          var nextLine = setData[i + 1];
          if (nextLine.match(/(P\/A|社員)\s*[0-9,]+\s*円/)) {
            sorted.push(nextLine);
            processed.push(i + 1);
          }
        }
      }
    }
  });
  
  // 残りの行を追加
  for (var i = 0; i < setData.length; i++) {
    if (processed.indexOf(i) === -1) {
      sorted.push(setData[i]);
    }
  }
  
  return sorted;
}

/**
 * デバッグ用：売上報告シートの内容を確認
 */
function debugReportSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = ss.getSheetByName('売上報告');
  
  if (!reportSheet) {
    console.log('売上報告シートが見つかりません');
    return;
  }
  
  // A1の内容
  var a1Value = reportSheet.getRange('A1').getValue();
  console.log('A1の内容:', a1Value);
  
  // B列の内容（最初の20行）
  var lastRow = Math.min(reportSheet.getLastRow(), 20);
  if (lastRow > 0) {
    var bValues = reportSheet.getRange(1, 2, lastRow, 1).getValues();
    console.log('B列の内容:');
    for (var i = 0; i < bValues.length; i++) {
      console.log('B' + (i+1) + ':', bValues[i][0]);
    }
  }
}