// シートオープン時にカスタムメニューを追加
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('📊 売上データ処理', [
    { name: '🔄 データ整形＋転記（すべて実行）', functionName: 'processSheet' },
    null, // セパレーター
    { name: '✏️ データ整形のみ（C列に出力）', functionName: 'formatDataOnly' },
    { name: '📝 転記のみ（C列→マスターシート）', functionName: 'runTransferOnly' },
    null, // セパレーター
    { name: '❓ 使い方・ヘルプ', functionName: 'showHelp' }
  ]);
}

function processSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("コピペ用");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("シート「コピペ用」が見つかりません。");
    return;
  }
  
  // A列の全データを取得（データがある最終行まで）
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    SpreadsheetApp.getUi().alert("シートにデータが存在しません。");
    return;
  }
  var data = sheet.getRange(1, 1, lastRow, 1).getValues();
  
  // デバッグ: 最初の50行を出力
  console.log("=== A列の生データ（最初の50行）===");
  for (var i = 0; i < Math.min(50, data.length); i++) {
    console.log("行" + (i+1) + ": [" + data[i][0] + "]");
  }
  
  // ①【アンカー行の検出】
  // 「日付」または「日時」を含む行のインデックスを記録
  var anchorIndices = [];
  for (var i = 0; i < data.length; i++) {
    var text = data[i][0].toString();
    // 時刻情報から始まる行も日付として扱う
    if (text.match(/^\d{2}:\d{2}\s+/) || text.indexOf("日付") > -1 || text.indexOf("日時") > -1) {
      anchorIndices.push(i);
    }
  }
  
  console.log("アンカー行インデックス:", anchorIndices);
  
  // ②【セットの切り出し】
  // 各アンカー行を起点として、次のアンカー行の直前までを1セットとする
  var sets = [];
  for (var j = 0; j < anchorIndices.length; j++) {
    var start = anchorIndices[j];
    var end = (j + 1 < anchorIndices.length) ? anchorIndices[j + 1] - 1 : data.length - 1;
    var setRows = data.slice(start, end + 1);
    sets.push({ rows: setRows });
  }
  
  console.log("セット数:", sets.length);
  
  // ③【セット内の並び替えおよび各セットごとの日付・店舗名の抽出】
  // セット内の並び替え用のキーワード順
  var desiredOrder = ["日時", "日付", "店舗", "担当者", "売上", "人件費", "仕入費"];
  
  // 出力時の店舗表示順（部分一致で判定）
  var storeOrder = ["マルキン三毳", "マルキン高崎", "マルキン土浦", "マルタツ羽川", "マルタツ結城",
                    "マルタツ小山", "マルタツ藤岡", "マルタツ真岡", "マルタツ野木", "マルタツ高崎",
                    "クロリ小山工場佐野", "クロリ", "ハレパン小山野木真岡", "晴れパン", "寅ジロー小山", "寅ジロー"];
  
  sets.forEach(function(setObj, setIndex) {
    var setRows = setObj.rows;
    console.log("=== セット" + (setIndex + 1) + "の処理 ===");
    
    // セット内の並び替え
    var orderedSet = [];
    var processedIndices = [];
    
    // 1. まず指定順序でキーワードに該当する行を収集
    desiredOrder.forEach(function(keyword) {
      for (var i = 0; i < setRows.length; i++) {
        if (processedIndices.indexOf(i) > -1) continue;
        
        var rowText = setRows[i][0].toString();
        var normalizedRowText = rowText.replace(/\s+/g, '');
        
        if (normalizedRowText.indexOf(keyword) > -1) {
          var processedText = processSpacesInAmount(rowText);
          orderedSet.push([processedText]);
          processedIndices.push(i);
          
          // 人件費の場合、次の行も確認（P/Aと社員が別行の場合）
          if (keyword === "人件費" && i + 1 < setRows.length) {
            var nextRow = setRows[i + 1][0].toString();
            if (nextRow.match(/社員\s*[0-9,]+\s*円/)) {
              var processedNext = processSpacesInAmount(nextRow);
              orderedSet.push([processedNext]);
              processedIndices.push(i + 1);
            }
          }
        }
      }
    });
    
    // 2. P/Aと社員の独立した行を収集
    for (var i = 0; i < setRows.length; i++) {
      if (processedIndices.indexOf(i) > -1) continue;
      
      var rowText = setRows[i][0].toString();
      if (rowText.match(/(P\/A|社員)\s*[0-9,]+\s*円/)) {
        var processedText = processSpacesInAmount(rowText);
        orderedSet.push([processedText]);
        processedIndices.push(i);
      }
    }
    
    // 3. 残りの行を追加
    for (var i = 0; i < setRows.length; i++) {
      if (processedIndices.indexOf(i) === -1) {
        var rowText = setRows[i][0].toString();
        if (rowText.trim()) {
          var processedText = processSpacesInAmount(rowText);
          orderedSet.push([processedText]);
        }
      }
    }
    
    setObj.orderedRows = orderedSet;
    
    // 【日付／日時】の抽出
    var setDate = null;
    for (var i = 0; i < setRows.length; i++) {
      var cellText = setRows[i][0].toString();
      if (cellText.match(/^\d{2}:\d{2}\s+/) || cellText.indexOf("日付") > -1 || cellText.indexOf("日時") > -1) {
        var parsedDate = parseDateFromText(cellText);
        if (parsedDate) {
          setDate = parsedDate;
          console.log("日付検出:", parsedDate);
          break;
        }
      }
    }
    
    // 日付が抽出できなかった場合は、遠い未来の日付（ソート時に後ろへ）をセット
    if (!setDate) {
      setDate = new Date(3000, 0, 1);
    }
    setObj.date = setDate;
    
    // 【店舗】の抽出
    var storeName = "";
    for (var i = 0; i < setRows.length; i++) {
      var cellText = setRows[i][0].toString();
      var normalizedText = cellText.replace(/\s+/g, '');
      if (normalizedText.indexOf("店舗") > -1) {
        storeName = cellText;
        console.log("店舗検出:", storeName);
        break;
      }
    }
    
    // 店舗名の順位を決定
    var rank = getStoreRank(storeName, storeOrder);
    setObj.storeRank = rank;
  });
  
  // ④【セット全体の並び替え】
  // 日付（古い順）でソートし、同じ日付の場合は店舗表示順に従う
  sets.sort(function(a, b) {
    var diff = a.date.getTime() - b.date.getTime();
    if (diff !== 0) {
      return diff;
    } else {
      return a.storeRank - b.storeRank;
    }
  });
  
  // ⑤【最終出力データの作成】
  // 各セットの並び替え済み行を連結し、セットごとに空行を挿入
  var output = [];
  sets.forEach(function(setObj) {
    output = output.concat(setObj.orderedRows);
    output.push([""]); // セット間に空白行を追加
  });
  
  // ⑥【出力先: シート「コピペ用」C列に出力】
  sheet.getRange(1, 3, output.length, 1).setValues(output);
  
  // === 🆕 転記処理を追加 ===
  try {
    // データ整形が正常に完了した場合のみ、転記処理を実行
    console.log('データ整形が完了しました。転記処理を開始します...');
    
    // 少し待機してから転記処理を実行（Googleスプレッドシートの更新を確実にするため）
    Utilities.sleep(1000);
    
    // 転記処理を実行
    transferToMasterSheet();
    
  } catch (error) {
    console.error('転記処理でエラーが発生しました:', error);
    SpreadsheetApp.getUi().alert(
      'エラー', 
      'データ整形は完了しましたが、転記処理中にエラーが発生しました。\n' +
      '「コピペ用」シートのデータを確認してから、再度実行してください。\n\n' +
      'エラー詳細: ' + error.message, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * データ整形のみを実行（転記なし）
 */
function formatDataOnly() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '📊 データ整形の実行',
    'A列のデータを整形してC列に出力します。\n' +
    '（マスターシートへの転記は行いません）\n\n' +
    '実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    try {
      // processSheetの処理から転記部分を除いたもの
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("コピペ用");
      if (!sheet) {
        ui.alert("⚠️ エラー", "シート「コピペ用」が見つかりません。", ui.ButtonSet.OK);
        return;
      }
      
      // A列の全データを取得（データがある最終行まで）
      var lastRow = sheet.getLastRow();
      if (lastRow < 1) {
        ui.alert("⚠️ エラー", "シートにデータが存在しません。", ui.ButtonSet.OK);
        return;
      }
      var data = sheet.getRange(1, 1, lastRow, 1).getValues();
      
      // デバッグ: 最初の50行を出力
      console.log("=== A列の生データ（最初の50行）===");
      for (var i = 0; i < Math.min(50, data.length); i++) {
        console.log("行" + (i+1) + ": [" + data[i][0] + "]");
      }
      
      // ①【アンカー行の検出】
      var anchorIndices = [];
      for (var i = 0; i < data.length; i++) {
        var text = data[i][0].toString();
        // 時刻情報から始まる行も日付として扱う
        if (text.match(/^\d{2}:\d{2}\s+/) || text.indexOf("日付") > -1 || text.indexOf("日時") > -1) {
          anchorIndices.push(i);
        }
      }
      
      console.log("アンカー行インデックス:", anchorIndices);
      
      // ②【セットの切り出し】
      var sets = [];
      for (var j = 0; j < anchorIndices.length; j++) {
        var start = anchorIndices[j];
        var end = (j + 1 < anchorIndices.length) ? anchorIndices[j + 1] - 1 : data.length - 1;
        var setRows = data.slice(start, end + 1);
        sets.push({ rows: setRows });
      }
      
      console.log("セット数:", sets.length);
      
      // ③【セット内の並び替えおよび各セットごとの日付・店舗名の抽出】
      var desiredOrder = ["日時", "日付", "店舗", "担当者", "売上", "人件費", "仕入費"];
      var storeOrder = ["マルキン三毳", "マルキン高崎", "マルキン土浦", "マルタツ羽川", "マルタツ結城",
                        "マルタツ小山", "マルタツ藤岡", "マルタツ真岡", "マルタツ野木", "マルタツ高崎",
                        "クロリ小山工場佐野", "クロリ", "ハレパン小山野木真岡", "晴れパン", "寅ジロー小山", "寅ジロー"];
      
      sets.forEach(function(setObj, setIndex) {
        var setRows = setObj.rows;
        console.log("=== セット" + (setIndex + 1) + "の処理 ===");
        
        // セット内の並び替え
        var orderedSet = [];
        var processedIndices = [];
        
        // 1. まず指定順序でキーワードに該当する行を収集
        desiredOrder.forEach(function(keyword) {
          for (var i = 0; i < setRows.length; i++) {
            if (processedIndices.indexOf(i) > -1) continue;
            
            var rowText = setRows[i][0].toString();
            var normalizedRowText = rowText.replace(/\s+/g, '');
            
            if (normalizedRowText.indexOf(keyword) > -1) {
              var processedText = processSpacesInAmount(rowText);
              orderedSet.push([processedText]);
              processedIndices.push(i);
              
              // 人件費の場合、次の行も確認（P/Aと社員が別行の場合）
              if (keyword === "人件費" && i + 1 < setRows.length) {
                var nextRow = setRows[i + 1][0].toString();
                if (nextRow.match(/社員\s*[0-9,]+\s*円/)) {
                  var processedNext = processSpacesInAmount(nextRow);
                  orderedSet.push([processedNext]);
                  processedIndices.push(i + 1);
                }
              }
            }
          }
        });
        
        // 2. P/Aと社員の独立した行を収集
        for (var i = 0; i < setRows.length; i++) {
          if (processedIndices.indexOf(i) > -1) continue;
          
          var rowText = setRows[i][0].toString();
          if (rowText.match(/(P\/A|社員)\s*[0-9,]+\s*円/)) {
            var processedText = processSpacesInAmount(rowText);
            orderedSet.push([processedText]);
            processedIndices.push(i);
          }
        }
        
        // 3. 残りの行を追加
        for (var i = 0; i < setRows.length; i++) {
          if (processedIndices.indexOf(i) === -1) {
            var rowText = setRows[i][0].toString();
            if (rowText.trim()) {
              var processedText = processSpacesInAmount(rowText);
              orderedSet.push([processedText]);
            }
          }
        }
        
        setObj.orderedRows = orderedSet;
        
        // 【日付／日時】の抽出
        var setDate = null;
        for (var i = 0; i < setRows.length; i++) {
          var cellText = setRows[i][0].toString();
          if (cellText.match(/^\d{2}:\d{2}\s+/) || cellText.indexOf("日付") > -1 || cellText.indexOf("日時") > -1) {
            var parsedDate = parseDateFromText(cellText);
            if (parsedDate) {
              setDate = parsedDate;
              console.log("日付検出:", parsedDate);
              break;
            }
          }
        }
        
        // 日付が抽出できなかった場合は、遠い未来の日付（ソート時に後ろへ）をセット
        if (!setDate) {
          setDate = new Date(3000, 0, 1);
        }
        setObj.date = setDate;
        
        // 【店舗】の抽出
        var storeName = "";
        for (var i = 0; i < setRows.length; i++) {
          var cellText = setRows[i][0].toString();
          var normalizedText = cellText.replace(/\s+/g, '');
          if (normalizedText.indexOf("店舗") > -1) {
            storeName = cellText;
            console.log("店舗検出:", storeName);
            break;
          }
        }
        
        // 店舗名の順位を決定
        var rank = getStoreRank(storeName, storeOrder);
        setObj.storeRank = rank;
      });
      
      // ④【セット全体の並び替え】
      sets.sort(function(a, b) {
        var diff = a.date.getTime() - b.date.getTime();
        if (diff !== 0) {
          return diff;
        } else {
          return a.storeRank - b.storeRank;
        }
      });
      
      // ⑤【最終出力データの作成】
      var output = [];
      sets.forEach(function(setObj) {
        output = output.concat(setObj.orderedRows);
        output.push([""]); // セット間に空白行を追加
      });
      
      // ⑥【出力先: シート「コピペ用」C列に出力】
      sheet.getRange(1, 3, output.length, 1).setValues(output);
      
      ui.alert(
        '✅ 完了',
        'データ整形が完了しました。\nC列に整形済みデータを出力しました。',
        ui.ButtonSet.OK
      );
      
    } catch (error) {
      console.error('データ整形でエラー:', error);
      ui.alert(
        '⚠️ エラー',
        'データ整形中にエラーが発生しました。\n\nエラー詳細: ' + error.message,
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * 転記処理のみを単独で実行したい場合のメニュー関数
 */
function runTransferOnly() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '📝 転記処理の実行',
    '「コピペ用」シートのC列データをマスターシートに転記します。\n' +
    'データが正しく整形されていることを確認してから実行してください。\n\n' +
    '実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    transferToMasterSheet();
  }
}

/**
 * ヘルプを表示
 */
function showHelp() {
  var ui = SpreadsheetApp.getUi();
  var helpMessage = 
    '📊 売上データ処理システムの使い方\n\n' +
    '【基本的な流れ】\n' +
    '1. A列にLINE売上報告データが自動で入力されます\n' +
    '2. 「データ整形＋転記」を実行すると、自動でマスターシートまで転記されます\n\n' +
    '【各メニューの説明】\n' +
    '🔄 データ整形＋転記（すべて実行）\n' +
    '　→ A列のデータを整形してC列に出力し、さらにマスターシートへ転記します\n\n' +
    '✏️ データ整形のみ（C列に出力）\n' +
    '　→ A列のデータを整形してC列に出力します（転記は行いません）\n' +
    '　→ 整形結果を確認したい場合に使用します\n\n' +
    '📝 転記のみ（C列→マスターシート）\n' +
    '　→ C列の整形済みデータをマスターシートへ転記します\n' +
    '　→ 整形後に手動で修正してから転記したい場合に使用します\n\n' +
    '【注意事項】\n' +
    '• マスターシートは「2507月_売上」のような形式で命名されている必要があります\n' +
    '• 転記時は日付と店舗名でデータを照合します';
  
  ui.alert('❓ ヘルプ', helpMessage, ui.ButtonSet.OK);
}

// ── ヘルパー関数 ──

/**
 * 金額表記のスペースを処理する関数
 * 例: "社員 6,840 円" → "社員 6,840円"
 * @param {string} text - 処理対象のテキスト
 * @returns {string} 処理後のテキスト
 */
function processSpacesInAmount(text) {
  // 数値と円の間のスペース（複数含む）を削除
  text = text.replace(/([0-9,]+)\s+円/g, '$1円');
  
  // 全てのスペースを半角1つに統一
  text = text.replace(/\s+/g, ' ');
  
  return text.trim();
}

// 日付文字列の解析（「yyyy/mm/dd」形式または「月日」形式に対応）
function parseDateFromText(text) {
  // 時刻とユーザー名を削除（例: "08:27 a_ki"）
  text = text.replace(/^\d{2}:\d{2}\s+[^\s]+\s+/, '');
  
  // 括弧内の曜日情報を削除（全角・半角括弧に対応）
  text = text.replace(/[\(（][^）\)]*[\)）]/g, '');
  
  // フルパターン：例 "2025/4/11" や "2025年4月11日"
  var fullRegex = /(\d{4})[\/\-年](\d{1,2})[\/\-月]?(\d{1,2})日?/;
  var match = text.match(fullRegex);
  if (match) {
    var year = parseInt(match[1], 10);
    var month = parseInt(match[2], 10) - 1;
    var day = parseInt(match[3], 10);
    return new Date(year, month, day);
  }
  
  // 月日パターン：例 "7月29日" "7月29"
  var mdRegex = /(\d{1,2})月(\d{1,2})日?/;
  match = text.match(mdRegex);
  if (match) {
    var year = new Date().getFullYear();
    var month = parseInt(match[1], 10) - 1;
    var day = parseInt(match[2], 10);
    return new Date(year, month, day);
  }
  return null;
}

// 店舗名の順位を取得（storeOrderリストに基づく部分一致判定）
// ※「マルタツ野木」については、テキストが"野木"のみの場合も該当するようにする。
// 一致しなければ非常に大きな値を返して後ろにソート
function getStoreRank(text, storeOrder) {
  // スペースを削除してから比較
  text = text.replace(/\s+/g, '').trim();
  
  for (var i = 0; i < storeOrder.length; i++) {
    var candidate = storeOrder[i];
    if (candidate === "マルタツ野木") {
      // 「マルタツ野木」と文字列が含まれている場合、または「野木」が含まれ、かつ「マルタツ」が含まれない場合も該当とする
      if (text.indexOf(candidate) > -1 || (text.indexOf("野木") > -1 && text.indexOf("マルタツ") === -1)) {
        return i;
      }
    } else if (candidate === "晴れパン" || candidate === "ハレパン小山野木真岡") {
      // 晴れパン、ハレパンのいずれかが含まれていれば該当
      if (text.indexOf("晴れパン") > -1 || text.indexOf("ハレパン") > -1) {
        return i;
      }
    } else {
      if (text.indexOf(candidate) > -1) {
        return i;
      }
    }
  }
  return Number.MAX_SAFE_INTEGER;
}