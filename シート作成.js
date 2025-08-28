/**
 * 現在のシートのB1日付を修正
 */
function fixCurrentSheetBaseDate() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();
    
    // マスターシートかチェック
    var match = sheetName.match(/^(\d{2})(\d{2})月_売上$/);
    if (!match) {
      SpreadsheetApp.getUi().alert('現在のシートはマスターシートではありません。');
      return;
    }
    
    var year = 2000 + parseInt(match[1], 10);
    var month = parseInt(match[2], 10);
    
    // B1に正しい日付を設定
    setBaseDateCell(sheet, year, month);
    
    SpreadsheetApp.getUi().alert('B1の日付を ' + year + '/' + month + '/1 に修正しました。');
    
  } catch (error) {
    console.error('B1修正エラー:', error);
    SpreadsheetApp.getUi().alert('エラー: ' + error.message);
  }
}/**
 * シート作成.gs
 * マスターシートの作成と管理
 */

/**
 * マスターシート作成ダイアログを表示
 */
function showCreateMasterSheetDialog() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutputFromFile('シート作成ダイアログ')
      .setWidth(400)
      .setHeight(300);
  ui.showModalDialog(html, '📅 マスターシート作成');
}

/**
 * 現在月のマスターシートを作成
 */
function createCurrentMonthSheet() {
  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth() + 1;
  createMasterSheetForMonth(year, month);
}

/**
 * 翌月のマスターシートを作成
 */
function createNextMonthSheet() {
  var now = new Date();
  var nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  var year = nextMonth.getFullYear();
  var month = nextMonth.getMonth() + 1;
  createMasterSheetForMonth(year, month);
}

/**
 * 指定した年月のマスターシートを作成
 * @param {number} year - 年（4桁）
 * @param {number} month - 月（1-12）
 */
function createMasterSheetForMonth(year, month) {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // シート名を生成
    var yearSuffix = String(year).slice(-2);
    var monthFormatted = month < 10 ? '0' + month : String(month);
    var sheetName = yearSuffix + monthFormatted + '月_売上';
    
    // 既存チェック
    var existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      ui.alert(
        '⚠️ 既存シート',
        'シート「' + sheetName + '」は既に存在します。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // テンプレートシートを探す（前月のシートを使用）
    var templateSheet = findTemplateSheet(ss, year, month);
    
    if (templateSheet) {
      // テンプレートをコピー
      var newSheet = templateSheet.copyTo(ss);
      newSheet.setName(sheetName);
      
      // シートを先頭に移動
      ss.setActiveSheet(newSheet);
      ss.moveActiveSheet(1);
      
      // B1に基準日付を設定
      setBaseDateCell(newSheet, year, month);
      
      // 日付関数を設定
      setDateFormulas(newSheet, year, month);
      
      // 曜日行を設定（2行目）
      setWeekdayFormulas(newSheet, year, month);
      
      // データと数式を設定
      setupDataAndFormulas(newSheet, year, month);
      
      ui.alert(
        '✅ 作成完了',
        'マスターシート「' + sheetName + '」を作成しました。\n' +
        'テンプレート: ' + templateSheet.getName(),
        ui.ButtonSet.OK
      );
    } else {
      // テンプレートがない場合は新規作成
      var response = ui.alert(
        '📋 新規作成の確認',
        'テンプレートとなるシートが見つかりません。\n' +
        '空のマスターシート「' + sheetName + '」を作成しますか？',
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        createEmptyMasterSheet(ss, sheetName, year, month);
        ui.alert(
          '✅ 作成完了',
          '空のマスターシート「' + sheetName + '」を作成しました。\n' +
          '必要に応じて店舗や項目を追加してください。',
          ui.ButtonSet.OK
        );
      }
    }
    
  } catch (error) {
    console.error('シート作成エラー:', error);
    showError('マスターシート作成中にエラーが発生しました。\n' + error.message);
  }
}

/**
 * B1セルに基準日付を設定
 * @param {Sheet} sheet - 対象シート
 * @param {number} year - 年
 * @param {number} month - 月
 */
function setBaseDateCell(sheet, year, month) {
  try {
    // 日本時間で月の1日を設定するため、文字列で直接設定
    var dateString = year + '/' + month + '/1';
    sheet.getRange('B1').setValue(dateString);
    sheet.getRange('B1').setNumberFormat('yyyy/MM/dd');
    
    // B1のスタイルを設定（目立たないように）
    sheet.getRange('B1').setBackground('#f0f0f0');
    sheet.getRange('B1').setFontColor('#666666');
    sheet.getRange('B1').setFontSize(9);
    
    console.log('基準日付設定: ' + dateString);
    
  } catch (error) {
    console.error('基準日付設定エラー:', error);
    throw error;
  }
}

/**
 * 日付関数を設定（C1から）
 * @param {Sheet} sheet - 対象シート
 * @param {number} year - 年
 * @param {number} month - 月
 */
function setDateFormulas(sheet, year, month) {
  try {
    var daysInMonth = new Date(year, month, 0).getDate();
    
    // 既存の日付列をクリア
    var maxCols = sheet.getMaxColumns();
    if (maxCols > 3) {
      sheet.getRange(1, 3, 1, maxCols - 2).clearContent();
    }
    
    // C1に最初の日付関数を設定
    sheet.getRange('C1').setFormula('=DATE(YEAR($B$1), MONTH($B$1), 1)');
    
    // D1以降は前日+1の関数を設定
    for (var day = 2; day <= daysInMonth; day++) {
      var col = 3 + (day - 1);
      var prevCol = col - 1;
      var formula = '=' + sheet.getRange(1, prevCol).getA1Notation() + '+1';
      sheet.getRange(1, col).setFormula(formula);
    }
    
    // 日付形式を設定
    var dateRange = sheet.getRange(1, 3, 1, daysInMonth);
    dateRange.setNumberFormat('M/d');
    
    console.log('日付関数設定完了: ' + daysInMonth + '日分');
    
  } catch (error) {
    console.error('日付関数設定エラー:', error);
    throw error;
  }
}

/**
 * テンプレートとなるシートを探す
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {number} year - 対象年
 * @param {number} month - 対象月
 * @returns {Sheet|null} テンプレートシート
 */
function findTemplateSheet(ss, year, month) {
  var sheets = ss.getSheets();
  var candidates = [];
  
  // マスターシートのパターン
  var pattern = /^(\d{2})(\d{2})月_売上$/;
  
  sheets.forEach(function(sheet) {
    var match = sheet.getName().match(pattern);
    if (match) {
      var sheetYear = 2000 + parseInt(match[1], 10);
      var sheetMonth = parseInt(match[2], 10);
      
      // 過去のシートを候補に追加
      if (sheetYear < year || (sheetYear === year && sheetMonth < month)) {
        candidates.push({
          sheet: sheet,
          year: sheetYear,
          month: sheetMonth,
          score: (year - sheetYear) * 12 + (month - sheetMonth)
        });
      }
    }
  });
  
  // 最も近い過去のシートを選択
  if (candidates.length > 0) {
    candidates.sort(function(a, b) {
      return a.score - b.score;
    });
    return candidates[0].sheet;
  }
  
  return null;
}

/**
 * 日付ヘッダーを更新（C1から開始）
 * @param {Sheet} sheet - 対象シート
 * @param {number} year - 年
 * @param {number} month - 月
 */
function updateDateHeaders(sheet, year, month) {
  try {
    // 月の日数を取得
    var daysInMonth = new Date(year, month, 0).getDate();
    
    // 既存の日付列をクリア（D列以降の可能性もあるため広範囲をクリア）
    var maxCols = sheet.getMaxColumns();
    if (maxCols > 3) {
      sheet.getRange(1, 3, 1, maxCols - 2).clearContent();
    }
    
    // C1から日付を設定
    for (var day = 1; day <= daysInMonth; day++) {
      var col = 3 + (day - 1); // C列は3列目から開始
      // 時刻を含まない日付オブジェクトを作成
      var date = new Date(year, month - 1, day, 0, 0, 0, 0);
      
      // 日付値として設定（時刻を含まない）
      sheet.getRange(1, col).setValue(date);
      
      // 日付形式を設定（M/d形式）
      sheet.getRange(1, col).setNumberFormat('M/d');
    }
    
    console.log('日付ヘッダー設定完了: ' + year + '年' + month + '月（' + daysInMonth + '日間）');
    
  } catch (error) {
    console.error('日付ヘッダー更新エラー:', error);
    throw error;
  }
}

/**
 * 曜日行を設定（2行目）
 * @param {Sheet} sheet - 対象シート
 * @param {number} year - 年
 * @param {number} month - 月
 */
function setWeekdayFormulas(sheet, year, month) {
  try {
    var daysInMonth = new Date(year, month, 0).getDate();
    
    // C2から曜日の数式を設定
    for (var day = 1; day <= daysInMonth; day++) {
      var col = 3 + (day - 1); // C列は3列目から開始
      var formula = '=TEXT(' + sheet.getRange(1, col).getA1Notation() + ', "ddd")';
      sheet.getRange(2, col).setFormula(formula);
    }
    
    // 2行目の書式設定
    var weekdayRange = sheet.getRange(2, 3, 1, daysInMonth);
    weekdayRange.setHorizontalAlignment('center');
    weekdayRange.setFontSize(10);
    weekdayRange.setBackground('#f0f0f0');
    
    console.log('曜日行設定完了');
    
  } catch (error) {
    console.error('曜日行設定エラー:', error);
    throw error;
  }
}

/**
 * データと数式を設定（高速化版）
 * @param {Sheet} sheet - 対象シート
 * @param {number} year - 年
 * @param {number} month - 月
 */
function setupDataAndFormulas(sheet, year, month) {
  try {
    var lastRow = sheet.getLastRow();
    var daysInMonth = new Date(year, month, 0).getDate();
    var lastCol = 2 + daysInMonth; // B列(2) + 日数
    
    // 一括でA列とB列の値を取得（高速化）
    var dataRange = sheet.getRange(3, 1, lastRow - 2, 2).getValues();
    
    // まず全体のデータ範囲を一括クリア（高速化）
    if (lastRow >= 3) {
      sheet.getRange(3, 3, lastRow - 2, daysInMonth).clearContent();
    }
    
    // 各店舗の項目情報を事前に収集（高速化）
    var storeItemMap = {};
    for (var i = 0; i < dataRange.length; i++) {
      var storeName = dataRange[i][0];
      var itemName = dataRange[i][1];
      if (storeName && storeName.toString().trim() !== '') {
        if (!storeItemMap[storeName]) {
          storeItemMap[storeName] = {};
        }
        storeItemMap[storeName][itemName] = i + 3; // 行番号
      }
    }
    
    // 数式を配列にまとめて一括設定（高速化）
    var formulasToSet = [];
    
    for (var i = 0; i < dataRange.length; i++) {
      var row = i + 3;
      var storeCell = dataRange[i][0];
      var itemCell = dataRange[i][1];
      
      if (storeCell && storeCell.toString().trim() !== '') {
        var storeName = storeCell.toString();
        var itemName = itemCell.toString();
        
        // 項目に応じて処理を分岐
        if (itemName.indexOf('目標累計') > -1) {
          // 目標累計は当日目標の累計関数を設定
          var targetRow = storeItemMap[storeName]['当日目標'];
          if (targetRow) {
            formulasToSet.push({
              type: 'cumulative',
              row: row,
              sourceRow: targetRow,
              startCol: 3,
              endCol: lastCol
            });
          }
          
        } else if (itemName.indexOf('当月累計') > -1) {
          // 当月累計は売上の累計
          var salesRow = storeItemMap[storeName]['当日売上'];
          if (salesRow) {
            formulasToSet.push({
              type: 'cumulative',
              row: row,
              sourceRow: salesRow,
              startCol: 3,
              endCol: lastCol
            });
          }
          
        } else if (itemName.indexOf('当日人件費') > -1) {
          // 当日人件費はP/Aと社員の合計
          var paRow = storeItemMap[storeName]['P/A'];
          var shainRow = storeItemMap[storeName]['社員'];
          formulasToSet.push({
            type: 'sum',
            row: row,
            rows: [paRow, shainRow].filter(function(r) { return r; }),
            startCol: 3,
            endCol: lastCol
          });
          
        } else if (itemName.indexOf('累計人件費') > -1) {
          // 累計人件費
          var laborRow = storeItemMap[storeName]['当日人件費'];
          if (laborRow) {
            formulasToSet.push({
              type: 'cumulative',
              row: row,
              sourceRow: laborRow,
              startCol: 3,
              endCol: lastCol
            });
          }
          
        } else if (itemName.indexOf('累計仕入費') > -1) {
          // 累計仕入費
          var purchaseRow = storeItemMap[storeName]['当日仕入費'];
          if (purchaseRow) {
            formulasToSet.push({
              type: 'cumulative',
              row: row,
              sourceRow: purchaseRow,
              startCol: 3,
              endCol: lastCol
            });
          }
        }
      }
    }
    
    // 収集した数式を一括で設定（高速化）
    setBatchFormulas(sheet, formulasToSet);
    
    console.log('データと数式の設定完了');
    
  } catch (error) {
    console.error('データ設定エラー:', error);
    throw error;
  }
}

/**
 * 数式を一括で設定（高速化）
 * @param {Sheet} sheet - 対象シート
 * @param {Array} formulasToSet - 設定する数式の配列
 */
function setBatchFormulas(sheet, formulasToSet) {
  // 各行の数式を準備
  var rowFormulas = {};
  
  formulasToSet.forEach(function(item) {
    if (!rowFormulas[item.row]) {
      rowFormulas[item.row] = new Array(item.endCol - item.startCol + 1);
    }
    
    if (item.type === 'cumulative') {
      // 累計の数式
      for (var col = item.startCol; col <= item.endCol; col++) {
        var idx = col - item.startCol;
        if (col === item.startCol) {
          rowFormulas[item.row][idx] = '=' + getA1Notation(item.sourceRow, col);
        } else {
          rowFormulas[item.row][idx] = '=' + getA1Notation(item.row, col - 1) + '+' + getA1Notation(item.sourceRow, col);
        }
      }
    } else if (item.type === 'sum') {
      // 合計の数式
      for (var col = item.startCol; col <= item.endCol; col++) {
        var idx = col - item.startCol;
        var sumParts = item.rows.map(function(r) {
          return getA1Notation(r, col);
        });
        if (sumParts.length > 0) {
          rowFormulas[item.row][idx] = '=' + sumParts.join('+');
        }
      }
    }
  });
  
  // 行ごとに数式を一括設定
  for (var row in rowFormulas) {
    var formulas = rowFormulas[row];
    var range = sheet.getRange(parseInt(row), 3, 1, formulas.length);
    range.setFormulas([formulas]);
  }
}

/**
 * A1記法を取得
 * @param {number} row - 行番号
 * @param {number} col - 列番号
 * @returns {string} A1記法
 */
function getA1Notation(row, col) {
  var colLetter = '';
  var temp = col - 1;
  
  while (temp >= 0) {
    colLetter = String.fromCharCode((temp % 26) + 65) + colLetter;
    temp = Math.floor(temp / 26) - 1;
  }
  
  return colLetter + row;
}

/**
 * 累計関数を設定（目標用 - 当日目標の累計）
 * @param {Sheet} sheet - 対象シート
 * @param {number} row - 行番号（累計目標の行）
 * @param {string} storeName - 店舗名
 * @param {number} startCol - 開始列
 * @param {number} endCol - 終了列
 */
function setCumulativeFormulas(sheet, row, storeName, startCol, endCol) {
  // 同じ店舗の当日目標行を探す
  var targetRow = findItemRow(sheet, storeName, '当日目標');
  
  if (targetRow > 0) {
    // C列（最初の日）- 当日目標を参照
    sheet.getRange(row, startCol).setFormula('=' + sheet.getRange(targetRow, startCol).getA1Notation());
    
    // D列以降（累計）
    for (var col = startCol + 1; col <= endCol; col++) {
      var formula = '=' + sheet.getRange(row, col - 1).getA1Notation() + 
                    '+' + sheet.getRange(targetRow, col).getA1Notation();
      sheet.getRange(row, col).setFormula(formula);
    }
  }
}

/**
 * 売上の累計関数を設定
 * @param {Sheet} sheet - 対象シート
 * @param {number} row - 行番号
 * @param {string} storeName - 店舗名
 * @param {number} startCol - 開始列
 * @param {number} endCol - 終了列
 */
function setCumulativeFormulasForSales(sheet, row, storeName, startCol, endCol) {
  // 同じ店舗の当日売上行を探す
  var salesRow = findItemRow(sheet, storeName, '当日売上');
  
  if (salesRow > 0) {
    // C列（最初の日）
    sheet.getRange(row, startCol).setFormula('=' + sheet.getRange(salesRow, startCol).getA1Notation());
    
    // D列以降（累計）
    for (var col = startCol + 1; col <= endCol; col++) {
      var formula = '=' + sheet.getRange(row, col - 1).getA1Notation() + 
                    '+' + sheet.getRange(salesRow, col).getA1Notation();
      sheet.getRange(row, col).setFormula(formula);
    }
  }
}

/**
 * 人件費の累計関数を設定
 * @param {Sheet} sheet - 対象シート
 * @param {number} row - 行番号
 * @param {string} storeName - 店舗名
 * @param {number} startCol - 開始列
 * @param {number} endCol - 終了列
 */
function setCumulativeFormulasForLaborCost(sheet, row, storeName, startCol, endCol) {
  // 同じ店舗の当日人件費行を探す
  var laborCostRow = findItemRow(sheet, storeName, '当日人件費');
  
  if (laborCostRow > 0) {
    // C列（最初の日）
    sheet.getRange(row, startCol).setFormula('=' + sheet.getRange(laborCostRow, startCol).getA1Notation());
    
    // D列以降（累計）
    for (var col = startCol + 1; col <= endCol; col++) {
      var formula = '=' + sheet.getRange(row, col - 1).getA1Notation() + 
                    '+' + sheet.getRange(laborCostRow, col).getA1Notation();
      sheet.getRange(row, col).setFormula(formula);
    }
  }
}

/**
 * 仕入費の累計関数を設定
 * @param {Sheet} sheet - 対象シート
 * @param {number} row - 行番号
 * @param {string} storeName - 店舗名
 * @param {number} startCol - 開始列
 * @param {number} endCol - 終了列
 */
function setCumulativeFormulasForPurchase(sheet, row, storeName, startCol, endCol) {
  // 同じ店舗の当日仕入費行を探す
  var purchaseRow = findItemRow(sheet, storeName, '当日仕入費');
  
  if (purchaseRow > 0) {
    // C列（最初の日）
    sheet.getRange(row, startCol).setFormula('=' + sheet.getRange(purchaseRow, startCol).getA1Notation());
    
    // D列以降（累計）
    for (var col = startCol + 1; col <= endCol; col++) {
      var formula = '=' + sheet.getRange(row, col - 1).getA1Notation() + 
                    '+' + sheet.getRange(purchaseRow, col).getA1Notation();
      sheet.getRange(row, col).setFormula(formula);
    }
  }
}

/**
 * 人件費の合計関数を設定（P/Aと社員の合計）
 * @param {Sheet} sheet - 対象シート
 * @param {number} row - 行番号
 * @param {string} storeName - 店舗名
 * @param {number} startCol - 開始列
 * @param {number} endCol - 終了列
 */
function setLaborCostFormulas(sheet, row, storeName, startCol, endCol) {
  // 同じ店舗のP/Aと社員の行を探す
  var paRow = findItemRow(sheet, storeName, 'P/A');
  var shainRow = findItemRow(sheet, storeName, '社員');
  
  // 各日の合計関数を設定
  for (var col = startCol; col <= endCol; col++) {
    var formulas = [];
    if (paRow > 0) {
      formulas.push(sheet.getRange(paRow, col).getA1Notation());
    }
    if (shainRow > 0) {
      formulas.push(sheet.getRange(shainRow, col).getA1Notation());
    }
    
    if (formulas.length > 0) {
      var formula = '=' + formulas.join('+');
      sheet.getRange(row, col).setFormula(formula);
    }
  }
}

/**
 * 特定の店舗・項目の行を探す
 * @param {Sheet} sheet - 対象シート
 * @param {string} storeName - 店舗名
 * @param {string} itemName - 項目名
 * @returns {number} 行番号（見つからない場合は-1）
 */
function findItemRow(sheet, storeName, itemName) {
  var lastRow = sheet.getLastRow();
  var aValues = sheet.getRange('A3:A' + lastRow).getValues();
  var bValues = sheet.getRange('B3:B' + lastRow).getValues();
  
  for (var i = 0; i < aValues.length; i++) {
    if (aValues[i][0] === storeName && bValues[i][0].toString().indexOf(itemName) > -1) {
      return i + 3;
    }
  }
  return -1;
}

/**
 * 空のマスターシートを作成
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {string} sheetName - シート名
 * @param {number} year - 年
 * @param {number} month - 月
 */
function createEmptyMasterSheet(ss, sheetName, year, month) {
  var newSheet = ss.insertSheet(sheetName);
  
  // シートを先頭に移動
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(1);
  
  // 基本的な列ヘッダーを設定
  newSheet.getRange('A1').setValue('店舗');
  newSheet.getRange('B1').setValue('項目');
  
  // B1に基準日付を設定
  setBaseDateCell(newSheet, year, month);
  
  // 日付関数を設定（C1から）
  setDateFormulas(newSheet, year, month);
  
  // 曜日行を設定（2行目）
  setWeekdayFormulas(newSheet, year, month);
  
  // ヘッダー行の書式設定
  var daysInMonth = new Date(year, month, 0).getDate();
  var headerRange = newSheet.getRange(1, 1, 1, 2 + daysInMonth);
  headerRange.setBackground('#4a4a4a');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // B1だけは基準日付なので別スタイル
  newSheet.getRange('B1').setBackground('#f0f0f0');
  newSheet.getRange('B1').setFontColor('#666666');
  
  // A1を再設定（項目ヘッダーのスタイルを維持）
  newSheet.getRange('A1').setValue('店舗');
  newSheet.getRange('A1').setBackground('#4a4a4a');
  newSheet.getRange('A1').setFontColor('#ffffff');
  
  // 列幅の調整
  newSheet.setColumnWidth(1, 150); // 店舗列
  newSheet.setColumnWidth(2, 100); // 項目列
  
  // 固定行の設定
  newSheet.setFrozenRows(2); // 日付と曜日の2行を固定
  newSheet.setFrozenColumns(2);
}