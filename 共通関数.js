/**
 * 共通関数.gs
 * 複数の処理で使用される汎用的なヘルパー関数
 */

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

/**
 * 日付文字列の解析（「yyyy/mm/dd」形式または「月日」形式に対応）
 * @param {string} text - 日付を含むテキスト
 * @returns {Date|null} 解析された日付オブジェクト
 */
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

/**
 * 店舗名の順位を取得（storeOrderリストに基づく部分一致判定）
 * @param {string} text - 店舗名を含むテキスト
 * @param {Array} storeOrder - 店舗の表示順序配列
 * @returns {number} 順位（見つからない場合は最大値）
 */
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

/**
 * 日付を表示用にフォーマット
 * @param {string} dateStr - 日付文字列
 * @returns {string} フォーマット済み日付（例：7月29日）
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
  SpreadsheetApp.getUi().alert('⚠️ エラー', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 成功メッセージを表示
 * @param {string} message - 成功メッセージ
 */
function showSuccess(message) {
  console.log('成功:', message);
  SpreadsheetApp.getUi().alert('✅ 完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 情報メッセージを表示
 * @param {string} message - 情報メッセージ
 */
function showInfo(message) {
  console.log('情報:', message);
  SpreadsheetApp.getUi().alert('ℹ️ お知らせ', message, SpreadsheetApp.getUi().ButtonSet.OK);
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