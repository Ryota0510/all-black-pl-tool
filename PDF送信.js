/**
 * PDF送信.gs
 * PDFの作成とLINEグループへの送信
 */

/**
 * 現在のシートをPDF化してLINEに送信
 */
function sendSheetAsPDF() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();
    
    // FLシートかチェック（例：2508月FL）
    if (!sheetName.match(/^\d{4}月FL$/)) {
      SpreadsheetApp.getUi().alert(
        '⚠️ エラー',
        '「○○○○月FL」形式のシートを開いてください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // LINE設定を取得
    var config = getLineConfig();
    if (!config.channelAccessToken || !config.groupId) {
      showLineConfigDialog();
      return;
    }
    
    // PDFを作成
    var pdf = createPDF(sheet);
    
    // LINEに送信
    sendPDFToLine(pdf, sheetName + '.pdf', config);
    
  } catch (error) {
    console.error('PDF送信エラー:', error);
    SpreadsheetApp.getUi().alert(
      '⚠️ エラー',
      'PDF送信中にエラーが発生しました。\n' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * シートからPDFを作成
 * @param {Sheet} sheet - 対象シート
 * @returns {Blob} PDFファイル
 */
function createPDF(sheet) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = spreadsheet.getId();
  var sheetId = sheet.getSheetId();
  
  // PDF作成のパラメータ
  var params = {
    'size': 'A4',                    // 用紙サイズ
    'portrait': true,                // 縦向き
    'fitw': true,                    // 幅に合わせる
    'sheetnames': false,             // シート名を表示しない
    'printtitle': false,             // タイトルを表示しない
    'pagenumbers': false,            // ページ番号を表示しない
    'gridlines': false,              // グリッドラインを表示しない
    'fzr': false,                    // 固定行を繰り返さない
    'fzc': false,                    // 固定列を繰り返さない
    'r1': 0,                         // 開始行（0ベース）
    'c1': 0,                         // 開始列（0ベース）
    'r2': 51,                        // 終了行（0ベース、A1~R52なので51）
    'c2': 17                         // 終了列（0ベース、R列は17）
  };
  
  // URLパラメータを作成
  var paramString = Object.keys(params).map(function(key) {
    return key + '=' + params[key];
  }).join('&');
  
  // PDF作成URL
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + 
            '/export?format=pdf&gid=' + sheetId + '&' + paramString;
  
  // PDFを取得
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    'headers': {
      'Authorization': 'Bearer ' + token
    }
  });
  
  return response.getBlob();
}

/**
 * PDFをLINEに送信
 * @param {Blob} pdfBlob - PDFファイル
 * @param {string} fileName - ファイル名
 * @param {Object} config - LINE設定
 */
function sendPDFToLine(pdfBlob, fileName, config) {
  try {
    console.log('LINE送信開始: ' + fileName);
    
    // まずメッセージを送信
    var messageUrl = 'https://api.line.me/v2/bot/message/push';
    var messagePayload = {
      'to': config.groupId,
      'messages': [{
        'type': 'text',
        'text': '📊 ' + fileName + ' を送信します。'
      }]
    };
    
    var messageOptions = {
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + config.channelAccessToken
      },
      'payload': JSON.stringify(messagePayload)
    };
    
    UrlFetchApp.fetch(messageUrl, messageOptions);
    
    // PDFファイルを送信
    // 注: LINE Messaging APIではファイルの直接送信はサポートされていないため、
    // Google Driveに一時保存してリンクを送信する方式を使用
    
    var file = DriveApp.createFile(pdfBlob);
    file.setName(fileName);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var fileUrl = file.getUrl();
    
    // ファイルリンクを送信
    var linkPayload = {
      'to': config.groupId,
      'messages': [{
        'type': 'text',
        'text': '📎 PDFファイル: ' + fileUrl
      }]
    };
    
    var linkOptions = {
      'method': 'post',
      'headers': {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + config.channelAccessToken
      },
      'payload': JSON.stringify(linkPayload)
    };
    
    UrlFetchApp.fetch(messageUrl, linkOptions);
    
    // 一定時間後にファイルを削除（オプション）
    // Utilities.sleep(60000); // 1分待機
    // file.setTrashed(true);
    
    SpreadsheetApp.getUi().alert(
      '✅ 送信完了',
      'PDFをLINEグループに送信しました。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('LINE送信エラー:', error);
    throw new Error('LINE送信に失敗しました: ' + error.message);
  }
}

/**
 * LINE設定を取得
 * @returns {Object} LINE設定
 */
function getLineConfig() {
  var documentProperties = PropertiesService.getDocumentProperties();
  return {
    channelAccessToken: documentProperties.getProperty('LINE_CHANNEL_ACCESS_TOKEN') || '',
    groupId: documentProperties.getProperty('LINE_GROUP_ID') || ''
  };
}

/**
 * LINE設定を保存
 * @param {string} channelAccessToken - チャンネルアクセストークン
 * @param {string} groupId - グループID
 */
function saveLineConfig(channelAccessToken, groupId) {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('LINE_CHANNEL_ACCESS_TOKEN', channelAccessToken);
  documentProperties.setProperty('LINE_GROUP_ID', groupId);
}

/**
 * LINE設定ダイアログを表示
 */
function showLineConfigDialog() {
  var html = HtmlService.createHtmlOutputFromFile('LINE設定ダイアログ')
      .setWidth(500)
      .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '⚙️ LINE API設定');
}

/**
 * 現在の設定を取得（ダイアログ用）
 */
function getCurrentConfig() {
  var config = getLineConfig();
  // セキュリティのため、トークンの一部をマスク
  if (config.channelAccessToken) {
    config.maskedToken = config.channelAccessToken.substring(0, 10) + '...' + 
                        config.channelAccessToken.substring(config.channelAccessToken.length - 10);
  }
  return config;
}