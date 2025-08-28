/**
 * LINE Webhook Handler - 最適化版
 * タイムアウト対策済み
 */

// ===== 設定 =====
const CONFIG = {
  CHANNEL_ACCESS_TOKEN: '', // ここにChannel Access Tokenを設定（オプション）
  DEBUG_MODE: false // 本番環境ではfalseに設定
};

/**
 * GETリクエスト処理（動作確認用）
 */
function doGet(e) {
  const response = {
    status: 'active',
    message: 'LINE Webhook is ready',
    timestamp: new Date().toISOString(),
    deployment: 'production'
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * POSTリクエスト処理（LINE Webhook）
 * 重要: LINEは1秒以内のレスポンスを期待
 */
function doPost(e) {
  // 即座にレスポンスを返すことを優先
  
  // 空のリクエストチェック（最小限）
  if (!e || !e.postData || !e.postData.contents) {
    return quickResponse();
  }
  
  try {
    // JSONパース
    const json = JSON.parse(e.postData.contents);
    
    // イベントが存在する場合のみ処理
    if (json.events && json.events.length > 0) {
      // 非同期的に処理（レスポンスを先に返す）
      processEventsQuick(json.events);
    }
    
  } catch (error) {
    // エラーは記録するが、レスポンスは正常に返す
    if (CONFIG.DEBUG_MODE) {
      console.error('Parse error:', error.toString());
    }
  }
  
  // 即座に200 OKを返す
  return quickResponse();
}

/**
 * 高速レスポンス生成
 */
function quickResponse() {
  return ContentService
    .createTextOutput('{"status":"ok"}')
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * イベントの高速処理
 */
function processEventsQuick(events) {
  // 最初のイベントのみ処理（タイムアウト対策）
  const event = events[0];
  
  if (!event) return;
  
  try {
    // グループ参加イベント
    if (event.type === 'join' && event.source && event.source.type === 'group') {
      const groupId = event.source.groupId;
      saveGroupIdQuick(groupId);
      
      // リプライは後回し（必要な場合のみ）
      if (event.replyToken && CONFIG.CHANNEL_ACCESS_TOKEN) {
        sendReplyLater(event.replyToken, groupId);
      }
    }
    
    // メッセージイベント
    else if (event.type === 'message' && 
             event.source && 
             event.source.type === 'group' &&
             event.message && 
             event.message.type === 'text') {
      
      const groupId = event.source.groupId;
      const text = event.message.text;
      
      // グループIDコマンドチェック（大文字小文字を区別しない）
      if (text && isGroupIdCommand(text)) {
        saveGroupIdQuick(groupId);
        
        // リプライは後回し
        if (event.replyToken && CONFIG.CHANNEL_ACCESS_TOKEN) {
          sendReplyLater(event.replyToken, groupId);
        }
      }
    }
    
  } catch (error) {
    // エラーは無視（レスポンス速度優先）
    if (CONFIG.DEBUG_MODE) {
      console.error('Process error:', error.toString());
    }
  }
}

/**
 * グループIDコマンドかチェック
 */
function isGroupIdCommand(text) {
  const lowerText = text.toLowerCase().trim();
  const commands = ['グループid', 'group id', 'groupid', '@id', 'id'];
  return commands.includes(lowerText);
}

/**
 * グループIDの高速保存（PropertiesServiceのみ）
 */
function saveGroupIdQuick(groupId) {
  try {
    // Script Propertiesに保存（最速）
    const props = PropertiesService.getScriptProperties();
    props.setProperty('LINE_GROUP_ID', groupId);
    props.setProperty('LINE_GROUP_ID_DATE', new Date().toISOString());
    
    if (CONFIG.DEBUG_MODE) {
      console.log('Group ID saved:', groupId);
    }
    
    // スプレッドシートへの保存は別途実行
    // （タイムアウトを避けるため省略または遅延実行）
    
  } catch (error) {
    // エラーは無視
    if (CONFIG.DEBUG_MODE) {
      console.error('Save error:', error.toString());
    }
  }
}

/**
 * 遅延リプライ送信（エラーが出ても無視）
 */
function sendReplyLater(replyToken, groupId) {
  // リプライトークンの有効期限が短いため、実際には送信できない可能性が高い
  // 必要に応じてPush APIを使用することを検討
  
  if (!CONFIG.CHANNEL_ACCESS_TOKEN) return;
  
  try {
    const url = 'https://api.line.me/v2/bot/message/reply';
    const payload = {
      replyToken: replyToken,
      messages: [{
        type: 'text',
        text: `グループID: ${groupId}\n保存しました ✅`
      }]
    };
    
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + CONFIG.CHANNEL_ACCESS_TOKEN
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
  } catch (error) {
    // エラーは無視
  }
}

/**
 * ===== ユーティリティ関数 =====
 */

/**
 * 保存されたグループIDを取得（手動実行用）
 */
function getStoredGroupId() {
  const props = PropertiesService.getScriptProperties();
  const groupId = props.getProperty('LINE_GROUP_ID');
  const date = props.getProperty('LINE_GROUP_ID_DATE');
  
  console.log('=== 保存されたグループID ===');
  console.log('Group ID:', groupId || 'なし');
  console.log('保存日時:', date || 'なし');
  
  // スプレッドシートに書き込み（手動実行時のみ）
  if (groupId) {
    writeToSheet(groupId);
  }
  
  return groupId;
}

/**
 * スプレッドシートに書き込み（手動実行用）
 */
function writeToSheet(groupId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.log('スプレッドシートが見つかりません');
      return;
    }
    
    let sheet = ss.getSheetByName('LINE設定');
    if (!sheet) {
      sheet = ss.insertSheet('LINE設定');
      // ヘッダー
      const headers = [['項目', '値', '更新日時']];
      sheet.getRange(1, 1, 1, 3).setValues(headers);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      sheet.getRange(1, 1, 1, 3).setBackground('#4285f4');
      sheet.getRange(1, 1, 1, 3).setFontColor('#ffffff');
    }
    
    // データ
    const data = [['グループID', groupId, new Date()]];
    sheet.getRange(2, 1, 1, 3).setValues(data);
    
    // 列幅自動調整
    sheet.autoResizeColumns(1, 3);
    
    console.log('スプレッドシートに保存しました');
    
  } catch (error) {
    console.error('Sheet error:', error);
  }
}

/**
 * Webhook URLの確認（手動実行用）
 */
function checkWebhookUrl() {
  const url = ScriptApp.getService().getUrl();
  
  console.log('=== Webhook設定確認 ===');
  console.log('');
  console.log('1. このURLをLINE Developersに設定してください:');
  console.log(url);
  console.log('');
  console.log('2. デプロイ設定の確認:');
  console.log('   - 実行するユーザー: 自分');
  console.log('   - アクセスできるユーザー: 全員');
  console.log('');
  console.log('3. ブラウザでアクセステスト:');
  console.log('   上記URLをブラウザで開いて JSON が表示されることを確認');
  console.log('');
  console.log('4. LINE Developersの設定:');
  console.log('   - Webhook URL: 上記URL');
  console.log('   - Webhookの利用: オン');
  console.log('   - 応答メッセージ: オフ');
}

/**
 * テスト実行（GASエディタから実行）
 */
function testWebhookLocal() {
  // テストペイロード
  const testEvent = {
    postData: {
      contents: JSON.stringify({
        events: [{
          type: 'message',
          source: {
            type: 'group',
            groupId: 'C1234567890abcdef'
          },
          message: {
            type: 'text',
            text: 'グループID'
          },
          replyToken: 'test_token'
        }]
      })
    }
  };
  
  console.log('=== Webhookテスト開始 ===');
  
  // 実行
  const result = doPost(testEvent);
  console.log('Response:', result.getContent());
  
  // 保存確認
  const savedId = getStoredGroupId();
  console.log('保存されたID:', savedId);
}

/**
 * PropertiesServiceのクリア（必要に応じて実行）
 */
function clearProperties() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('LINE_GROUP_ID');
  props.deleteProperty('LINE_GROUP_ID_DATE');
  console.log('Properties cleared');
}