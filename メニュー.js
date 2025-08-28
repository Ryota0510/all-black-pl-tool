/**
 * メニュー.gs
 * カスタムメニューの作成と管理
 */

// シートオープン時にカスタムメニューを追加
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // メインメニュー
  ui.createMenu('📊 売上データ処理')
    .addItem('📋 LINE貼り付け → B列に展開', 'processLineMessage')
    .addItem('🚀 LINE貼り付け → 展開＋転記', 'processLineAndTransfer')
    .addSeparator()
    .addItem('📝 B列 → マスターシート転記', 'transferFromReportSheet')
    .addSeparator()
    .addItem('➕ 新しい月のシートを作成', 'showCreateMasterSheetDialog')
    .addItem('📊 今月のシートを作成', 'createCurrentMonthSheet')
    .addItem('📈 来月のシートを作成', 'createNextMonthSheet')
    .addSeparator()
    .addItem('📄 PDF作成 → LINE送信', 'sendSheetAsPDF')
    .addItem('⚙️ LINE API設定', 'showLineConfigDialog')
    .addItem('📱 グループID取得方法', 'showWebhookSetupGuide')
    .addSeparator()
    .addItem('❓ 使い方・ヘルプ', 'showHelp')
    .addToUi();
  
  // ボタンの作成（売上報告シート）
  createButtons();
}

/**
 * 売上報告シートにボタンを作成
 */
function createButtons() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var reportSheet = ss.getSheetByName('売上報告');
    
    if (reportSheet) {
      // 既存の図形を削除（重複防止）
      var drawings = reportSheet.getDrawings();
      drawings.forEach(function(drawing) {
        if (drawing.getContainerInfo() && 
            drawing.getOnAction() === 'processLineMessage') {
          drawing.remove();
        }
      });
      
      // ボタンを作成
      var button = reportSheet.insertDrawing(
        SpreadsheetApp.newDrawing()
          .setWidth(200)
          .setHeight(40)
          .setPosition(2, 4, 0, 0) // B4セルの位置
          .build()
      );
      
      // ボタンにスクリプトを割り当て
      button.setOnAction('processLineMessage');
      
      console.log('ボタンを作成しました');
    }
  } catch (error) {
    console.error('ボタン作成エラー:', error);
  }
}

/**
 * ヘルプを表示
 */
function showHelp() {
  var ui = SpreadsheetApp.getUi();
  var helpMessage = 
    '📊 売上データ処理システムの使い方\n\n' +
    '【基本的な使い方】\n' +
    '1. 売上報告シートのA1セルにLINEメッセージを貼り付け\n' +
    '2. 「LINE貼り付け → 展開＋転記」を実行\n' +
    '3. 自動的にB列に展開され、マスターシートへ転記されます\n\n' +
    '【メニューの説明】\n' +
    '📋 LINE貼り付け → B列に展開\n' +
    '　→ A1のLINEメッセージをB列に展開します（転記なし）\n\n' +
    '🚀 LINE貼り付け → 展開＋転記（一括実行）\n' +
    '　→ B列への展開とマスターシートへの転記を一括で実行\n\n' +
    '✅ B列 → マスターシート（確認あり）\n' +
    '　→ B列のデータを確認しながら転記します\n\n' +
    '⚡ B列 → マスターシート（一括転記）\n' +
    '　→ 確認なしで一括転記します\n\n' +
    '【注意事項】\n' +
    '• マスターシートは「2507月_売上」のような形式で命名\n' +
    '• 転記時は日付と店舗名でデータを照合します\n' +
    '• 不要なデータ（天気、達成率など）は自動的に除外されます';
  
  ui.alert('❓ ヘルプ', helpMessage, ui.ButtonSet.OK);
}

/**
 * LINEメッセージを展開して転記まで一括実行
 */
function processLineAndTransfer() {
  // まずLINEメッセージを処理
  if (processLineMessage(true)) { // 成功メッセージを抑制
    // その後転記処理
    Utilities.sleep(1000);
    transferFromReportSheet();
  }
}

/**
 * 売上報告シートから転記（確認なし）
 */
function transferFromReportSheet() {
  transferReportToMaster(true); // batchMode = true（確認なし）
}