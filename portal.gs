/**
 * portal.gs - Webアプリポータル エントリーポイント
 * CDmedical 健診結果管理システム
 *
 * デプロイ方法:
 * 1. Apps Script エディタで「デプロイ」→「新しいデプロイ」
 * 2. 種類: ウェブアプリ
 * 3. 実行ユーザー: 自分
 * 4. アクセスできるユーザー: 組織内全員（または特定ユーザー）
 */

/**
 * ポータル用エントリーポイント（別デプロイ用）
 * ※ メインのdoGetはUI.gsで定義
 * @param {Object} e - リクエストパラメータ
 * @returns {HtmlOutput} HTMLページ
 */
function doGetPortal(e) {
  try {
    const template = HtmlService.createTemplateFromFile('templates/portal');

    // テンプレートに初期データを渡す
    template.initialData = JSON.stringify({
      user: Session.getActiveUser().getEmail(),
      timestamp: new Date().toISOString(),
      version: '1.0.0'
    });

    const htmlOutput = template.evaluate()
      .setTitle('CDmedical 健診システム')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

    return htmlOutput;
  } catch (error) {
    console.error('Portal initialization error:', error);
    return HtmlService.createHtmlOutput(
      '<h1>エラーが発生しました</h1><p>' + error.message + '</p>'
    );
  }
}

/**
 * HTMLファイルをインクルードするヘルパー関数
 * @param {string} filename - インクルードするファイル名
 * @returns {string} HTMLコンテンツ
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * スプレッドシートIDを取得
 * @returns {string} スプレッドシートID
 */
function getSpreadsheetId() {
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}

/**
 * 現在のユーザー情報を取得
 * @returns {Object} ユーザー情報
 */
function getCurrentUser() {
  return {
    email: Session.getActiveUser().getEmail(),
    timezone: Session.getScriptTimeZone()
  };
}
