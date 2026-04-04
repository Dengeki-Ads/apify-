/**
 * notify.gs — sponsored_by 別 views 集計・通知
 */

/**
 * メイン関数。data シートから sponsored_by 別の views 合計を集計し通知する。
 */
function notifyViewsSummary() {
  try {
    const sheet = getSheet('data');
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow <= 1) {
      Logger.log('[NOTIFY SKIP] data sheet has no data rows.');
      return;
    }

    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0];

    // ヘッダーから列インデックスを動的に取得
    const sponsoredByIdx = headers.indexOf('sponsored_by');
    const viewsIdx = headers.indexOf('views');

    if (sponsoredByIdx === -1 || viewsIdx === -1) {
      const missing = [];
      if (sponsoredByIdx === -1) missing.push('sponsored_by');
      if (viewsIdx === -1) missing.push('views');
      throw new Error(`Required column(s) not found: ${missing.join(', ')}`);
    }

    // sponsored_by 別に views を集計
    const summary = {};
    for (let i = 1; i < data.length; i++) {
      const sponsor = data[i][sponsoredByIdx];
      const views = Number(data[i][viewsIdx]) || 0;

      if (!sponsor) continue;

      if (!summary[sponsor]) {
        summary[sponsor] = 0;
      }
      summary[sponsor] += views;
    }

    if (Object.keys(summary).length === 0) {
      Logger.log('[NOTIFY SKIP] No sponsored_by data found.');
      return;
    }

    const sheetName = sheet.getName();
    sendNotification(summary, sheetName);

    Logger.log(`[NOTIFY OK] Summary sent. Groups: ${Object.keys(summary).length}`);

  } catch (err) {
    Logger.log(`[NOTIFY ERROR] ${err.message}`);
    try {
      sendErrorNotification(
        'Views集計レポート: エラー',
        `notifyViewsSummary() でエラーが発生しました。\n\n${err.message}`
      );
    } catch (mailErr) {
      Logger.log(`Failed to send error notification: ${mailErr.message}`);
    }
  }
}
