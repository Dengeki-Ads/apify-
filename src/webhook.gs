/**
 * webhook.gs — Apify Webhook受信 → データ書き込み
 */

/**
 * Apify Webhook POST受信エントリポイント。
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const runId = payload.resource.id;
    const runStatus = payload.resource.status;
    const datasetId = payload.resource.defaultDatasetId;

    if (runStatus !== 'SUCCEEDED') {
      updateLogRow(runId, {
        completed_at: new Date(),
        status: '失敗',
        error_detail: `Actor run status: ${runStatus}`,
      });
      sendErrorNotification(
        'TikTok Scraper: Actor実行失敗',
        `RunID: ${runId}\nStatus: ${runStatus}`
      );
      return ContentService.createTextOutput('ERROR').setMimeType(ContentService.MimeType.TEXT);
    }

    const url = buildDatasetUrl(datasetId);
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const statusCode = response.getResponseCode();

    if (statusCode !== 200) {
      const errMsg = `Dataset fetch failed (HTTP ${statusCode}): ${response.getContentText()}`;
      updateLogRow(runId, { completed_at: new Date(), status: '失敗', error_detail: errMsg });
      sendErrorNotification('TikTok Scraper: データ取得失敗', errMsg);
      return ContentService.createTextOutput('ERROR').setMimeType(ContentService.MimeType.TEXT);
    }

    const items = JSON.parse(response.getContentText());

    if (!items || items.length === 0) {
      updateLogRow(runId, { completed_at: new Date(), status: '失敗', error_detail: 'Dataset is empty' });
      sendErrorNotification('TikTok Scraper: データ取得失敗', `RunID: ${runId}\nDataset is empty.`);
      return ContentService.createTextOutput('ERROR').setMimeType(ContentService.MimeType.TEXT);
    }

    const fetchedAt = new Date();
    const flatItems = items.map((item) => flattenObject(item));

    // 全アイテムからユニークキーを収集
    const keySet = new Set();
    const dataKeys = [];
    for (const flat of flatItems) {
      for (const key of Object.keys(flat)) {
        if (!keySet.has(key)) {
          keySet.add(key);
          dataKeys.push(key);
        }
      }
    }

    const headers = ['fetched_at', 'run_id', ...dataKeys];

    const rows = flatItems.map((flat) => {
      const row = [fetchedAt, runId];
      for (const key of dataKeys) {
        row.push(flat[key] !== undefined ? flat[key] : '');
      }
      return row;
    });

    checkAndArchive();
    overwriteDataSheet(headers, rows);
    filterColumns();
    addHashtagFormulaColumns();
    syncToLookerStudio();

    updateLogRow(runId, {
      completed_at: new Date(),
      status: '完了',
      result_count: rows.length,
    });

    Logger.log(`Webhook processed. RunID: ${runId}, Items: ${rows.length}`);
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    Logger.log(`doPost error: ${err.message}\n${err.stack}`);

    try {
      sendErrorNotification(
        'TikTok Scraper: Webhook処理エラー',
        `Error: ${err.message}\n\nStack: ${err.stack}`
      );
    } catch (mailErr) {
      Logger.log(`Failed to send error notification: ${mailErr.message}`);
    }

    return ContentService.createTextOutput('ERROR').setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * ヘルスチェック用GETエンドポイント。
 */
function doGet(e) {
  return ContentService.createTextOutput('TikTok Scraper Webhook is active.')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * エラー通知メールを送信する。
 */
const sendErrorNotification = (subject, body) => {
  const email = getConfig('NOTIFY_EMAIL');
  GmailApp.sendEmail(email, subject, body);
  Logger.log(`Error notification sent to ${email}`);
};

/**
 * ネストされたオブジェクトを「.」区切りでフラット化する。
 * 配列はJSON文字列化して1セルに格納。
 */
const flattenObject = (obj, prefix = '') => {
  const result = {};

  for (const [key, value] of Object.entries(obj)) {
    const fullKey = prefix ? `${prefix}.${key}` : key;

    if (value === null || value === undefined) {
      result[fullKey] = '';
    } else if (Array.isArray(value)) {
      result[fullKey] = JSON.stringify(value);
    } else if (typeof value === 'object') {
      Object.assign(result, flattenObject(value, fullKey));
    } else {
      result[fullKey] = value;
    }
  }

  return result;
};
