/**
 * webhook.gs — Apify Webhook受信 → データ書き込み
 */

/**
 * Apify Webhook POST受信エントリポイント。
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    // ルーティング: token フィールドがあればプロパティ更新API
    if (payload.token !== undefined) {
      return handlePropertyUpdate(payload);
    }

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

    // raw原本を別スプレッドシートに保存（加工前）
    let rawInfo = null;
    try {
      rawInfo = saveRawSpreadsheet(headers, rows, {
        runId,
        datasetId,
        fetchedAt,
      });
    } catch (rawErr) {
      Logger.log(`[RAW ERROR] ${rawErr.message}`);
    }

    // 月別シートに振り分けて書き込み
    const grouped = groupRowsByMonth(headers, rows, dataKeys);
    const updatedMonths = [];

    for (const [monthName, monthData] of Object.entries(grouped)) {
      overwriteMonthlySheet(monthName, headers, monthData.rows);
      applySheetTransformations(monthName);
      updatedMonths.push(monthName);
    }

    // 統合シート再生成
    let consolidatedStatus = 'skip';
    try {
      consolidatedStatus = rebuildConsolidatedSheet(updatedMonths);
    } catch (consErr) {
      Logger.log(`[CONSOLIDATED ERROR] ${consErr.message}`);
      consolidatedStatus = 'error';
    }

    // Looker Studio同期（統合シートから）
    syncToLookerStudio();

    updateLogRow(runId, {
      completed_at: new Date(),
      status: '完了',
      result_count: rows.length,
      target_month: updatedMonths.join(','),
      raw_file_id: rawInfo ? rawInfo.fileId : '',
      raw_file_url: rawInfo ? rawInfo.fileUrl : '',
      consolidated_status: consolidatedStatus,
    });

    Logger.log(`Webhook processed. RunID: ${runId}, Items: ${rows.length}, Months: ${updatedMonths.join(',')}`);
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
 * データ行を月別にグルーピングする。
 * uploadedAtFormatted列からYYYY-MMを抽出して振り分け。
 */
const groupRowsByMonth = (headers, rows, dataKeys) => {
  const dateColIndex = headers.indexOf('uploadedAtFormatted');
  const grouped = {};

  for (const row of rows) {
    let monthName = 'unknown';

    if (dateColIndex !== -1) {
      const dateVal = String(row[dateColIndex] || '');
      const match = dateVal.match(/^(\d{4}-\d{2})/);
      if (match) {
        monthName = match[1];
      }
    }

    if (!grouped[monthName]) {
      grouped[monthName] = { rows: [] };
    }
    grouped[monthName].rows.push(row);
  }

  return grouped;
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
