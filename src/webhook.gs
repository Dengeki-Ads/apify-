/**
 * webhook.gs — Apify Webhook受信 → データ書き込み
 */

/**
 * Apify Webhook POST受信エントリポイント。
 * シート操作は行わず、Apifyからのデータ取得とraw保存のみを担当。
 * その後のシート整形・書込・統合・Looker同期は processSheetData (別実行) に委譲する。
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    // ルーティング: token フィールドがあればプロパティ更新API
    if (payload.token !== undefined || payload.action !== undefined || payload.bootstrapToken !== undefined) {
      return handlePropertyUpdate(payload);
    }

    const runId = payload.resource.id;
    const runStatus = payload.resource.status;
    const datasetId = payload.resource.defaultDatasetId;

    if (runStatus !== 'SUCCEEDED') {
      Logger.log(`[DOPOST] Actor run not succeeded. RunID: ${runId}, Status: ${runStatus}`);
      return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    }

    const url = buildDatasetUrl(datasetId);
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const statusCode = response.getResponseCode();

    if (statusCode !== 200) {
      Logger.log(`[DOPOST] Dataset fetch failed (HTTP ${statusCode}): ${response.getContentText()}`);
      return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    }

    const items = JSON.parse(response.getContentText());

    if (!items || items.length === 0) {
      Logger.log(`[DOPOST] Dataset is empty. RunID: ${runId}`);
      return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
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

    // raw原本を別スプレッドシートに保存（メインスプレッドシートではなく Drive 上の別ファイル）
    // 後続の processSheetData はこのファイルから読み戻す
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

    if (!rawInfo) {
      Logger.log(`[DOPOST] Failed to save raw spreadsheet. RunID: ${runId}`);
      return ContentService.createTextOutput('ERROR').setMimeType(ContentService.MimeType.TEXT);
    }

    // 月別振り分け・整形・書込・統合・Looker同期は processSheetData に委譲
    enqueueJob({
      type: 'process_data',
      runId,
      datasetId,
      fetchedAt: fetchedAt.toISOString(),
      rawFileId: rawInfo.fileId,
      rawFileUrl: rawInfo.fileUrl,
      resultCount: rows.length,
    });
    scheduleProcessSheetData();

    Logger.log(`Webhook accepted. RunID: ${runId}, Items: ${rows.length}, Raw: ${rawInfo.fileId}`);
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    Logger.log(`doPost error: ${err.message}\n${err.stack}`);
    return ContentService.createTextOutput('ERROR').setMimeType(ContentService.MimeType.TEXT);
  }
}

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
