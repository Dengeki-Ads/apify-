/**
 * api.gs — 外部からの Script Properties 更新 API
 */

/** 更新を許可するキーのホワイトリスト */
const ALLOWED_KEYS = ['SponsoredBy'];

/**
 * プロパティ更新リクエストを処理する。
 * doPost() から呼び出される。
 */
const handlePropertyUpdate = (payload) => {
  const authToken = PropertiesService.getScriptProperties().getProperty('AUTH_TOKEN');
  if (!authToken) {
    return jsonResponse({ status: 'error', message: 'AUTH_TOKEN is not configured' });
  }

  if (payload.token !== authToken) {
    Logger.log('[API] Unauthorized request.');
    return jsonResponse({ status: 'error', message: 'Unauthorized' });
  }

  const updated = [];
  for (const [key, value] of Object.entries(payload)) {
    if (key === 'token') continue;

    if (!ALLOWED_KEYS.includes(key)) {
      Logger.log(`[API WARN] Key "${key}" is not allowed. Skipped.`);
      continue;
    }

    PropertiesService.getScriptProperties().setProperty(key, value);
    updated.push(key);
  }

  Logger.log(`[API OK] Updated: ${updated.length > 0 ? updated.join(', ') : 'none'}`);
  return jsonResponse({ status: 'ok', updated: updated });
};

/**
 * JSON レスポンスを生成するヘルパー。
 */
const jsonResponse = (obj) => {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
};
