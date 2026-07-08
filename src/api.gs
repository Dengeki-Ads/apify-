/**
 * api.gs — 外部からの Script Properties 更新 API
 */

/**
 * 更新を許可するキーのホワイトリスト。
 * 末尾 '*' は前方一致パターン（例 'VIEWS_GOAL_*' は VIEWS_GOAL_ で始まる任意のキーを許可）。
 * それ以外は完全一致。
 */
const ALLOWED_KEYS = ['SponsoredBy', 'UploadedBy', 'VIEWS_GOAL_*'];

/**
 * プロパティ更新リクエストを処理する。
 * doPost() から呼び出される。
 */
const handlePropertyUpdate = (payload) => {
  const props = PropertiesService.getScriptProperties();
  const authToken =
    props.getProperty('HERMES_PROP_API_TOKEN') ||
    props.getProperty('AUTH_TOKEN');
  if (!authToken) {
    return jsonResponse({ status: 'error', message: 'HERMES_PROP_API_TOKEN or AUTH_TOKEN is not configured' });
  }

  if (payload.token !== authToken) {
    Logger.log('[API] Unauthorized request.');
    return jsonResponse({ status: 'error', message: 'Unauthorized' });
  }

  const updated = [];

  for (const [key, value] of Object.entries(payload)) {
    if (key === 'token') continue;

    if (!ALLOWED_KEYS.some(pattern => pattern.endsWith('*') ? key.startsWith(pattern.slice(0, -1)) : key === pattern)) {
      Logger.log(`[API WARN] Key "${key}" is not allowed. Skipped.`);
      continue;
    }

    props.setProperty(key, value);
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
