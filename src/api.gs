/**
 * api.gs — 外部からの Script Properties 更新 API
 */

/** 更新を許可するキーのホワイトリスト */
const ALLOWED_KEYS = ['SponsoredBy', 'UploadedBy', 'START_URLS'];

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
  const props = PropertiesService.getScriptProperties();

  for (const [key, value] of Object.entries(payload)) {
    if (key === 'token') continue;

    if (!ALLOWED_KEYS.includes(key)) {
      Logger.log(`[API WARN] Key "${key}" is not allowed. Skipped.`);
      continue;
    }

    props.setProperty(key, value);
    updated.push(key);
  }

  // START_URLS の追加・削除（差分更新）
  if (payload.add_urls || payload.remove_urls) {
    const current = (props.getProperty('START_URLS') || '').split(',').map((s) => s.trim()).filter((s) => s);
    let urls = current;

    if (payload.add_urls) {
      const toAdd = payload.add_urls.split(',').map((s) => s.trim()).filter((s) => s);
      toAdd.forEach((u) => { if (!urls.includes(u)) urls.push(u); });
      updated.push('add_urls(' + toAdd.length + ')');
    }

    if (payload.remove_urls) {
      const toRemove = payload.remove_urls.split(',').map((s) => s.trim()).filter((s) => s);
      urls = urls.filter((u) => !toRemove.includes(u));
      updated.push('remove_urls(' + toRemove.length + ')');
    }

    props.setProperty('START_URLS', urls.join(','));
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
