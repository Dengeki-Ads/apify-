/**
 * config.gs — スクリプトプロパティ管理
 */

let _spreadsheet = null;

/**
 * スクリプトプロパティからキーを取得する。未設定時はエラー。
 */
const getConfig = (key) => {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error(`Property "${key}" is not set. Set it in Script Properties.`);
  }
  return value;
};

/**
 * スプレッドシートをIDで取得する（実行中キャッシュ）。
 */
const getSpreadsheet = () => {
  if (!_spreadsheet) {
    _spreadsheet = SpreadsheetApp.openById(getConfig('SPREADSHEET_ID'));
  }
  return _spreadsheet;
};

/**
 * Apify Task起動URLを組み立てる。
 */
const buildTaskRunUrl = () => {
  const apiKey = getConfig('APIFY_API_KEY');
  const taskId = getConfig('APIFY_TASK_ID');
  return `https://api.apify.com/v2/actor-tasks/${taskId}/runs?token=${apiKey}`;
};

/**
 * Apify Dataset取得URLを組み立てる。
 */
const buildDatasetUrl = (datasetId) => {
  const apiKey = getConfig('APIFY_API_KEY');
  return `https://api.apify.com/v2/datasets/${datasetId}/items?token=${apiKey}`;
};

/**
 * Apify Task APIから現在のタスク入力設定を取得する。
 */
const fetchTaskInput = () => {
  const apiKey = getConfig('APIFY_API_KEY');
  const taskId = getConfig('APIFY_TASK_ID');
  const url = `https://api.apify.com/v2/actor-tasks/${taskId}?token=${apiKey}`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error(`Failed to fetch task input (HTTP ${response.getResponseCode()})`);
  }
  return JSON.parse(response.getContentText()).data.input;
};

/**
 * Apifyタスクの期間文字列を返す（"since_until" 形式）。
 */
const fetchTaskPeriod = () => {
  const input = fetchTaskInput();
  const since = input.since || '';
  const until = input.until || '';
  return `${since}_${until}`;
};

/**
 * Apify Taskの現在の取得期間をログに表示する。手動実行用。
 */
function checkPeriod() {
  // Task保存済み設定
  const input = fetchTaskInput();
  Logger.log('[Task設定] since: ' + input.since + ', until: ' + input.until);
  Logger.log('[Task設定] maxItems: ' + input.maxItems);
  Logger.log('[Task設定] startUrls数: ' + (input.startUrls ? input.startUrls.length : 'なし'));
  Logger.log('[Task設定] 全体: ' + JSON.stringify(input, null, 2));

  // 直近のRunで実際に使われた入力
  const apiKey = getConfig('APIFY_API_KEY');
  const taskId = getConfig('APIFY_TASK_ID');
  const url = `https://api.apify.com/v2/actor-tasks/${taskId}/runs/last?token=${apiKey}`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() === 200) {
    const run = JSON.parse(response.getContentText()).data;
    Logger.log('[最終Run] ID: ' + run.id + ', status: ' + run.status);
    Logger.log('[最終Run] startedAt: ' + run.startedAt);
    Logger.log('[最終Run] datasetId: ' + run.defaultDatasetId);
    Logger.log('[最終Run] itemCount: ' + (run.stats && run.stats.inputItemCount));
    // Runの入力を取得
    const inputUrl = `https://api.apify.com/v2/key-value-stores/${run.defaultKeyValueStoreId}/records/INPUT?token=${apiKey}`;
    const inputRes = UrlFetchApp.fetch(inputUrl, { muteHttpExceptions: true });
    if (inputRes.getResponseCode() === 200) {
      const runInput = JSON.parse(inputRes.getContentText());
      Logger.log('[最終Run入力] since: ' + runInput.since + ', until: ' + runInput.until);
      Logger.log('[最終Run入力] 全体: ' + JSON.stringify(runInput, null, 2));
    }
  }
}

/**
 * COLUMNS_TO_KEEP をパースし、常に保持する列を付与して返す。
 */
const getColumnsToKeep = () => {
  const raw = getConfig('COLUMNS_TO_KEEP');
  return raw.split(',').map((s) => s.trim()).filter((s) => s);
};

/**
 * 統合シート名を取得する。未設定時は「統合」。
 */
const getConsolidatedSheetName = () => {
  const name = PropertiesService.getScriptProperties().getProperty('CONSOLIDATED_SHEET_NAME');
  return name || '統合';
};

/**
 * スクリプトプロパティからキーを取得する。未設定時はnullを返す（エラーなし）。
 */
const getOptionalConfig = (key) => {
  return PropertiesService.getScriptProperties().getProperty(key);
};
