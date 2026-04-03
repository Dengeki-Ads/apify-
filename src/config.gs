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
 * COLUMNS_TO_KEEP をパースし、常に保持する列を付与して返す。
 */
const getColumnsToKeep = () => {
  const raw = getConfig('COLUMNS_TO_KEEP');
  return raw.split(',').map((s) => s.trim()).filter((s) => s);
};
