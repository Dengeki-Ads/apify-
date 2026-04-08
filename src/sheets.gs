/**
 * sheets.gs — Google Sheets操作ユーティリティ
 */

const LOG_HEADERS = [
  'triggered_at', 'completed_at', 'run_id', 'target_period', 'target_month',
  'status', 'result_count', 'raw_file_id', 'raw_file_url',
  'consolidated_status', 'error_detail'
];

/**
 * 名前でシートを取得する。存在しない場合は自動作成。
 */
const getSheet = (sheetName) => {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`Sheet "${sheetName}" created.`);
  }
  return sheet;
};

/**
 * dataシートをクリアしてデータを上書きする。
 */
const overwriteDataSheet = (headers, rows) => {
  if (!rows || rows.length === 0) return;

  const sheet = getSheet('data');
  if (sheet.getLastRow() > 0) {
    sheet.clear();
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
};

/**
 * 月別シート（YYYY-MM形式）にデータを上書きする。
 */
const overwriteMonthlySheet = (sheetName, headers, rows) => {
  if (!rows || rows.length === 0) return;

  const sheet = getSheet(sheetName);
  if (sheet.getLastRow() > 0) {
    sheet.clear();
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log(`[MONTHLY] ${sheetName}: ${rows.length} rows written.`);
};

/**
 * シート名がYYYY-MM形式の月別シートかどうか判定する。
 */
const isMonthlySheetName = (name) => {
  return /^\d{4}-\d{2}$/.test(name);
};

/**
 * 全月別シートの名前リストを取得する（ソート済み）。
 */
const listMonthlySheets = () => {
  const ss = getSpreadsheet();
  return ss.getSheets()
    .map((s) => s.getName())
    .filter(isMonthlySheetName)
    .sort();
};

/** 統合対象から除外するシート名 */
const EXCLUDED_SHEETS = new Set(['log', 'settings']);

/**
 * 統合対象の全シート名を取得する（統合シート自身と除外リストを除く）。
 */
const listSourceSheets = () => {
  const ss = getSpreadsheet();
  const consolidatedName = getConsolidatedSheetName();
  return ss.getSheets()
    .map((s) => s.getName())
    .filter((name) => name !== consolidatedName && !EXCLUDED_SHEETS.has(name))
    .sort();
};

/**
 * 統合シートを更新する。
 * 全ソースシート（統合・log・settings以外）を結合して統合シートを再生成する。
 */
const rebuildConsolidatedSheet = (updatedMonths) => {
  return fullRebuildConsolidatedSheet();
};

/**
 * 全ソースシートを結合して統合シートを生成する。
 */
const fullRebuildConsolidatedSheet = () => {
  const ss = getSpreadsheet();
  const consolidatedName = getConsolidatedSheetName();
  const tmpName = consolidatedName + '_tmp';
  const sourceNames = listSourceSheets();

  if (sourceNames.length === 0) {
    Logger.log('[CONSOLIDATED SKIP] No source sheets found.');
    return 'skip';
  }

  let tmpSheet = ss.getSheetByName(tmpName);
  if (tmpSheet) ss.deleteSheet(tmpSheet);
  tmpSheet = ss.insertSheet(tmpName);

  SpreadsheetApp.flush();

  let headers = null;
  let allRows = [];

  for (const name of sourceNames) {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) continue;

    const sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

    if (!headers) {
      headers = sheetHeaders;
    }

    const rows = readSheetDisplayRows(ss, name);
    allRows = [...allRows, ...rows];

    Logger.log(`[CONSOLIDATED] "${name}": ${rows.length} rows`);
  }

  if (!headers || allRows.length === 0) {
    ss.deleteSheet(tmpSheet);
    Logger.log('[CONSOLIDATED SKIP] No data in source sheets.');
    return 'skip';
  }

  writeRowsToSheet(tmpSheet, headers, allRows);

  const oldSheet = ss.getSheetByName(consolidatedName);
  if (oldSheet) ss.deleteSheet(oldSheet);
  tmpSheet.setName(consolidatedName);

  Logger.log(`[CONSOLIDATED OK] ${consolidatedName}: ${allRows.length} rows from ${sourceNames.length} sheets.`);
  return 'ok';
};

/**
 * シートからデータ行をすべて表示値（文字列）として読み取る。ヘッダーは含まない。
 */
const readSheetDisplayRows = (ss, sheetName) => {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const display = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  const rows = [];
  for (let i = 1; i < display.length; i++) {
    const row = display[i];
    if (row.every((cell) => cell === '')) continue;
    rows.push(row);
  }
  return rows;
};

/**
 * シートにヘッダーとデータ行を書き込む。
 * 日付自動変換を防ぐため、全セルを書式なしテキスト(@)に設定してから書き込む。
 */
const writeRowsToSheet = (sheet, headers, rows) => {
  sheet.clear();

  // 全データ範囲を書式なしテキストに設定（日付自動変換防止）
  const totalRows = rows.length + 1;
  const totalCols = headers.length;
  sheet.getRange(1, 1, totalRows, totalCols).setNumberFormat('@');

  // ヘッダーとデータを書き込み
  sheet.getRange(1, 1, 1, totalCols).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
};

/**
 * 統合シートを手動でフルリビルドする。GASエディタから実行する想定。
 */
function rebuildConsolidatedManual() {
  const result = fullRebuildConsolidatedSheet();
  Logger.log(`[MANUAL] Consolidated rebuild result: ${result}`);
}

/**
 * Apifyタスクの期間が変わったかチェックし、変わっていたらdataシートをアーカイブする。
 */
const checkAndArchive = () => {
  const props = PropertiesService.getScriptProperties();
  const currentPeriod = props.getProperty('CURRENT_PERIOD');
  if (!currentPeriod) return;
  const lastPeriod = props.getProperty('LAST_PERIOD');

  props.setProperty('LAST_PERIOD', currentPeriod);

  if (!lastPeriod || lastPeriod === currentPeriod) {
    return;
  }

  const ss = getSpreadsheet();
  const dataSheet = ss.getSheetByName('data');
  if (!dataSheet || dataSheet.getLastRow() <= 1) {
    return;
  }

  const archiveName = formatPeriodName(lastPeriod);

  let finalName = archiveName;
  let counter = 1;
  while (ss.getSheetByName(finalName)) {
    finalName = `${archiveName}_${counter}`;
    counter++;
  }

  dataSheet.setName(finalName);
  Logger.log(`[ARCHIVE] data sheet archived as "${finalName}"`);
};

/**
 * 期間文字列 "YYYY-MM-DD_YYYY-MM-DD" を "M/D-M/D" 形式に変換する。
 */
const formatPeriodName = (periodStr) => {
  const [sinceStr, untilStr] = periodStr.split('_');
  const formatDate = (dateStr) => {
    if (!dateStr) return '?';
    const d = new Date(dateStr);
    return `${d.getMonth() + 1}/${d.getDate()}`;
  };
  return `${formatDate(untilStr)}-${formatDate(sinceStr)}`;
};

/**
 * rawデータを別スプレッドシートとしてGoogle Driveに保存する。
 */
const saveRawSpreadsheet = (headers, rows, meta) => {
  const folderId = getOptionalConfig('RAW_OUTPUT_FOLDER_ID');
  if (!folderId) {
    Logger.log('[RAW SKIP] RAW_OUTPUT_FOLDER_ID is not set.');
    return null;
  }

  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const fileName = `raw_apify_${now}`;
  const monthLabel = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');

  Logger.log(`[RAW STEP1] Getting/creating subfolder "${monthLabel}" in parent "${folderId}"`);
  const monthFolderId = getOrCreateSubfolderId(folderId, monthLabel);
  Logger.log(`[RAW STEP1 OK] monthFolderId: ${monthFolderId}`);

  Logger.log(`[RAW STEP2] Creating spreadsheet "${fileName}" in folder "${monthFolderId}"`);
  const fileId = createSpreadsheetInFolder(fileName, monthFolderId);
  Logger.log(`[RAW STEP2 OK] fileId: ${fileId}`);

  Logger.log(`[RAW STEP3] Writing data to spreadsheet`);
  writeRawData(fileId, headers, rows, meta);
  Logger.log(`[RAW STEP3 OK] Done`);

  const ss = SpreadsheetApp.openById(fileId);
  Logger.log(`[RAW OK] Saved "${fileName}" to Drive folder "${monthLabel}". FileID: ${fileId}`);
  return {
    fileId: fileId,
    fileUrl: ss.getUrl(),
  };
};

/**
 * 親フォルダ内にサブフォルダIDを取得する。なければDrive APIで作成。
 */
const getOrCreateSubfolderId = (parentFolderId, folderName) => {
  const query = `'${parentFolderId}' in parents and name = '${folderName}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`;
  const result = Drive.Files.list({
    q: query,
    fields: 'files(id)',
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
  });

  if (result.files && result.files.length > 0) {
    Logger.log(`[DRIVE] Found existing subfolder "${folderName}" (ID: ${result.files[0].id})`);
    return result.files[0].id;
  }

  const folderResource = {
    name: folderName,
    mimeType: 'application/vnd.google-apps.folder',
    parents: [parentFolderId],
  };
  const created = Drive.Files.create(folderResource, null, { supportsAllDrives: true });
  Logger.log(`[DRIVE] Created subfolder "${folderName}" (ID: ${created.id})`);
  return created.id;
};

/**
 * Drive API v3 で指定フォルダ内にスプレッドシートを作成し、IDを返す。
 */
const createSpreadsheetInFolder = (fileName, folderId) => {
  const fileResource = {
    name: fileName,
    mimeType: 'application/vnd.google-apps.spreadsheet',
    parents: [folderId],
  };
  const created = Drive.Files.create(fileResource, null, { supportsAllDrives: true });
  return created.id;
};

/**
 * スプレッドシートにraw_dataとmetaシートのデータを書き込む。
 */
const writeRawData = (fileId, headers, rows, meta) => {
  const ss = SpreadsheetApp.openById(fileId);

  const rawSheet = ss.getActiveSheet();
  rawSheet.setName('raw_data');
  if (headers && rows && rows.length > 0) {
    rawSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    rawSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  const metaSheet = ss.insertSheet('meta');
  const metaHeaders = ['run_id', 'dataset_id', 'fetched_at', 'target_period', 'record_count'];
  const metaValues = [
    meta.runId || '',
    meta.datasetId || '',
    meta.fetchedAt || '',
    meta.targetPeriod || '',
    rows ? rows.length : 0,
  ];
  metaSheet.getRange(1, 1, 1, metaHeaders.length).setValues([metaHeaders]);
  metaSheet.getRange(2, 1, 1, metaValues.length).setValues([metaValues]);
};

/**
 * 手動テスト用: raw保存の各ステップを個別に実行して障害箇所を特定する。
 */
function testRawSaveSteps() {
  const folderId = getOptionalConfig('RAW_OUTPUT_FOLDER_ID');
  Logger.log(`[TEST] RAW_OUTPUT_FOLDER_ID: ${folderId}`);
  if (!folderId) {
    Logger.log('[TEST FAIL] RAW_OUTPUT_FOLDER_ID is not set.');
    return;
  }

  try {
    const parentCheck = Drive.Files.get(folderId, { supportsAllDrives: true, fields: 'id,name,mimeType' });
    Logger.log(`[TEST STEP1 OK] Parent folder: name="${parentCheck.name}", mimeType="${parentCheck.mimeType}"`);
  } catch (e) {
    Logger.log(`[TEST STEP1 FAIL] Cannot access parent folder: ${e.message}`);
    return;
  }

  try {
    const subId = getOrCreateSubfolderId(folderId, 'test-subfolder');
    Logger.log(`[TEST STEP2 OK] Subfolder ID: ${subId}`);
  } catch (e) {
    Logger.log(`[TEST STEP2 FAIL] Cannot create subfolder: ${e.message}`);
    return;
  }

  try {
    const subId = getOrCreateSubfolderId(folderId, 'test-subfolder');
    const ssId = createSpreadsheetInFolder('test_raw_save', subId);
    Logger.log(`[TEST STEP3 OK] Spreadsheet created. ID: ${ssId}`);

    writeRawData(ssId, ['col1', 'col2'], [['a', 'b']], { runId: 'test', datasetId: 'test', fetchedAt: new Date() });
    Logger.log(`[TEST STEP4 OK] Data written.`);

    const ss = SpreadsheetApp.openById(ssId);
    Logger.log(`[TEST DONE] URL: ${ss.getUrl()}`);
  } catch (e) {
    Logger.log(`[TEST STEP3-4 FAIL] ${e.message}`);
    return;
  }
}

/**
 * logシートに1行追記する。
 */
const appendLogRow = (logData) => {
  const sheet = getSheet('log');
  const row = [
    logData.triggered_at || '',
    logData.completed_at || '',
    logData.run_id || '',
    logData.target_period || '',
    logData.target_month || '',
    logData.status || '',
    logData.result_count || '',
    logData.raw_file_id || '',
    logData.raw_file_url || '',
    logData.consolidated_status || '',
    logData.error_detail || '',
  ];
  sheet.appendRow(row);
};

/**
 * logシートのrunId一致行を更新する（下から検索）。
 */
const updateLogRow = (runId, updates) => {
  const sheet = getSheet('log');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const runIds = sheet.getRange(2, 3, lastRow - 1, 1).getValues();

  let targetRow = -1;
  for (let i = runIds.length - 1; i >= 0; i--) {
    if (runIds[i][0] === runId) {
      targetRow = i + 2;
      break;
    }
  }

  if (targetRow === -1) {
    Logger.log(`Warning: RunID "${runId}" not found in log sheet.`);
    return;
  }

  const columnMap = {
    completed_at: 2,
    target_period: 4,
    target_month: 5,
    status: 6,
    result_count: 7,
    raw_file_id: 8,
    raw_file_url: 9,
    consolidated_status: 10,
    error_detail: 11,
  };
  for (const [key, col] of Object.entries(columnMap)) {
    if (updates[key] !== undefined) {
      sheet.getRange(targetRow, col).setValue(updates[key]);
    }
  }
};

/**
 * logシートにヘッダー行を書き込む（初回セットアップ用）。
 */
const initializeLogHeaders = () => {
  const sheet = getSheet('log');
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
    Logger.log('Log headers initialized.');
  } else {
    Logger.log('Log sheet already has data. Skipping header initialization.');
  }
};
