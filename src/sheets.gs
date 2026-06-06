/**
 * sheets.gs — Google Sheets操作ユーティリティ
 */

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

/** 統合対象から除外するシート名 */
const EXCLUDED_SHEETS = new Set(['settings', 'log', 'notified']);

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
    // 列数をヘッダーに合わせて揃える
    const normalized = rows.map((row) => {
      if (row.length === totalCols) return row;
      if (row.length > totalCols) return row.slice(0, totalCols);
      return [...row, ...Array(totalCols - row.length).fill('')];
    });
    sheet.getRange(2, 1, normalized.length, totalCols).setValues(normalized);
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

