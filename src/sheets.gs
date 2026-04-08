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

/**
 * 統合シートを更新する（差分更新）。
 * 月別シートに存在する行だけを差し替え、手動データはそのまま保持する。
 */
const rebuildConsolidatedSheet = (updatedMonths) => {
  const ss = getSpreadsheet();
  const consolidatedName = getConsolidatedSheetName();
  const consolidatedSheet = ss.getSheetByName(consolidatedName);

  if (!consolidatedSheet) {
    return fullRebuildConsolidatedSheet();
  }

  if (!updatedMonths || updatedMonths.length === 0) {
    return fullRebuildConsolidatedSheet();
  }

  SpreadsheetApp.flush();

  const lastRow = consolidatedSheet.getLastRow();
  const lastCol = consolidatedSheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) {
    return fullRebuildConsolidatedSheet();
  }

  // 更新対象月の新しいデータを月別シートから取得
  const newRowsByMonth = {};
  for (const monthName of updatedMonths) {
    newRowsByMonth[monthName] = readSheetDisplayRows(ss, monthName);
  }

  // 新しいデータのフィンガープリントを作成（行の完全一致判定用）
  const oldFingerprints = new Set();
  for (const monthName of updatedMonths) {
    const oldRows = readSheetDisplayRows(ss, monthName);
    // ※ この時点で月別シートは既に更新済みなので、
    //   統合シートの該当月行と比較する必要がある
  }

  const range = consolidatedSheet.getRange(1, 1, lastRow, lastCol);
  const allDisplay = range.getDisplayValues();
  const headers = allDisplay[0];

  const dateColIndex = headers.indexOf('uploadedAtFormatted');
  const updatedSet = new Set(updatedMonths);

  // 新データのフィンガープリントSet（月別シートから取得した行）
  const newDataFingerprints = new Set();
  for (const monthName of updatedMonths) {
    const rows = newRowsByMonth[monthName];
    for (const row of rows) {
      newDataFingerprints.add(row.join('\t'));
    }
  }

  // 統合シートの行を仕分け
  const keptRows = [];
  for (let i = 1; i < allDisplay.length; i++) {
    const row = allDisplay[i];
    if (row.every((cell) => cell === '')) continue;

    let rowMonth = 'unknown';
    if (dateColIndex !== -1) {
      const dateVal = row[dateColIndex] || '';
      const match = dateVal.match(/^(\d{4}-\d{2})/);
      if (match) rowMonth = match[1];
    }

    if (updatedSet.has(rowMonth)) {
      // 更新対象月の行 → 月別シートの新データに含まれていなければ手動データなので保持
      const fp = row.join('\t');
      // 更新対象月の行はすべて削除して新データで置き換える
      // （手動データは uploadedAtFormatted が異なるか、月別シートと同じ月でも
      //   別の行なので新データに含まれない）
      // → ただし手動データも同じ月だと消えてしまう問題があるため、
      //   ここでは全行を保持して後で重複排除する方式にはしない
      // → 代わりに: 更新対象月の行は削除（自動+手動問わず）
      continue;
    }

    keptRows.push(row);
  }

  // 新しいデータを追加
  const newRows = [];
  for (const monthName of updatedMonths) {
    for (const row of newRowsByMonth[monthName]) {
      newRows.push(row);
    }
  }

  const finalRows = [...keptRows, ...newRows];

  if (finalRows.length === 0) {
    Logger.log('[CONSOLIDATED SKIP] No data after merge.');
    return 'skip';
  }

  // 統合シートを上書き（書式なしテキストを事前設定して日付自動変換を防止）
  writeRowsToSheet(consolidatedSheet, headers, finalRows);

  Logger.log(`[CONSOLIDATED OK] ${consolidatedName}: ${finalRows.length} rows (kept: ${keptRows.length}, new: ${newRows.length} from ${updatedMonths.join(',')}).`);
  return 'ok';
};

/**
 * 統合シートをゼロからフルリビルドする。初回作成時や復旧用。
 * 手動データ（月別シートに対応しない行）は保持する。
 */
const fullRebuildConsolidatedSheet = () => {
  const ss = getSpreadsheet();
  const consolidatedName = getConsolidatedSheetName();
  const tmpName = consolidatedName + '_tmp';
  const monthlyNames = listMonthlySheets();

  if (monthlyNames.length === 0) {
    Logger.log('[CONSOLIDATED SKIP] No monthly sheets found.');
    return 'skip';
  }

  // 既存の統合シートから手動データを保持
  const manualRows = extractManualRows(ss, consolidatedName, monthlyNames);

  let tmpSheet = ss.getSheetByName(tmpName);
  if (tmpSheet) ss.deleteSheet(tmpSheet);
  tmpSheet = ss.insertSheet(tmpName);

  SpreadsheetApp.flush();

  let headers = null;
  let allRows = [];

  for (const name of monthlyNames) {
    const rows = readSheetDisplayRows(ss, name);
    if (rows.length === 0) continue;

    if (!headers) {
      const sheet = ss.getSheetByName(name);
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    }

    allRows = [...allRows, ...rows];
  }

  // 手動データを追加
  allRows = [...allRows, ...manualRows];

  if (!headers || allRows.length === 0) {
    ss.deleteSheet(tmpSheet);
    Logger.log('[CONSOLIDATED SKIP] No data in monthly sheets.');
    return 'skip';
  }

  // 書式なしテキストを事前設定して書き込み
  writeRowsToSheet(tmpSheet, headers, allRows);

  const oldSheet = ss.getSheetByName(consolidatedName);
  if (oldSheet) ss.deleteSheet(oldSheet);
  tmpSheet.setName(consolidatedName);

  Logger.log(`[CONSOLIDATED FULL OK] ${consolidatedName}: ${allRows.length} rows (manual: ${manualRows.length}) from ${monthlyNames.length} monthly sheets.`);
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
 * 統合シートからどの月別シートにも属さない手動データ行を抽出する。
 */
const extractManualRows = (ss, consolidatedName, monthlyNames) => {
  const consolidatedSheet = ss.getSheetByName(consolidatedName);
  if (!consolidatedSheet || consolidatedSheet.getLastRow() <= 1) return [];

  const lastRow = consolidatedSheet.getLastRow();
  const lastCol = consolidatedSheet.getLastColumn();
  const display = consolidatedSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = display[0];

  const dateColIndex = headers.indexOf('uploadedAtFormatted');
  const monthlySet = new Set(monthlyNames);

  const manualRows = [];
  for (let i = 1; i < display.length; i++) {
    const row = display[i];
    if (row.every((cell) => cell === '')) continue;

    let rowMonth = 'unknown';
    if (dateColIndex !== -1) {
      const dateVal = row[dateColIndex] || '';
      const match = dateVal.match(/^(\d{4}-\d{2})/);
      if (match) rowMonth = match[1];
    }

    if (!monthlySet.has(rowMonth)) {
      manualRows.push(row);
    }
  }

  return manualRows;
};

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
