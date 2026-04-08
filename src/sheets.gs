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
 * 統合シートを更新する。
 * updatedMonths に含まれる月の行だけ差し替え、それ以外の月はそのまま保持する。
 * 統合シートが存在しない場合はフルリビルドする。
 */
const rebuildConsolidatedSheet = (updatedMonths) => {
  const ss = getSpreadsheet();
  const consolidatedName = getConsolidatedSheetName();
  const consolidatedSheet = ss.getSheetByName(consolidatedName);

  // 統合シートが存在しない場合はフルリビルド
  if (!consolidatedSheet) {
    return fullRebuildConsolidatedSheet();
  }

  // updatedMonths が未指定の場合もフルリビルド
  if (!updatedMonths || updatedMonths.length === 0) {
    return fullRebuildConsolidatedSheet();
  }

  SpreadsheetApp.flush();

  const lastRow = consolidatedSheet.getLastRow();
  const lastCol = consolidatedSheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) {
    return fullRebuildConsolidatedSheet();
  }

  const range = consolidatedSheet.getRange(1, 1, lastRow, lastCol);
  const allData = range.getValues();
  const displayData = range.getDisplayValues();
  const headers = allData[0];

  // uploadedAtFormatted列のインデックスを取得
  const dateColIndex = headers.indexOf('uploadedAtFormatted');
  if (dateColIndex === -1) {
    Logger.log('[CONSOLIDATED WARN] uploadedAtFormatted column not found. Falling back to full rebuild.');
    return fullRebuildConsolidatedSheet();
  }

  // 日付変換されてしまう列を特定
  const dateSafeCols = findDateSafeCols(headers);

  const updatedSet = new Set(updatedMonths);

  // 更新対象月以外の行を保持（手動貼り付けデータ含む）
  const keptRows = [];
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const dateVal = String(row[dateColIndex] || '');
    const match = dateVal.match(/^(\d{4}-\d{2})/);
    const rowMonth = match ? match[1] : 'unknown';

    if (!updatedSet.has(rowMonth)) {
      // 保持する行の日付列を表示値で補正
      for (const c of dateSafeCols) {
        row[c] = displayData[i][c];
      }
      keptRows.push(row);
    }
  }

  // 更新対象月のデータを月別シートから取得（表示値で日付列を補正）
  const newRows = readMonthlySheetRows(ss, updatedMonths, headers);

  const finalRows = [...keptRows, ...newRows];

  if (finalRows.length === 0) {
    Logger.log('[CONSOLIDATED SKIP] No data after merge.');
    return 'skip';
  }

  // 統合シートを上書き
  consolidatedSheet.clear();
  consolidatedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  consolidatedSheet.getRange(2, 1, finalRows.length, finalRows[0].length).setValues(finalRows);

  Logger.log(`[CONSOLIDATED OK] ${consolidatedName}: ${finalRows.length} rows (kept: ${keptRows.length}, updated: ${newRows.length} from ${updatedMonths.join(',')}).`);
  return 'ok';
};

/**
 * 統合シートをゼロからフルリビルドする。初回作成時や復旧用。
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

  // 既存の統合シートから手動貼り付けデータを保持
  const manualRows = extractManualRows(ss, consolidatedName, monthlyNames);

  let tmpSheet = ss.getSheetByName(tmpName);
  if (tmpSheet) ss.deleteSheet(tmpSheet);
  tmpSheet = ss.insertSheet(tmpName);

  SpreadsheetApp.flush();

  let headers = null;
  let allRows = [];

  for (const name of monthlyNames) {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) continue;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const range = sheet.getRange(1, 1, lastRow, lastCol);
    const data = range.getValues();
    const display = range.getDisplayValues();

    if (!headers) {
      headers = data[0];
    }

    const dateSafeCols = findDateSafeCols(headers);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row.every((cell) => cell === '' || cell === null)) continue;
      for (const c of dateSafeCols) {
        row[c] = display[i][c];
      }
      allRows.push(row);
    }
  }

  // 手動データを追加
  allRows = [...allRows, ...manualRows];

  if (!headers || allRows.length === 0) {
    ss.deleteSheet(tmpSheet);
    Logger.log('[CONSOLIDATED SKIP] No data in monthly sheets.');
    return 'skip';
  }

  tmpSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  tmpSheet.getRange(2, 1, allRows.length, allRows[0].length).setValues(allRows);

  const oldSheet = ss.getSheetByName(consolidatedName);
  if (oldSheet) ss.deleteSheet(oldSheet);
  tmpSheet.setName(consolidatedName);

  Logger.log(`[CONSOLIDATED FULL OK] ${consolidatedName}: ${allRows.length} rows (manual: ${manualRows.length}) from ${monthlyNames.length} monthly sheets.`);
  return 'ok';
};

/**
 * 月別シートからデータ行を読み取る。日付列は表示値で補正。
 */
const readMonthlySheetRows = (ss, monthNames, headers) => {
  const dateSafeCols = findDateSafeCols(headers);
  const rows = [];

  for (const monthName of monthNames) {
    const sheet = ss.getSheetByName(monthName);
    if (!sheet || sheet.getLastRow() <= 1) continue;

    const mLastRow = sheet.getLastRow();
    const mLastCol = sheet.getLastColumn();
    const range = sheet.getRange(1, 1, mLastRow, mLastCol);
    const mData = range.getValues();
    const mDisplay = range.getDisplayValues();

    for (let i = 1; i < mData.length; i++) {
      const row = mData[i];
      if (row.every((cell) => cell === '' || cell === null)) continue;
      for (const c of dateSafeCols) {
        row[c] = mDisplay[i][c];
      }
      rows.push(row);
    }
  }

  return rows;
};

/**
 * upload_month, upload_date 列のインデックスを返す。
 */
const findDateSafeCols = (headers) => {
  const cols = [];
  headers.forEach((h, i) => {
    if (h === 'upload_month' || h === 'upload_date') {
      cols.push(i);
    }
  });
  return cols;
};

/**
 * 統合シートからどの月別シートにも属さない手動データ行を抽出する。
 */
const extractManualRows = (ss, consolidatedName, monthlyNames) => {
  const consolidatedSheet = ss.getSheetByName(consolidatedName);
  if (!consolidatedSheet || consolidatedSheet.getLastRow() <= 1) return [];

  const lastRow = consolidatedSheet.getLastRow();
  const lastCol = consolidatedSheet.getLastColumn();
  const range = consolidatedSheet.getRange(1, 1, lastRow, lastCol);
  const allData = range.getValues();
  const displayData = range.getDisplayValues();
  const headers = allData[0];

  const dateColIndex = headers.indexOf('uploadedAtFormatted');
  const monthlySet = new Set(monthlyNames);
  const dateSafeCols = findDateSafeCols(headers);

  const manualRows = [];
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    let rowMonth = 'unknown';
    if (dateColIndex !== -1) {
      const dateVal = String(row[dateColIndex] || '');
      const match = dateVal.match(/^(\d{4}-\d{2})/);
      if (match) rowMonth = match[1];
    }

    // 月別シートに対応しない行 = 手動データ
    if (!monthlySet.has(rowMonth)) {
      for (const c of dateSafeCols) {
        row[c] = displayData[i][c];
      }
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
/**
 * rawデータを別スプレッドシートとしてGoogle Driveに保存する。
 * 各ステップを個別メソッドに分割して障害箇所を特定しやすくしている。
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

  // Step 1: 月別サブフォルダ取得/作成
  Logger.log(`[RAW STEP1] Getting/creating subfolder "${monthLabel}" in parent "${folderId}"`);
  const monthFolderId = getOrCreateSubfolderId(folderId, monthLabel);
  Logger.log(`[RAW STEP1 OK] monthFolderId: ${monthFolderId}`);

  // Step 2: スプレッドシート作成（フォルダ内）
  Logger.log(`[RAW STEP2] Creating spreadsheet "${fileName}" in folder "${monthFolderId}"`);
  const fileId = createSpreadsheetInFolder(fileName, monthFolderId);
  Logger.log(`[RAW STEP2 OK] fileId: ${fileId}`);

  // Step 3: データ書き込み
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

  // raw_data シート
  const rawSheet = ss.getActiveSheet();
  rawSheet.setName('raw_data');
  if (headers && rows && rows.length > 0) {
    rawSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    rawSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  // meta シート
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

  // Step 1: 親フォルダにアクセスできるか
  try {
    const parentCheck = Drive.Files.get(folderId, { supportsAllDrives: true, fields: 'id,name,mimeType' });
    Logger.log(`[TEST STEP1 OK] Parent folder: name="${parentCheck.name}", mimeType="${parentCheck.mimeType}"`);
  } catch (e) {
    Logger.log(`[TEST STEP1 FAIL] Cannot access parent folder: ${e.message}`);
    return;
  }

  // Step 2: サブフォルダ作成
  try {
    const subId = getOrCreateSubfolderId(folderId, 'test-subfolder');
    Logger.log(`[TEST STEP2 OK] Subfolder ID: ${subId}`);
  } catch (e) {
    Logger.log(`[TEST STEP2 FAIL] Cannot create subfolder: ${e.message}`);
    return;
  }

  // Step 3: スプレッドシート作成
  try {
    const subId = getOrCreateSubfolderId(folderId, 'test-subfolder');
    const ssId = createSpreadsheetInFolder('test_raw_save', subId);
    Logger.log(`[TEST STEP3 OK] Spreadsheet created. ID: ${ssId}`);

    // Step 4: データ書き込み
    writeRawData(ssId, ['col1', 'col2'], [['a', 'b']], { runId: 'test', datasetId: 'test', fetchedAt: new Date() });
    Logger.log(`[TEST STEP4 OK] Data written.`);

    const ss = SpreadsheetApp.openById(ssId);
    Logger.log(`[TEST DONE] URL: ${ss.getUrl()}`);
  } catch (e) {
    Logger.log(`[TEST STEP3-4 FAIL] ${e.message}`);
    return;
  }
};

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
