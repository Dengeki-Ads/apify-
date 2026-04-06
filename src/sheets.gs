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
 * 全月別シートを連結して統合シートを再生成する。
 * 一時シートを使い、完成後に入れ替える。
 */
const rebuildConsolidatedSheet = () => {
  const ss = getSpreadsheet();
  const consolidatedName = getConsolidatedSheetName();
  const tmpName = consolidatedName + '_tmp';
  const monthlyNames = listMonthlySheets();

  if (monthlyNames.length === 0) {
    Logger.log('[CONSOLIDATED SKIP] No monthly sheets found.');
    return 'skip';
  }

  // 一時シート作成
  let tmpSheet = ss.getSheetByName(tmpName);
  if (tmpSheet) ss.deleteSheet(tmpSheet);
  tmpSheet = ss.insertSheet(tmpName);

  // 数式の計算結果を確定させる
  SpreadsheetApp.flush();

  let headers = null;
  let allRows = [];

  for (const name of monthlyNames) {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) continue;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    if (!headers) {
      headers = data[0];
    }

    // データ行を追加（ヘッダー除く、空行除外）
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row.every((cell) => cell === '' || cell === null)) continue;
      allRows.push(row);
    }
  }

  if (!headers || allRows.length === 0) {
    ss.deleteSheet(tmpSheet);
    Logger.log('[CONSOLIDATED SKIP] No data in monthly sheets.');
    return 'skip';
  }

  // 一時シートに書き込み
  tmpSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  tmpSheet.getRange(2, 1, allRows.length, allRows[0].length).setValues(allRows);

  // 旧統合シートと入れ替え
  const oldSheet = ss.getSheetByName(consolidatedName);
  if (oldSheet) ss.deleteSheet(oldSheet);
  tmpSheet.setName(consolidatedName);

  Logger.log(`[CONSOLIDATED OK] ${consolidatedName}: ${allRows.length} rows from ${monthlyNames.length} monthly sheets.`);
  return 'ok';
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

  // 月別サブフォルダを取得/作成
  const parentFolder = DriveApp.getFolderById(folderId);
  const monthLabel = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
  const monthFolder = getOrCreateSubfolder(parentFolder, monthLabel);

  // 空のスプレッドシートをフォルダ内に直接作成
  const blankFile = monthFolder.createFile(fileName, '', MimeType.GOOGLE_SHEETS);
  const ss = SpreadsheetApp.openById(blankFile.getId());

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

  Logger.log(`[RAW OK] Saved "${fileName}" to Drive folder "${monthLabel}". FileID: ${ss.getId()}`);
  return {
    fileId: ss.getId(),
    fileUrl: ss.getUrl(),
  };
};

/**
 * 親フォルダ内にサブフォルダを取得する。なければ作成。
 */
const getOrCreateSubfolder = (parentFolder, folderName) => {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
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
