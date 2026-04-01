/**
 * sheets.gs — Google Sheets操作ユーティリティ
 */

const LOG_HEADERS = ['triggered_at', 'completed_at', 'run_id', 'status', 'result_count', 'error_detail'];

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
 * dataシートにデータを追記する（動的ヘッダー対応）。
 */
const appendDataRows = (headers, rows) => {
  if (!rows || rows.length === 0) return;

  const sheet = getSheet('data');
  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    return;
  }

  let existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const existingSet = new Set(existingHeaders);
  const newKeys = headers.filter((h) => !existingSet.has(h));

  if (newKeys.length > 0) {
    const startCol = existingHeaders.length + 1;
    sheet.getRange(1, startCol, 1, newKeys.length).setValues([newKeys]);
    existingHeaders = [...existingHeaders, ...newKeys];
  }

  const headerIndexMap = new Map(headers.map((h, i) => [h, i]));

  const alignedRows = rows.map((row) =>
    existingHeaders.map((h) => {
      const idx = headerIndexMap.get(h);
      return idx !== undefined ? row[idx] : '';
    })
  );

  sheet.getRange(lastRow + 1, 1, alignedRows.length, alignedRows[0].length).setValues(alignedRows);
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
    logData.status || '',
    logData.result_count || '',
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

  const columnMap = { completed_at: 2, status: 4, result_count: 5, error_detail: 6 };
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
