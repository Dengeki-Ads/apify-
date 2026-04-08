/**
 * sync.gs — 統合シートを Looker Studio 用スプレッドシートに同期
 */

/**
 * 指定シートの内容を Looker Studio 用スプレッドシートに上書き同期する。
 * DEST_SPREADSHEET_ID または DEST_SHEET_NAME が未設定の場合はスキップ。
 */
const syncToLookerStudio = (sourceSheetName) => {
  const effectiveSource = sourceSheetName || getConsolidatedSheetName();

  const props = PropertiesService.getScriptProperties();
  const destId = props.getProperty('DEST_SPREADSHEET_ID');
  const destSheetName = props.getProperty('DEST_SHEET_NAME');

  if (!destId || !destSheetName) {
    Logger.log('[SYNC SKIP] DEST_SPREADSHEET_ID or DEST_SHEET_NAME is not set.');
    return;
  }

  // 数式の計算結果を確定させる
  SpreadsheetApp.flush();

  const ss = getSpreadsheet();
  const sourceSheet = ss.getSheetByName(effectiveSource);
  if (!sourceSheet) {
    Logger.log(`[SYNC SKIP] Source sheet "${effectiveSource}" not found.`);
    return;
  }

  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) {
    Logger.log(`[SYNC SKIP] ${effectiveSource} sheet is empty. Skipping to avoid overwriting with empty data.`);
    return;
  }

  const range = sourceSheet.getRange(1, 1, lastRow, lastCol);
  const data = range.getValues();
  const displayData = range.getDisplayValues();

  // 日付オブジェクトに変換されてしまう列を表示値で置換
  const headers = data[0];
  const dateSafeCols = [];
  headers.forEach((h, i) => {
    if (h === 'upload_month' || h === 'upload_date') {
      dateSafeCols.push(i);
    }
  });

  if (dateSafeCols.length > 0) {
    for (let r = 1; r < data.length; r++) {
      for (const c of dateSafeCols) {
        data[r][c] = displayData[r][c];
      }
    }
  }

  const destSs = SpreadsheetApp.openById(destId);
  let destSheet = destSs.getSheetByName(destSheetName);
  if (!destSheet) {
    destSheet = destSs.insertSheet(destSheetName);
    Logger.log(`[SYNC] Destination sheet "${destSheetName}" created.`);
  }

  destSheet.clearContents();
  destSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  Logger.log(`[SYNC OK] Synced ${lastRow} rows from "${effectiveSource}" to Looker Studio spreadsheet.`);
};
