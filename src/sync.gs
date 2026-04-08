/**
 * sync.gs — 統合シートを Looker Studio 用スプレッドシートに同期
 */

/**
 * 指定シートの内容を Looker Studio 用スプレッドシートに上書き同期する。
 * DEST_SPREADSHEET_ID または DEST_SHEET_NAME が未設定の場合はスキップ。
 * 全セルを書式なしテキストに設定して日付自動変換を防止する。
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

  // 表示値（文字列）として取得して日付変換を完全に防止
  const data = sourceSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  const destSs = SpreadsheetApp.openById(destId);
  let destSheet = destSs.getSheetByName(destSheetName);
  if (!destSheet) {
    destSheet = destSs.insertSheet(destSheetName);
    Logger.log(`[SYNC] Destination sheet "${destSheetName}" created.`);
  }

  destSheet.clearContents();

  // 書き込み先を書式なしテキストに設定してから書き込む
  destSheet.getRange(1, 1, data.length, data[0].length).setNumberFormat('@');
  destSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  Logger.log(`[SYNC OK] Synced ${lastRow} rows from "${effectiveSource}" to Looker Studio spreadsheet.`);
};
