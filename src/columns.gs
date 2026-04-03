/**
 * columns.gs — data シートの不要列自動削除
 */

/**
 * data シートからホワイトリスト外の列を削除する。
 * COLUMNS_TO_KEEP が未設定の場合はスキップ（既存動作に影響しない）。
 */
const filterColumns = () => {
  const prop = PropertiesService.getScriptProperties().getProperty('COLUMNS_TO_KEEP');
  if (!prop) {
    Logger.log('[COLUMNS SKIP] COLUMNS_TO_KEEP is not set. Skipping column filter.');
    return;
  }

  const sheet = getSheet('data');
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    Logger.log('[COLUMNS SKIP] data sheet is empty.');
    return;
  }

  const columnsToKeep = getColumnsToKeep();
  const keepSet = new Set(columnsToKeep);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 保持リストにあるがシートに存在しない列を警告
  const headerSet = new Set(headers);
  const missing = columnsToKeep.filter((col) => !headerSet.has(col));
  if (missing.length > 0) {
    Logger.log(`[COLUMNS WARN] These columns are in COLUMNS_TO_KEEP but not found in sheet: ${missing.join(', ')}`);
  }

  // 削除対象の列番号を収集（1-indexed）
  const columnsToDelete = [];
  headers.forEach((header, index) => {
    if (!keepSet.has(header)) {
      columnsToDelete.push({ col: index + 1, name: header });
    }
  });

  if (columnsToDelete.length === 0) {
    Logger.log('[COLUMNS SKIP] No columns to delete.');
    return;
  }

  // 降順ソートして後ろから削除
  columnsToDelete.sort((a, b) => b.col - a.col);
  const deletedNames = columnsToDelete.map((c) => c.name);

  for (const { col } of columnsToDelete) {
    sheet.deleteColumn(col);
  }

  Logger.log(`[COLUMNS OK] Deleted ${deletedNames.length} column(s): ${deletedNames.join(', ')}`);
};
