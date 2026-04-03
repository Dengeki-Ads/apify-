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

/**
 * data シートにキーワード抽出列を追加する共通処理。
 * hashtags 列から指定キーワードに一致する値を REGEXEXTRACT で抽出。
 */
const addExtractColumn = (propertyKey, headerName) => {
  const props = PropertiesService.getScriptProperties();
  const keyword = props.getProperty(propertyKey);
  if (!keyword) {
    Logger.log(`[EXTRACT SKIP] ${propertyKey} is not set.`);
    return;
  }

  const sheet = getSheet('data');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log(`[EXTRACT SKIP] data sheet has no data rows.`);
    return;
  }

  // ヘッダーを毎回読み直す（前の列追加で変わっている可能性がある）
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const hashtagCol = headers.indexOf('hashtags');
  if (hashtagCol === -1) {
    Logger.log('[EXTRACT WARN] "hashtags" column not found in data sheet.');
    return;
  }
  const hashtagColLetter = columnToLetter(hashtagCol + 1);

  let formulaColIndex = headers.indexOf(headerName);
  if (formulaColIndex === -1) {
    formulaColIndex = headers.length;
    sheet.getRange(1, formulaColIndex + 1).setValue(headerName);
  }

  // REGEXREPLACE ルールを取得（例: "belmise:ベルミス"）
  const replaceRule = props.getProperty(`${propertyKey}_Replace`);

  // 数式を組み立て: LOWER + TRIM → (REGEXREPLACE) → REGEXEXTRACT
  const formulas = [];
  for (let row = 2; row <= lastRow; row++) {
    let innerExpr = `LOWER(TRIM(${hashtagColLetter}${row}))`;
    if (replaceRule) {
      const [from, to] = replaceRule.split(':');
      innerExpr = `REGEXREPLACE(${innerExpr},"${from}","${to}")`;
    }
    formulas.push([`=IFERROR(REGEXEXTRACT(${innerExpr},"${keyword}"),"")`]);
  }

  sheet.getRange(2, formulaColIndex + 1, formulas.length, 1).setFormulas(formulas);
  Logger.log(`[EXTRACT OK] Added "${headerName}" column with keyword "${keyword}"${replaceRule ? `, replace: ${replaceRule}` : ''}.`);
};

/**
 * UploadedBy / SponsoredBy の2列を追加する。
 */
const addHashtagFormulaColumns = () => {
  addExtractColumn('UploadedBy', 'uploaded_by');
  addExtractColumn('SponsoredBy', 'sponsored_by');
};

/**
 * 列番号（1-indexed）をアルファベット列名に変換する。
 */
const columnToLetter = (col) => {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
};
