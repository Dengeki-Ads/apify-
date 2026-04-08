/**
 * columns.gs — シートの不要列自動削除・補助列追加
 */

/**
 * 指定シートからホワイトリスト外の列を削除する。
 * COLUMNS_TO_KEEP が未設定の場合はスキップ。
 */
const filterColumns = (sheetName = 'data') => {
  const prop = PropertiesService.getScriptProperties().getProperty('COLUMNS_TO_KEEP');
  if (!prop) {
    Logger.log(`[COLUMNS SKIP] COLUMNS_TO_KEEP is not set. Skipping column filter.`);
    return;
  }

  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    Logger.log(`[COLUMNS SKIP] ${sheetName} sheet is empty.`);
    return;
  }

  const columnsToKeep = getColumnsToKeep();
  const keepSet = new Set(columnsToKeep);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const headerSet = new Set(headers);
  const missing = columnsToKeep.filter((col) => !headerSet.has(col));
  if (missing.length > 0) {
    Logger.log(`[COLUMNS WARN] These columns are in COLUMNS_TO_KEEP but not found in ${sheetName}: ${missing.join(', ')}`);
  }

  const columnsToDelete = [];
  headers.forEach((header, index) => {
    if (!keepSet.has(header)) {
      columnsToDelete.push({ col: index + 1, name: header });
    }
  });

  if (columnsToDelete.length === 0) {
    Logger.log(`[COLUMNS SKIP] No columns to delete in ${sheetName}.`);
    return;
  }

  columnsToDelete.sort((a, b) => b.col - a.col);
  const deletedNames = columnsToDelete.map((c) => c.name);

  for (const { col } of columnsToDelete) {
    sheet.deleteColumn(col);
  }

  Logger.log(`[COLUMNS OK] Deleted ${deletedNames.length} column(s) from ${sheetName}: ${deletedNames.join(', ')}`);
};

/**
 * 指定シートにキーワード抽出列を追加する共通処理。
 */
const addExtractColumn = (propertyKey, headerName, sheetName = 'data') => {
  const props = PropertiesService.getScriptProperties();
  const keyword = props.getProperty(propertyKey);
  if (!keyword) {
    Logger.log(`[EXTRACT SKIP] ${propertyKey} is not set.`);
    return;
  }

  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log(`[EXTRACT SKIP] ${sheetName} sheet has no data rows.`);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const hashtagCol = headers.indexOf('hashtags');
  if (hashtagCol === -1) {
    Logger.log(`[EXTRACT WARN] "hashtags" column not found in ${sheetName}.`);
    return;
  }
  const hashtagColLetter = columnToLetter(hashtagCol + 1);

  let formulaColIndex = headers.indexOf(headerName);
  if (formulaColIndex === -1) {
    formulaColIndex = headers.length;
    sheet.getRange(1, formulaColIndex + 1).setValue(headerName);
  }

  const replaceRule = props.getProperty(`${propertyKey}_Replace`);
  const cleanKeyword = keyword.replace(/^"+|"+$/g, '');

  const formulas = [];
  for (let row = 2; row <= lastRow; row++) {
    let innerExpr = `LOWER(TRIM(${hashtagColLetter}${row}))`;
    if (replaceRule) {
      const [from, to] = replaceRule.split(':').map((s) => s.replace(/^"+|"+$/g, ''));
      innerExpr = `REGEXREPLACE(${innerExpr},"${from}","${to}")`;
    }
    formulas.push([`=IFERROR(REGEXEXTRACT(${innerExpr},"${cleanKeyword}"),"")`]);
  }

  sheet.getRange(2, formulaColIndex + 1, formulas.length, 1).setFormulas(formulas);
  Logger.log(`[EXTRACT OK] Added "${headerName}" column to ${sheetName} with keyword "${keyword}"${replaceRule ? `, replace: ${replaceRule}` : ''}.`);
};

/**
 * uploadedAtFormatted 列から年月を抽出して「YYYY年X月」形式の列を追加する。
 */
const addUploadMonthColumn = (sheetName = 'data') => {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log(`[MONTH SKIP] ${sheetName} sheet has no data rows.`);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const dateCol = headers.indexOf('uploadedAtFormatted');
  if (dateCol === -1) {
    Logger.log(`[MONTH WARN] "uploadedAtFormatted" column not found in ${sheetName}.`);
    return;
  }
  const dateColLetter = columnToLetter(dateCol + 1);

  const headerName = 'upload_month';
  let formulaColIndex = headers.indexOf(headerName);
  if (formulaColIndex === -1) {
    formulaColIndex = headers.length;
    sheet.getRange(1, formulaColIndex + 1).setValue(headerName);
  }

  const formulas = [];
  for (let row = 2; row <= lastRow; row++) {
    formulas.push([`=IFERROR(REGEXEXTRACT(${dateColLetter}${row},"(\\\d{4})")&"年"&REGEXEXTRACT(${dateColLetter}${row},"-(\\\d{2})-")*1&"月","")`]);
  }

  sheet.getRange(2, formulaColIndex + 1, formulas.length, 1).setFormulas(formulas);
  Logger.log(`[MONTH OK] Added "upload_month" column to ${sheetName}.`);
};

/**
 * uploadedAtFormatted 列からアップロード日（YYYY/MM/DD）を抽出する列を追加する。
 */
const addUploadDateColumn = (sheetName = 'data') => {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log(`[DATE SKIP] ${sheetName} sheet has no data rows.`);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const dateCol = headers.indexOf('uploadedAtFormatted');
  if (dateCol === -1) {
    Logger.log(`[DATE WARN] "uploadedAtFormatted" column not found in ${sheetName}.`);
    return;
  }
  const dateColLetter = columnToLetter(dateCol + 1);

  const headerName = 'upload_date';
  let formulaColIndex = headers.indexOf(headerName);
  if (formulaColIndex === -1) {
    formulaColIndex = headers.length;
    sheet.getRange(1, formulaColIndex + 1).setValue(headerName);
  }

  const formulas = [];
  for (let row = 2; row <= lastRow; row++) {
    formulas.push([`=IFERROR(TEXT(DATEVALUE(LEFT(${dateColLetter}${row},10)),"YYYY/MM/DD"),"")`]);
  }

  sheet.getRange(2, formulaColIndex + 1, formulas.length, 1).setFormulas(formulas);
  Logger.log(`[DATE OK] Added "upload_date" column to ${sheetName}.`);
};

/**
 * UploadedBy / SponsoredBy / upload_month / upload_date の列を追加する。
 */
const addHashtagFormulaColumns = (sheetName = 'data') => {
  addExtractColumn('UploadedBy', 'uploaded_by', sheetName);
  addExtractColumn('SponsoredBy', 'sponsored_by', sheetName);
  addUploadMonthColumn(sheetName);
  addUploadDateColumn(sheetName);
};

/**
 * 月別シート向け加工処理（列フィルタ + 補助列追加）。
 */
const applySheetTransformations = (sheetName) => {
  filterColumns(sheetName);
  addHashtagFormulaColumns(sheetName);
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
