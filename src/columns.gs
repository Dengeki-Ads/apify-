/**
 * columns.gs — シートの不要列自動削除・補助列追加
 */

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
    formulas.push([`=IFERROR(TEXT(REGEXEXTRACT(${dateColLetter}${row},"(\\\d{4})"),"0")&"年"&TEXT(REGEXEXTRACT(${dateColLetter}${row},"-(\\\d{2})-")*1,"0")&"月","")`]);
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
 * 統合シートにupload_monthとupload_dateを追加する。手動で1回実行する想定。
 */
function addDateColumnsToConsolidated() {
  const sheetName = getConsolidatedSheetName();
  addUploadMonthColumn(sheetName);
  addUploadDateColumn(sheetName);
  Logger.log(`[MANUAL OK] Added upload_month and upload_date to "${sheetName}".`);
}

// ============================================================================
// メモリ上の (headers, rows) を操作する整形関数群
// process.gs の processSheetData から呼び出される。シートには触らない。
// ============================================================================

/**
 * COLUMNS_TO_KEEP の列だけを残した新しい headers / rows を返す。
 */
const filterColumnsInMemory = (headers, rows) => {
  const prop = PropertiesService.getScriptProperties().getProperty('COLUMNS_TO_KEEP');
  if (!prop) return { headers, rows };

  const keepSet = new Set(getColumnsToKeep());
  const keepIndices = [];
  const newHeaders = [];
  headers.forEach((h, i) => {
    if (keepSet.has(h)) {
      keepIndices.push(i);
      newHeaders.push(h);
    }
  });

  // フェイルセーフ: COLUMNS_TO_KEEP がどのヘッダーにも一致しない場合、
  // 列を全捨てして空シートにするのではなく、全列を残して警告する。
  // （区切り文字の記法揺れや列名変更で丸ごとデータが消える事故を防ぐ）
  if (newHeaders.length === 0) {
    Logger.log(`[COLUMNS WARN] COLUMNS_TO_KEEP matched no headers. Keeping all columns as fail-safe. headers=${JSON.stringify(headers)}`);
    return { headers, rows };
  }

  if (newHeaders.length === headers.length) {
    return { headers, rows };
  }

  const newRows = rows.map((row) => keepIndices.map((i) => (row[i] !== undefined ? row[i] : '')));
  return { headers: newHeaders, rows: newRows };
};

/**
 * uploadedAtFormatted の値から "YYYY年X月" を計算する。
 * 形式が想定外なら "" を返す。
 */
const computeUploadMonthValue = (raw) => {
  const str = String(raw == null ? '' : raw);
  const yearMatch = str.match(/(\d{4})/);
  const monthMatch = str.match(/-(\d{2})-/);
  if (!yearMatch || !monthMatch) return '';
  return `${yearMatch[1]}年${parseInt(monthMatch[1], 10)}月`;
};

/**
 * uploadedAtFormatted の先頭10文字を "YYYY/MM/DD" に変換する。
 */
const computeUploadDateValue = (raw) => {
  const str = String(raw == null ? '' : raw).substring(0, 10);
  const m = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  return `${m[1]}/${m[2]}/${m[3]}`;
};

/**
 * hashtags 値から keyword をマッチさせて抽出。replaceRule があれば適用。
 * REGEXEXTRACT 互換: キャプチャグループがあれば最初のグループを返す、なければ全マッチを返す。
 */
const computeExtractValue = (raw, keyword, replaceRule) => {
  let str = String(raw == null ? '' : raw).toLowerCase().trim();
  if (replaceRule) {
    const [from, to] = replaceRule.split(':').map((s) => s.replace(/^"+|"+$/g, ''));
    try {
      str = str.replace(new RegExp(from, 'g'), to);
    } catch (e) {
      return '';
    }
  }
  try {
    // 完全一致: keyword がハッシュタグの区切り（引用符・空白・カンマ・括弧・# 等）で
    // 挟まれた完全なトークンと一致する場合のみ抽出する。
    // 例) keyword="sh" は "shop" には一致せず、"sh" 単体にのみ一致する。
    // 日本語等の非ASCII値でも機能するよう \b ではなく区切り文字の有無で境界判定する。
    const boundary = `[^\\s"',\\[\\]{}#:]`;
    const m = str.match(new RegExp(`(?<!${boundary})(?:${keyword})(?!${boundary})`));
    if (!m) return '';
    return m.length > 1 ? m[1] : m[0];
  } catch (e) {
    return '';
  }
};

/**
 * upload_month 列を計算値として追加する。
 */
const addUploadMonthColumnInMemory = (headers, rows) => {
  const dateCol = headers.indexOf('uploadedAtFormatted');
  if (dateCol === -1) return { headers, rows };

  return appendComputedColumn(headers, rows, 'upload_month', (row) => computeUploadMonthValue(row[dateCol]));
};

/**
 * upload_date 列を計算値として追加する。
 */
const addUploadDateColumnInMemory = (headers, rows) => {
  const dateCol = headers.indexOf('uploadedAtFormatted');
  if (dateCol === -1) return { headers, rows };

  return appendComputedColumn(headers, rows, 'upload_date', (row) => computeUploadDateValue(row[dateCol]));
};

/**
 * propertyKey で指定したキーワードを hashtags 列から抽出した値を headerName 列に追加する。
 */
const addExtractColumnInMemory = (headers, rows, propertyKey, headerName) => {
  const props = PropertiesService.getScriptProperties();
  const keyword = props.getProperty(propertyKey);
  if (!keyword) return { headers, rows };

  const hashtagCol = headers.indexOf('hashtags');
  if (hashtagCol === -1) return { headers, rows };

  const replaceRule = props.getProperty(`${propertyKey}_Replace`);
  const cleanKeyword = keyword.replace(/^"+|"+$/g, '');

  return appendComputedColumn(headers, rows, headerName, (row) =>
    computeExtractValue(row[hashtagCol], cleanKeyword, replaceRule)
  );
};

/**
 * 計算済み値で1列を追加（上書き）する共通ヘルパー。
 */
const appendComputedColumn = (headers, rows, colName, computeFn) => {
  let colIdx = headers.indexOf(colName);
  let newHeaders = headers;
  if (colIdx === -1) {
    colIdx = headers.length;
    newHeaders = [...headers, colName];
  }
  const newRows = rows.map((row) => {
    const newRow = [...row];
    while (newRow.length <= colIdx) newRow.push('');
    newRow[colIdx] = computeFn(row);
    return newRow;
  });
  return { headers: newHeaders, rows: newRows };
};

/**
 * 月別シート向けの加工処理（メモリ版）。filterColumnsInMemory + 各種補助列追加。
 */
const applySheetTransformationsInMemory = (headers, rows) => {
  let result = filterColumnsInMemory(headers, rows);
  result = addExtractColumnInMemory(result.headers, result.rows, 'UploadedBy', 'uploaded_by');
  result = addExtractColumnInMemory(result.headers, result.rows, 'SponsoredBy', 'sponsored_by');
  result = addUploadMonthColumnInMemory(result.headers, result.rows);
  result = addUploadDateColumnInMemory(result.headers, result.rows);
  return result;
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
