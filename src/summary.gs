/**
 * summary.gs — sponsored_by ごとの集計シートを生成する
 *
 * 集計シート(既定名「集計」)を sponsored_by の値ごとに1行で構築する:
 *   - sponsored_by : 案件名
 *   - sum_views    : その案件の全投稿の views 合計
 *   - views_goal   : スクリプトプロパティ VIEWS_GOAL_<sponsored_by> の目標値
 *   - achieved     : sum_views >= views_goal なら 1、未達なら 0(バイナリ)
 *
 * 対象は「VIEWS_GOAL_ プロパティが設定済みの sponsored_by」のみ。
 * データソースは最新月の月別シート(YYYY-MM のうち最新)。processSheetData の統合再構築後に自動更新される。
 */

// 集計シートのヘッダー(この順序で書き込む)
const SUMMARY_HEADERS = ['sponsored_by', 'sum_views', 'views_goal', 'achieved'];

// views_goal を保持するスクリプトプロパティのキー接頭辞。
// 例) VIEWS_GOAL_転スラ = 1000000
const VIEWS_GOAL_PREFIX = 'VIEWS_GOAL_';

/**
 * VIEWS_GOAL_<sponsored_by> 形式のスクリプトプロパティを走査し、
 * sponsored_by -> 目標値(数値) の Map を返す。
 */
const getViewsGoals = () => {
  const props = PropertiesService.getScriptProperties().getProperties();
  const map = new Map();
  Object.keys(props).forEach((key) => {
    if (key.indexOf(VIEWS_GOAL_PREFIX) !== 0) return;
    const sponsor = normalizeSponsor(key.slice(VIEWS_GOAL_PREFIX.length));
    if (!sponsor) return;
    const goal = parseNumericValue(props[key]);
    if (goal === null) {
      Logger.log(`[SUMMARY WARN] Property "${key}" is not numeric: "${props[key]}". Skipped.`);
      return;
    }
    map.set(sponsor, goal);
  });
  return map;
};

/**
 * YYYY-MM 形式の月別シートのうち最新月のシート名を返す。無ければ null。
 * listSourceSheets() の結果(統合・集計・除外シートを除いたもの)から
 * 月別命名のシートだけを抽出し、名前の昇順で最後(=最新月)を返す。
 */
const getLatestMonthSheetName = () => {
  const monthly = listSourceSheets().filter((name) => /^\d{4}-\d{2}$/.test(name));
  if (monthly.length === 0) return null;
  monthly.sort();
  return monthly[monthly.length - 1];
};

/**
 * sponsored_by の照合キーを正規化する。
 * NFKC で全角/半角・濁点の合成差を吸収し、前後空白を除去する。
 * これにより VIEWS_GOAL_<カタカナ> のキーと sheet 上の sponsored_by 値の
 * 表記ズレ(半角カナ/全角カナ、パ vs パ 等)を無視して合算できる。
 */
const normalizeSponsor = (raw) => String(raw == null ? '' : raw).normalize('NFKC').trim();

/**
 * 表示値(文字列。桁区切りカンマ等を含む場合あり)を数値に変換する。
 * 数値として解釈できない場合は null を返す。
 */
const parseNumericValue = (raw) => {
  const cleaned = String(raw == null ? '' : raw).replace(/[^0-9.\-]/g, '');
  if (cleaned === '' || cleaned === '-' || cleaned === '.') return null;
  const n = Number(cleaned);
  return isNaN(n) ? null : n;
};

/**
 * 最新月の月別シートを読み、sponsored_by ごとに views を合計した集計行を計算する。
 * 返り値: { sourceName, rows } (rows は [sponsored_by, sum_views, views_goal, achieved] の配列)。
 * 対象シートが空、月別シートが無い、または VIEWS_GOAL_ プロパティが1件も無い場合は null。
 */
const computeSummaryData = () => {
  const goals = getViewsGoals();
  if (goals.size === 0) {
    Logger.log('[SUMMARY SKIP] No VIEWS_GOAL_* properties are set.');
    return null;
  }

  const ss = getSpreadsheet();
  const sourceName = getLatestMonthSheetName();
  if (!sourceName) {
    Logger.log('[SUMMARY SKIP] No monthly (YYYY-MM) sheet found.');
    return null;
  }
  const sourceSheet = ss.getSheetByName(sourceName);
  if (!sourceSheet || sourceSheet.getLastRow() <= 1) {
    Logger.log(`[SUMMARY SKIP] Latest month sheet "${sourceName}" is empty or missing.`);
    return null;
  }

  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();
  const values = sourceSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0];
  const sponsorCol = headers.indexOf('sponsored_by');
  const viewsCol = headers.indexOf('views');

  if (sponsorCol === -1 || viewsCol === -1) {
    Logger.log(`[SUMMARY SKIP] Missing columns in "${sourceName}". sponsored_by=${sponsorCol}, views=${viewsCol}`);
    return null;
  }

  // goal が設定済みの sponsor だけを対象に views を合計する。
  const sums = {};
  goals.forEach((_goal, sponsor) => { sums[sponsor] = 0; });

  for (let i = 1; i < values.length; i++) {
    const sponsor = normalizeSponsor(values[i][sponsorCol]);
    if (!sponsor || !Object.prototype.hasOwnProperty.call(sums, sponsor)) continue;
    const v = parseNumericValue(values[i][viewsCol]);
    if (v !== null) sums[sponsor] += v;
  }

  // 行を組み立てる(sponsored_by 昇順で決定的に)。
  const sponsors = Array.from(goals.keys()).sort();
  const rows = sponsors.map((sponsor) => {
    const sumViews = sums[sponsor] || 0;
    const goal = goals.get(sponsor);
    const achieved = sumViews >= goal ? 1 : 0; // 同値は達成扱い(1)
    return [sponsor, sumViews, goal, achieved];
  });

  return { sourceName, rows };
};

/**
 * 集計行から「達成済み(achieved===1)」の sponsored_by 名(正規化済み)の Set を返す。
 */
const achievedSponsorsFromRows = (rows) =>
  new Set((rows || []).filter((r) => r[3] === 1).map((r) => r[0]));

/**
 * 最新月の月別シートを読み、sponsored_by ごとに views を合計して集計シートを再構築する。
 * VIEWS_GOAL_ 未設定・月別シート無し・対象シート空などの場合は 'skip'。
 */
const rebuildSummarySheet = () => {
  const data = computeSummaryData();
  if (!data) return 'skip';
  writeSummaryRows(getSummarySheetName(), SUMMARY_HEADERS, data.rows);
  Logger.log(`[SUMMARY OK] "${getSummarySheetName()}": ${data.rows.length} sponsor row(s) written.`);
  return 'ok';
};

/**
 * 集計シートにヘッダーと数値行を書き込む。
 * sum_views / views_goal / achieved は数値のまま保持する(集計・比較用)。
 */
const writeSummaryRows = (sheetName, headers, rows) => {
  const sheet = getSheet(sheetName);
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
};

/**
 * 手動実行用: 集計シートを今すぐ再構築する。GASエディタから実行する想定。
 */
function rebuildSummaryManual() {
  const result = rebuildSummarySheet();
  Logger.log(`[MANUAL] Summary rebuild result: ${result}`);
}
