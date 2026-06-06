/**
 * process.gs — シート操作を集約した遅延実行ハンドラ
 *
 * 設計:
 *   - doPost はシートに直接書き込まず、ジョブを ScriptProperties のキューに積み、
 *     processSheetData の時間トリガーをスケジュールするだけ。
 *   - processSheetData は別実行として起動し、最初に統合対象の全ソースシートを
 *     SHEET_CACHE に読み込み(getDisplayValues=文字列のみ)、月別シートは raw データから直接書き戻し、
 *     最後にキャッシュ文字列から統合シートを再生成する(displayify 不要)。
 */

// ============================================================================
// グローバル: 実行スコープのシートキャッシュ
// ============================================================================
// processSheetData の1実行内でのみ有効。実行が終われば破棄される。
// 統合シート構築用に「文字列」で全ソースシートを保持する。
let SHEET_CACHE = null;

const PENDING_JOBS_KEY = 'PENDING_SHEET_JOBS';
const PROCESS_FN_NAME = 'processSheetData';
const PROCESS_DEFER_MS = 60 * 1000; // 1分後

// ============================================================================
// 通知条件の定数とデフォルト値
// ============================================================================
const NOTIFIED_SHEET_NAME = 'notified';
const NOTIFIED_HEADERS = ['key', 'notified_at', 'views', 'sponsored_by', 'upload_date', 'tier'];
const NOTIFY_KEY_COLUMN = 'uploadedAtFormatted';

const DEFAULTS = {
  NOTIFY_VIEWS_MIN: 10000,
  NOTIFY_VIEWS_MIN_100K: 100000,
  NOTIFY_MAX_DAYS: 21,
  NOTIFY_SPONSOR_INCLUDE: 'ルルットリリィ|杖と剣のウィストリア|転スラ|ぷちきゅあ',
  SLACK_MESSAGE_TEXT_100K: 'CVに繋がるコメントお願いします！',
};

/**
 * 通知条件パラメータを ScriptProperties から取得する。未設定時はデフォルト値。
 * リスト系プロパティは | 区切りで保存される運用。
 */
const getNotifyConfig = () => {
  const props = PropertiesService.getScriptProperties();
  return {
    viewsMin: Number(props.getProperty('NOTIFY_VIEWS_MIN') || DEFAULTS.NOTIFY_VIEWS_MIN),
    viewsMin100K: Number(props.getProperty('NOTIFY_VIEWS_MIN_100K') || DEFAULTS.NOTIFY_VIEWS_MIN_100K),
    maxDays: Number(props.getProperty('NOTIFY_MAX_DAYS') || DEFAULTS.NOTIFY_MAX_DAYS),
    sponsorInclude: new Set(
      (props.getProperty('NOTIFY_SPONSOR_INCLUDE') || DEFAULTS.NOTIFY_SPONSOR_INCLUDE)
        .split('|').map((s) => s.trim()).filter((s) => s)
    ),
    messageText: props.getProperty('SLACK_MESSAGE_TEXT') || '',
    messageText100K: props.getProperty('SLACK_MESSAGE_TEXT_100K') || DEFAULTS.SLACK_MESSAGE_TEXT_100K,
  };
};

// ============================================================================
// ジョブキュー (ScriptProperties)
// ============================================================================

const enqueueJob = (job) => {
  const lock = LockService.getScriptLock();
  lock.waitLock(10 * 1000);
  try {
    const props = PropertiesService.getScriptProperties();
    const existing = JSON.parse(props.getProperty(PENDING_JOBS_KEY) || '[]');
    existing.push(job);
    props.setProperty(PENDING_JOBS_KEY, JSON.stringify(existing));
    Logger.log(`[QUEUE] Enqueued job type=${job.type} runId=${job.runId || ''} (queue size: ${existing.length})`);
  } finally {
    lock.releaseLock();
  }
};

const dequeueAllJobs = () => {
  const lock = LockService.getScriptLock();
  lock.waitLock(10 * 1000);
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = JSON.parse(props.getProperty(PENDING_JOBS_KEY) || '[]');
    props.deleteProperty(PENDING_JOBS_KEY);
    // null/undefined を除外して後段の処理を守る
    const jobs = Array.isArray(raw) ? raw.filter((j) => j !== null && j !== undefined) : [];
    if (Array.isArray(raw) && raw.length !== jobs.length) {
      Logger.log(`[QUEUE WARN] Filtered ${raw.length - jobs.length} null/undefined job(s) from queue.`);
    }
    return jobs;
  } finally {
    lock.releaseLock();
  }
};

// ============================================================================
// トリガー管理
// ============================================================================

const scheduleProcessSheetData = () => {
  const existing = ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === PROCESS_FN_NAME);
  for (const t of existing) {
    ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger(PROCESS_FN_NAME)
    .timeBased()
    .after(PROCESS_DEFER_MS)
    .create();
  Logger.log(`[SCHEDULE] ${PROCESS_FN_NAME} scheduled ${PROCESS_DEFER_MS / 1000}s from now. (cleared ${existing.length} existing trigger(s))`);
};

const cleanupFiringTrigger = (triggerUid) => {
  if (!triggerUid) return;
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getUniqueId() === triggerUid)
    .forEach((t) => ScriptApp.deleteTrigger(t));
};

// ============================================================================
// SHEET_CACHE 初期化(全ソースシートを「文字列」で読込)
// ============================================================================

const initSheetCache = () => {
  const ss = getSpreadsheet();
  SHEET_CACHE = {
    sourceSheets: {},          // 'YYYY-MM' などソースシート名 -> { headers, rows } (すべて文字列)
    notified10k: new Set(),    // 1万通知済みのキー(uploadedAtFormatted値)
    notified100k: new Set(),   // 10万通知済みのキー(uploadedAtFormatted値)
    newNotifiedRecords: [],    // 今回新たに通知したレコード(notified シートへ追記)
  };

  const tList = Date.now();
  const sourceNames = listSourceSheets();
  Logger.log(`  [TIMING:init] listSourceSheets (count=${sourceNames.length}): ${((Date.now() - tList) / 1000).toFixed(2)}s`);

  for (const name of sourceNames) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    const tRead = Date.now();
    SHEET_CACHE.sourceSheets[name] = readSheetToMemory(sheet, true); // useDisplayValues=true
    const data = SHEET_CACHE.sourceSheets[name];
    Logger.log(`  [TIMING:init] read "${name}" (rows=${data.rows.length}, cols=${data.headers.length}): ${((Date.now() - tRead) / 1000).toFixed(2)}s`);
  }

  const tNotified = Date.now();
  loadNotifiedKeys();
  Logger.log(`  [TIMING:init] loadNotifiedKeys (10k=${SHEET_CACHE.notified10k.size}, 100k=${SHEET_CACHE.notified100k.size}): ${((Date.now() - tNotified) / 1000).toFixed(2)}s`);
};

/**
 * notified シートからキー列のみを Set として読み込む。
 * シートが存在しなければ何もしない(後で append 時に作る)。
 */
const loadNotifiedKeys = () => {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(NOTIFIED_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return;

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const keyCol = headers.indexOf('key');
  const tierCol = headers.indexOf('tier');

  if (keyCol === -1) {
    Logger.log("[NOTIFIED WARN] 'key' column not found in notified sheet.");
    return;
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
  for (const row of values) {
    const key = String(row[keyCol] || '');
    if (!key) continue;
    const tier = tierCol !== -1 ? String(row[tierCol] || '') : '';
    if (tier === '10w') {
      SHEET_CACHE.notified100k.add(key);
    } else {
      // 空 or '1w' (既存データの暗黙のマイグレーション)
      SHEET_CACHE.notified10k.add(key);
    }
  }
};

/**
 * 今回追加された通知レコードを notified シートに append する。
 * シートがなければ作成しヘッダーも書く。
 */
const flushNewNotifiedRecords = () => {
  const records = SHEET_CACHE.newNotifiedRecords;
  if (!records || records.length === 0) return;

  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(NOTIFIED_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(NOTIFIED_SHEET_NAME);
    sheet.getRange(1, 1, 1, NOTIFIED_HEADERS.length).setValues([NOTIFIED_HEADERS]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, NOTIFIED_HEADERS.length).setValues([NOTIFIED_HEADERS]);
  } else {
    // 既存シートのマイグレーション: 'tier' 列を必要に応じて追加
    const lastCol = sheet.getLastColumn();
    const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (existingHeaders.indexOf('tier') === -1) {
      sheet.getRange(1, lastCol + 1).setValue('tier');
      Logger.log("[NOTIFIED MIGRATE] Added 'tier' column to existing notified sheet.");
    }
  }

  const startRow = sheet.getLastRow() + 1;
  const rows = records.map((r) => [
    r.key,
    r.notified_at,
    r.views,
    r.sponsored_by,
    r.upload_date,
    r.tier || '',
  ]);
  sheet.getRange(startRow, 1, rows.length, NOTIFIED_HEADERS.length).setValues(rows);
  Logger.log(`[NOTIFIED] Appended ${rows.length} record(s) to "${NOTIFIED_SHEET_NAME}" sheet.`);
};

/**
 * Sheet 全体を { headers, rows } のメモリ表現に読み出す。
 * useDisplayValues=true で文字列、false で型を保持して読む。
 */
const readSheetToMemory = (sheet, useDisplayValues) => {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return { headers: [], rows: [] };
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const values = useDisplayValues ? range.getDisplayValues() : range.getValues();
  return {
    headers: values[0] || [],
    rows: values.slice(1).filter((r) => r.some((c) => c !== '' && c !== null && c !== undefined)),
  };
};

/**
 * シートに headers と rows を書き込む(書式は弄らない=型を保持)。
 * 月別シートへの直接書込で使用。
 */
const writeRowsPreserveTypes = (sheet, headers, rows) => {
  sheet.clear();
  if (headers.length === 0) return;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length === 0) return;
  const normalized = rows.map((r) => {
    if (r.length === headers.length) return r;
    if (r.length > headers.length) return r.slice(0, headers.length);
    return [...r, ...Array(headers.length - r.length).fill('')];
  });
  sheet.getRange(2, 1, normalized.length, headers.length).setValues(normalized);
};

// ============================================================================
// 統合シート再構築(キャッシュは既に文字列なので displayify 不要)
// ============================================================================

const rebuildConsolidatedFromCache = () => {
  const lap = (label, t) => Logger.log(`  [TIMING:cons] ${label}: ${((Date.now() - t) / 1000).toFixed(2)}s`);

  const ss = getSpreadsheet();
  const consolidatedName = getConsolidatedSheetName();
  const tmpName = consolidatedName + '_tmp';
  const sourceNames = Object.keys(SHEET_CACHE.sourceSheets).sort();

  if (sourceNames.length === 0) {
    Logger.log('[CONSOLIDATED SKIP] No source sheets in cache.');
    return 'skip';
  }

  const tConcat = Date.now();
  let headers = null;
  let allRows = [];

  for (const name of sourceNames) {
    const data = SHEET_CACHE.sourceSheets[name];
    if (!data || data.rows.length === 0) continue;
    if (!headers) headers = data.headers;
    allRows = [...allRows, ...data.rows];
  }
  lap(`concat (rows=${allRows.length})`, tConcat);

  if (!headers || allRows.length === 0) {
    Logger.log('[CONSOLIDATED SKIP] No data in cached source sheets.');
    return 'skip';
  }

  const tTmp = Date.now();
  let tmpSheet = ss.getSheetByName(tmpName);
  if (tmpSheet) ss.deleteSheet(tmpSheet);
  tmpSheet = ss.insertSheet(tmpName);
  SpreadsheetApp.flush();
  lap('create tmp sheet', tTmp);

  const tWrite = Date.now();
  writeRowsToSheet(tmpSheet, headers, allRows);
  lap(`writeRowsToSheet (rows=${allRows.length}, cols=${headers.length})`, tWrite);

  const tRename = Date.now();
  const oldSheet = ss.getSheetByName(consolidatedName);
  if (oldSheet) ss.deleteSheet(oldSheet);
  tmpSheet.setName(consolidatedName);
  lap('delete old + rename', tRename);

  Logger.log(`[CONSOLIDATED OK] ${consolidatedName}: ${allRows.length} rows from ${sourceNames.length} sheets.`);
  return 'ok';
};

/**
 * 任意の値を「表示用文字列」に変換する。新規データのみに使用(統合シートには不要)。
 */
const displayify = (v) => {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  return String(v);
};

// ============================================================================
// raw spreadsheet からデータを読み戻すヘルパー
// ============================================================================

const readRawSpreadsheetData = (fileId) => {
  const ss = SpreadsheetApp.openById(fileId);
  const rawSheet = ss.getSheetByName('raw_data');
  if (!rawSheet || rawSheet.getLastRow() <= 1) {
    return { headers: [], rows: [] };
  }
  return readSheetToMemory(rawSheet, false);
};

// ============================================================================
// 通知判定ロジック
// ============================================================================

/**
 * upload_date (YYYY/MM/DD) と今日(Asia/Tokyo)の差を日数で返す。
 * パース不能なら Infinity を返して通知対象外にする。
 */
const daysSinceUpload = (uploadDateStr) => {
  const m = String(uploadDateStr || '').match(/^(\d{4})\/(\d{2})\/(\d{2})/);
  if (!m) return Infinity;
  const uploadDate = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const todayM = todayStr.match(/^(\d{4})\/(\d{2})\/(\d{2})/);
  const today = new Date(Number(todayM[1]), Number(todayM[2]) - 1, Number(todayM[3]));
  return Math.floor((today.getTime() - uploadDate.getTime()) / (1000 * 60 * 60 * 24));
};

/**
 * 整形済みデータ(headers, rows)を1行ずつチェックし、通知条件を満たす行に対して通知を実行する。
 * 通知済みのキーは SHEET_CACHE.notifiedKeys と newNotifiedRecords に追加される。
 * 条件: views >= viewsMin && sponsored_by NOT IN exclude && (今日 - upload_date) <= maxDays && 未通知
 */
const checkAndNotifyRows = (headers, rows) => {
  const cfg = getNotifyConfig();
  const keyCol = headers.indexOf(NOTIFY_KEY_COLUMN);
  const viewsCol = headers.indexOf('views');
  const sponsorCol = headers.indexOf('sponsored_by');
  const uploadDateCol = headers.indexOf('upload_date');
  const urlCol = headers.indexOf('postPage');
  const uploadedByCol = headers.indexOf('uploaded_by');

  if (keyCol === -1) {
    Logger.log(`[NOTIFY SKIP] "${NOTIFY_KEY_COLUMN}" column not found.`);
    return 0;
  }
  if (viewsCol === -1 || sponsorCol === -1 || uploadDateCol === -1) {
    Logger.log(`[NOTIFY SKIP] Missing required columns. views=${viewsCol}, sponsored_by=${sponsorCol}, upload_date=${uploadDateCol}`);
    return 0;
  }

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  let notifiedCount = 0;

  for (const row of rows) {
    const key = String(row[keyCol] || '');
    if (!key) continue;

    const views = Number(row[viewsCol]) || 0;
    if (views < cfg.viewsMin) continue; // 1万にも満たないなら何もしない

    const sponsor = String(row[sponsorCol] || '');
    // ホワイトリスト方式: sponsored_by が NOTIFY_SPONSOR_INCLUDE に含まれている時のみ通知。
    // 空 sponsor や未登録案件は通知対象外。
    if (!sponsor || !cfg.sponsorInclude.has(sponsor)) continue;

    const uploadDate = String(row[uploadDateCol] || '');
    const diffDays = daysSinceUpload(uploadDate);
    if (diffDays > cfg.maxDays) continue;

    const videoUrl = urlCol !== -1 ? String(row[urlCol] || '') : '';
    const creator = uploadedByCol !== -1 ? String(row[uploadedByCol] || '') : '';

    // ========== 10万 tier(高い方を優先) ==========
    if (views >= cfg.viewsMin100K) {
      if (SHEET_CACHE.notified100k.has(key)) continue;

      const text = [
        '<!channel>',
        '10万再生突破しました！',
        cfg.messageText100K,
        `案件名：${sponsor || '(なし)'}`,
        `制作者：${creator || '(なし)'}`,
        `投稿日：${uploadDate}`,
        `投稿URL：${videoUrl}`,
      ].join('\n');

      try {
        notifySlack({ text });
      } catch (e) {
        Logger.log(`[NOTIFY FAIL] tier=10w key=${key}: ${e.message}`);
        continue;
      }

      SHEET_CACHE.notified100k.add(key);
      SHEET_CACHE.newNotifiedRecords.push({
        key,
        notified_at: nowStr,
        views,
        sponsored_by: sponsor,
        upload_date: uploadDate,
        tier: '10w',
      });
      notifiedCount++;
      continue;
    }

    // ========== 1万 tier(10万には満たない場合のみ) ==========
    if (SHEET_CACHE.notified10k.has(key)) continue;

    try {
      notifySlack({
        title: '一万再生突破しました！',
        message: '<!channel> ' + sponsor + cfg.messageText,
        level: 'info',
        fields: {
          '案件': sponsor || '(なし)',
          '制作者': creator || '(なし)',
          '投稿日': uploadDate,
          '動画URL': videoUrl,
        },
      });
    } catch (e) {
      Logger.log(`[NOTIFY FAIL] tier=1w key=${key}: ${e.message}`);
      continue;
    }

    SHEET_CACHE.notified10k.add(key);
    SHEET_CACHE.newNotifiedRecords.push({
      key,
      notified_at: nowStr,
      views,
      sponsored_by: sponsor,
      upload_date: uploadDate,
      tier: '1w',
    });
    notifiedCount++;
  }

  return notifiedCount;
};

// ============================================================================
// Slack 送信(Bot Token + chat.postMessage)
// ============================================================================

const SLACK_API_URL = 'https://slack.com/api/chat.postMessage';

/**
 * 汎用 Slack 通知関数。失敗時は throw する。
 * data = { title, message, level, fields }
 */
const notifySlack = (data) => {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('SLACK_BOT_TOKEN');
  const channel = props.getProperty('SLACK_CHANNEL_ID');

  if (!token || !channel) {
    throw new Error('SLACK_BOT_TOKEN or SLACK_CHANNEL_ID is not set');
  }

  const payload = {
    channel: channel,
    text: typeof data.text === 'string' ? data.text : buildText(data),
  };

  const response = UrlFetchApp.fetch(SLACK_API_URL, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const body = response.getContentText();
  const result = JSON.parse(body);

  if (response.getResponseCode() !== 200 || !result.ok) {
    throw new Error('Slack notification failed: ' + body);
  }

  Logger.log('[NOTIFY] Slack message sent: ts=' + result.ts);
};

/**
 * notifySlack 用のテキスト本文を組み立てる。
 */
const buildText = (data) => {
  const icon = levelIcon(data.level);
  const title = data.title || 'GAS通知';
  const message = data.message || '';
  const fields = data.fields || {};
  const lines = [icon + ' *' + title + '*'];
  if (message) lines.push(message);

  Object.keys(fields).forEach(function (key) {
    if (fields[key] !== undefined && fields[key] !== null && fields[key] !== '') {
      lines.push('*' + key + '*: ' + fields[key]);
    }
  });

  return lines.join('\n');
};

/**
 * level 文字列に対応するアイコンを返す。
 */
const levelIcon = (level) => {
  if (level === 'success') return ':white_check_mark:';
  if (level === 'warning') return ':warning:';
  if (level === 'error') return ':rotating_light:';
  return ':bell:';
};

// ============================================================================
// 手動運用関数(Slack設定確認・テスト送信)
// ============================================================================

/**
 * Slack関連 ScriptProperties の設定状況をログに表示する。
 */
function checkSlackNotificationProperties() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('SLACK_BOT_TOKEN');
  const channel = props.getProperty('SLACK_CHANNEL_ID');

  if (!token || !channel) {
    throw new Error('SLACK_BOT_TOKEN or SLACK_CHANNEL_ID is not set');
  }

  Logger.log('SLACK_BOT_TOKEN: ' + token.slice(0, 8) + '...');
  Logger.log('SLACK_CHANNEL_ID: ' + channel);
  Logger.log('SLACK_MESSAGE_TEXT: ' + (props.getProperty('SLACK_MESSAGE_TEXT') || '(未設定)'));
}

/**
 * Slack送信の動作確認用テスト関数。GASエディタから手動実行する。
 */
function sendSlackNotification() {
  notifySlack({
    title: 'GASからの通知',
    message: '特定の条件を検知しました。',
    level: 'info',
    fields: {
      source: 'Google Apps Script',
      detectedAt: new Date().toISOString(),
      status: '通知対象',
    },
  });
}

// ============================================================================
// ジョブ処理本体
// ============================================================================

/**
 * raw spreadsheet からデータを読み、月別に振り分け、整形して
 * 月別シートに直接書込 + キャッシュ(文字列)も更新する。
 */
const processDataJob = (job) => {
  const lap = (label, t) => Logger.log(`  [TIMING:job] ${label}: ${((Date.now() - t) / 1000).toFixed(2)}s`);

  const tRaw = Date.now();
  const rawData = readRawSpreadsheetData(job.rawFileId);
  lap(`readRawSpreadsheetData (rows=${rawData.rows.length})`, tRaw);

  if (!rawData.rows || rawData.rows.length === 0) {
    Logger.log(`[PROCESS] Raw data not readable (fileId=${job.rawFileId})`);
    return new Set();
  }

  const tGroup = Date.now();
  const grouped = groupRowsByMonth(rawData.headers, rawData.rows, []);
  lap(`groupRowsByMonth (months=${Object.keys(grouped).length})`, tGroup);

  const updatedMonths = new Set();
  const stagedByMonth = {};

  // メモリ上で全月の整形を完了させる(部分失敗で月別シートが壊れないよう)
  const tTrans = Date.now();
  for (const [monthName, monthData] of Object.entries(grouped)) {
    stagedByMonth[monthName] = applySheetTransformationsInMemory(rawData.headers, monthData.rows);
  }
  lap('applySheetTransformationsInMemory(all months)', tTrans);

  // 通知条件チェック + 通知実行(月別シート書込前)
  const tNotify = Date.now();
  let totalNotified = 0;
  for (const result of Object.values(stagedByMonth)) {
    totalNotified += checkAndNotifyRows(result.headers, result.rows);
  }
  lap(`checkAndNotifyRows (notified=${totalNotified})`, tNotify);

  // 月別シートに直接書込(typed) — キャッシュ経由ではない
  const tWrite = Date.now();
  for (const [monthName, result] of Object.entries(stagedByMonth)) {
    const sheet = getSheet(monthName);
    writeRowsPreserveTypes(sheet, result.headers, result.rows);
    updatedMonths.add(monthName);
  }
  lap(`write monthly sheets (count=${updatedMonths.size})`, tWrite);

  // キャッシュは display strings で更新(統合シート構築用)
  const tCache = Date.now();
  for (const [monthName, result] of Object.entries(stagedByMonth)) {
    SHEET_CACHE.sourceSheets[monthName] = {
      headers: result.headers.map(displayify),
      rows: result.rows.map((r) => r.map(displayify)),
    };
  }
  lap(`update cache (display strings, new rows only)`, tCache);

  return updatedMonths;
};

// ============================================================================
// エントリポイント: 時間トリガーから起動される
// ============================================================================

function processSheetData(e) {
  const t0 = Date.now();
  const lap = (label, t) => Logger.log(`[TIMING] ${label}: ${((Date.now() - t) / 1000).toFixed(2)}s`);

  cleanupFiringTrigger(e && e.triggerUid);

  const tDeq = Date.now();
  const jobs = dequeueAllJobs();
  lap('dequeueAllJobs', tDeq);

  if (jobs.length === 0) {
    Logger.log('[PROCESS] No pending jobs. Exit.');
    return;
  }

  Logger.log(`[PROCESS] Starting. ${jobs.length} job(s) in queue.`);

  const tInit = Date.now();
  initSheetCache();
  lap(`initSheetCache (sources=${Object.keys(SHEET_CACHE.sourceSheets).length})`, tInit);

  let hasProcessedData = false;
  for (const job of jobs) {
    if (!job) {
      Logger.log('[PROCESS WARN] Null/undefined job entry, skipping.');
      continue;
    }
    const tJob = Date.now();
    try {
      if (job.type === 'process_data') {
        processDataJob(job);
        hasProcessedData = true;
        lap(`processDataJob runId=${job.runId}`, tJob);
      } else {
        Logger.log(`[PROCESS] Skipping unknown job type: ${job.type}`);
      }
    } catch (err) {
      Logger.log(`[PROCESS ERROR] runId=${(job && job.runId) || ''} type=${(job && job.type) || ''}: ${err.message}\n${err.stack}`);
    }
  }

  if (hasProcessedData) {
    const tNotifyFlush = Date.now();
    try {
      flushNewNotifiedRecords();
      lap(`flushNewNotifiedRecords (count=${SHEET_CACHE.newNotifiedRecords.length})`, tNotifyFlush);
    } catch (notifyErr) {
      Logger.log(`[NOTIFIED FLUSH ERROR] ${notifyErr.message}`);
    }

    const tCons = Date.now();
    try {
      rebuildConsolidatedFromCache();
      lap('rebuildConsolidatedFromCache', tCons);
    } catch (consErr) {
      Logger.log(`[CONSOLIDATED ERROR] ${consErr.message}`);
    }

    const tSync = Date.now();
    try {
      syncToLookerStudio();
      lap('syncToLookerStudio', tSync);
    } catch (syncErr) {
      Logger.log(`[SYNC ERROR] ${syncErr.message}`);
    }
  }

  lap('TOTAL processSheetData', t0);
  Logger.log(`[PROCESS] Done. processed=${jobs.length}`);
}

// ============================================================================
// 手動運用関数
// ============================================================================

function clearPendingJobs() {
  PropertiesService.getScriptProperties().deleteProperty(PENDING_JOBS_KEY);
  Logger.log('[MANUAL] Pending job queue cleared.');
}

function runProcessSheetDataNow() {
  processSheetData({});
}

function inspectProcessState() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === PROCESS_FN_NAME);
  Logger.log(`[INSPECT] processSheetData triggers: ${triggers.length}`);
  triggers.forEach((t) => {
    Logger.log(`  - uid=${t.getUniqueId()} type=${t.getEventType()}`);
  });

  const queue = JSON.parse(PropertiesService.getScriptProperties().getProperty(PENDING_JOBS_KEY) || '[]');
  Logger.log(`[INSPECT] Pending jobs: ${queue.length}`);
  queue.forEach((j) => {
    Logger.log(`  - type=${j.type} runId=${j.runId || ''} rawFileId=${j.rawFileId || ''}`);
  });
}

function clearAllProcessTriggers() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === PROCESS_FN_NAME);
  triggers.forEach((t) => ScriptApp.deleteTrigger(t));
  Logger.log(`[MANUAL] Deleted ${triggers.length} processSheetData trigger(s).`);
}
