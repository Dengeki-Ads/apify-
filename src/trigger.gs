/**
 * trigger.gs — 定時起動 → Apify Actor起動
 */

/**
 * 基準日からoffsetMonth分ずらした月の期間を生成する。
 * 月初（1日）〜翌月1日の範囲を返す。ただし終了日が未来になる場合は当日を上限とする。
 *   - 当月: 1日 〜 当日（例: 6/12実行 → 6/1〜6/12）
 *   - 過去月: 1日 〜 翌月1日（満了。例: 6/12実行の先月 → 5/1〜6/1）
 * Apifyの仕様: since=期間終了日, until=期間開始日
 */
const buildMonthPeriod = (baseDate, offsetMonth) => {
  const tz = 'Asia/Tokyo';
  const year = baseDate.getFullYear();
  const month = baseDate.getMonth() + offsetMonth;

  const firstOfMonth = new Date(year, month, 1);
  const firstOfNextMonth = new Date(year, month + 1, 1); // 次の月の1日
  // 終了日は「次の月の1日」。ただし未来日にはせず、当日(baseDate)を上限にする。
  const endDate = firstOfNextMonth.getTime() > baseDate.getTime() ? baseDate : firstOfNextMonth;

  const until = Utilities.formatDate(firstOfMonth, tz, 'yyyy-MM-dd');
  const since = Utilities.formatDate(endDate, tz, 'yyyy-MM-dd');
  const monthLabel = Utilities.formatDate(firstOfMonth, tz, 'yyyy-MM');

  return { since, until, month: monthLabel };
};

/**
 * 当月+先月の2つの Apify Task を起動。
 * 8時と21時のトリガー、及び手動実行用。
 *
 * 日付による分岐:
 *   - 21日以前: 当月+先月の両方を起動（従来通り）
 *   - 22日以降: 当月のみ起動（runDailyJobCurrentOnly相当）
 */
function runDailyJob() {
  const baseDate = new Date();

  // 22日以降は当月のみに切り替える
  if (baseDate.getDate() >= 22) {
    runDailyJobCurrentOnly();
    return;
  }

  startApifyTask(buildMonthPeriod(baseDate, 0));   // 当月
  startApifyTask(buildMonthPeriod(baseDate, -1));  // 先月
}

/**
 * 当月のみの Apify Task を起動。
 * 14時のトリガー用。
 */
function runDailyJobCurrentOnly() {
  const baseDate = new Date();
  startApifyTask(buildMonthPeriod(baseDate, 0));   // 当月
}

/**
 * 対象期間でApify Taskを起動する。
 */
const startApifyTask = (period) => {
  let runId = null;

  try {
    const webhookUrl = getConfig('GAS_WEBAPP_URL');

    const webhooks = [{
      eventTypes: [
        'ACTOR.RUN.SUCCEEDED',
        'ACTOR.RUN.FAILED',
        'ACTOR.RUN.TIMED_OUT',
        'ACTOR.RUN.ABORTED',
      ],
      requestUrl: webhookUrl,
    }];
    const webhooksParam = Utilities.base64Encode(JSON.stringify(webhooks));
    const url = `${buildTaskRunUrl()}&webhooks=${encodeURIComponent(webhooksParam)}`;

    // startUrlsはTaskに保存済みのものが使われる。期間だけ上書き送信。
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        since: period.since,
        until: period.until,
      }),
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    const body = JSON.parse(response.getContentText());

    if (statusCode !== 201) {
      throw new Error(`Apify API error (HTTP ${statusCode}): ${JSON.stringify(body)}`);
    }

    runId = body.data.id;

    Logger.log(`Actor run started. RunID: ${runId}, month: ${period.month}, period: ${period.until} - ${period.since}`);

  } catch (e) {
    Logger.log(`startApifyTask failed: month=${period.month}, error=${e.message}`);
  }
};

/**
 * 毎日のタイマートリガーを作成する。手動で1回実行。
 */
function setupDailyTrigger() {
  // 既存トリガーを全削除
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'runDailyJob' || t.getHandlerFunction() === 'runDailyJobCurrentOnly')
    .forEach((t) => ScriptApp.deleteTrigger(t));

  // 8時: 両方
  ScriptApp.newTrigger('runDailyJob')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  // 14時: 当月のみ
  ScriptApp.newTrigger('runDailyJobCurrentOnly')
    .timeBased()
    .everyDays(1)
    .atHour(14)
    .create();
  // 21時: 両方
  ScriptApp.newTrigger('runDailyJob')
    .timeBased()
    .everyDays(1)
    .atHour(21)
    .create();

  Logger.log('Daily triggers created: runDailyJob@8h, runDailyJobCurrentOnly@14h, runDailyJob@21h.');
}

/**
 * runDailyJob/runDailyJobCurrentOnlyのトリガーを削除する。
 */
function deleteDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'runDailyJob' || t.getHandlerFunction() === 'runDailyJobCurrentOnly');
  triggers.forEach((t) => ScriptApp.deleteTrigger(t));
  Logger.log(`Deleted ${triggers.length} trigger(s) for runDailyJob/runDailyJobCurrentOnly.`);
}
