/**
 * trigger.gs — 定時起動 → Apify Actor起動
 */

/**
 * 実行日から対象期間リストを返す。
 * 当月は毎日、先月は21日のみ。
 */
const determineTargetPeriods = (baseDate) => {
  const periods = [];
  const tz = 'Asia/Tokyo';
  const now = new Date(baseDate);

  // 当月
  periods.push(buildMonthPeriod(now, 0));

  // 毎月21日は先月も対象
  const day = parseInt(Utilities.formatDate(now, tz, 'dd'), 10);
  if (day === 21) {
    periods.push(buildMonthPeriod(now, -1));
  }

  return periods;
};

/**
 * 基準日からoffsetMonth分ずらした月の期間を生成する。
 * 月初（1日）〜月末（最終日）の範囲を返す。
 * Apifyの仕様: since=期間終了日, until=期間開始日
 */
const buildMonthPeriod = (baseDate, offsetMonth) => {
  const tz = 'Asia/Tokyo';
  const year = baseDate.getFullYear();
  const month = baseDate.getMonth() + offsetMonth;

  const firstOfMonth = new Date(year, month, 1);
  const lastOfMonth = new Date(year, month + 1, 0); // 月末日

  const until = Utilities.formatDate(firstOfMonth, tz, 'yyyy-MM-dd');
  const since = Utilities.formatDate(lastOfMonth, tz, 'yyyy-MM-dd');
  const monthLabel = Utilities.formatDate(firstOfMonth, tz, 'yyyy-MM');

  return { since, until, month: monthLabel };
};

/**
 * 月ラベルからシート名を返す。
 */
const getMonthSheetName = (monthLabel) => {
  return monthLabel;
};

/**
 * メイン関数。定時トリガーまたは手動で実行する。
 * 対象月ごとにApify Taskを起動する。
 */
function runDailyJob() {
  const baseDate = new Date();
  const periods = determineTargetPeriods(baseDate);

  periods.forEach((period) => {
    startApifyTask(period);
  });
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

    appendLogRow({
      triggered_at: new Date(),
      run_id: runId,
      target_period: `${period.until} ~ ${period.since}`,
      target_month: period.month,
      status: '実行中',
    });

    PropertiesService.getScriptProperties().setProperty('CURRENT_PERIOD', `${period.until}_${period.since}`);

    Logger.log(`Actor run started. RunID: ${runId}, month: ${period.month}, period: ${period.until} - ${period.since}`);

  } catch (e) {
    Logger.log(`startApifyTask failed: ${e.message}`);

    appendLogRow({
      triggered_at: new Date(),
      run_id: runId || '',
      target_period: `${period.until} ~ ${period.since}`,
      target_month: period.month,
      status: '失敗',
      error_detail: e.message,
    });

    sendErrorNotification(
      'TikTok Scraper: Actor起動失敗',
      `startApifyTask() でエラーが発生しました。\nMonth: ${period.month}\n\n${e.message}`
    );
  }
};

/**
 * 毎日のタイマートリガーを作成する。手動で1回実行。
 */
function setupDailyTrigger() {
  const hour = parseInt(getConfig('TRIGGER_HOUR'), 10);

  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'runDailyJob')
    .forEach((t) => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('runDailyJob')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();

  Logger.log(`Daily trigger created for runDailyJob at ${hour}:00-${hour + 1}:00.`);
}

/**
 * runDailyJobのトリガーを削除する。
 */
function deleteDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'runDailyJob');

  triggers.forEach((t) => ScriptApp.deleteTrigger(t));
  Logger.log(`Deleted ${triggers.length} trigger(s) for runDailyJob.`);
}
