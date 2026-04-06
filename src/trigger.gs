/**
 * trigger.gs — 定時起動 → Apify Actor起動
 */

/**
 * メイン関数。定時トリガーまたは手動で実行する。
 */
function runDailyJob() {
  let runId = null;

  try {
    const webhookUrl = getConfig('GAS_WEBAPP_URL');

    // 実行月の期間を自動計算
    // Apifyの仕様: since=期間終了日, until=期間開始日
    const now = new Date();
    const firstOfMonth = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth(), 1), 'Asia/Tokyo', 'yyyy-MM-dd');
    const firstOfNextMonth = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() + 1, 1), 'Asia/Tokyo', 'yyyy-MM-dd');

    // Task APIを使用（期間のみオーバーライド）
    // webhooksはBase64エンコードしてクエリパラメータで渡す
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

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        since: firstOfNextMonth,
        until: firstOfMonth,
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
      status: '実行中',
    });

    // 今回の期間をScript Propertiesに保存（checkAndArchiveで参照）
    PropertiesService.getScriptProperties().setProperty('CURRENT_PERIOD', `${firstOfNextMonth}_${firstOfMonth}`);

    Logger.log(`Actor run started. RunID: ${runId}, period: ${firstOfMonth} - ${firstOfNextMonth}`);

  } catch (e) {
    Logger.log(`runDailyJob failed: ${e.message}`);

    appendLogRow({
      triggered_at: new Date(),
      run_id: runId || '',
      status: '失敗',
      error_detail: e.message,
    });

    sendErrorNotification(
      'TikTok Scraper: Actor起動失敗',
      `runDailyJob() でエラーが発生しました。\n\n${e.message}`
    );
  }
}

/**
 * Script PropertiesのSTART_URLSを使ってApify Actorを起動する。
 * START_URLS: カンマ区切りのTikTok URL（例: "https://www.tiktok.com/@user1,https://www.tiktok.com/@user2"）
 * 期間は自動計算。手動で1回実行する想定。
 */
function runWithCustomUrls() {
  const raw = getConfig('START_URLS');
  const startUrls = raw.split(',').map((s) => s.trim()).filter((s) => s);

  if (startUrls.length === 0) {
    throw new Error('START_URLS is empty.');
  }

  const webhookUrl = getConfig('GAS_WEBAPP_URL');
  const now = new Date();
  const firstOfMonth = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth(), 1), 'Asia/Tokyo', 'yyyy-MM-dd');
  const firstOfNextMonth = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() + 1, 1), 'Asia/Tokyo', 'yyyy-MM-dd');

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

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      since: firstOfNextMonth,
      until: firstOfMonth,
      startUrls: startUrls,
    }),
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  const body = JSON.parse(response.getContentText());

  if (statusCode !== 201) {
    throw new Error(`Apify API error (HTTP ${statusCode}): ${JSON.stringify(body)}`);
  }

  Logger.log(`Actor run started with ${startUrls.length} custom URLs. RunID: ${body.data.id}`);
}

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
