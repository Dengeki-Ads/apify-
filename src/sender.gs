/**
 * sender.gs — 通知送信（Email / Webhook 切り替え対応）
 */

/**
 * 集計結果を通知する。NOTIFY_METHOD に応じて送信先を切り替え。
 * 未設定または "email" の場合は Gmail、"webhook" の場合は Webhook 送信。
 */
const sendNotification = (summary, sheetName) => {
  const method = PropertiesService.getScriptProperties().getProperty('NOTIFY_METHOD') || 'email';

  if (method === 'webhook') {
    sendWebhook(summary, sheetName);
  } else {
    sendSummaryEmail(summary, sheetName);
  }
};

/**
 * 集計結果を Gmail で送信する。
 */
const sendSummaryEmail = (summary, sheetName) => {
  const email = getConfig('NOTIFY_EMAIL');
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  const subject = `【views集計レポート】${sheetName} - ${now}`;

  let body = `集計期間：${sheetName}\n`;
  body += `実行日時：${now}\n\n`;
  body += `■ sponsored_by 別 views 合計\n\n`;

  let total = 0;
  for (const [sponsor, views] of Object.entries(summary)) {
    body += `${sponsor} : ${views.toLocaleString()}\n`;
    total += views;
  }

  body += `\n■ 合計\n`;
  body += `         ${total.toLocaleString()}\n`;

  GmailApp.sendEmail(email, subject, body);
  Logger.log(`[SENDER] Email sent to ${email}`);
};

/**
 * 集計結果を Webhook で送信する（将来実装用スタブ）。
 */
const sendWebhook = (summary, sheetName) => {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('WEBHOOK_URL');
  if (!webhookUrl) {
    Logger.log('[SENDER WARN] WEBHOOK_URL is not set. Falling back to email.');
    sendSummaryEmail(summary, sheetName);
    return;
  }

  const payload = {
    sheetName: sheetName,
    timestamp: new Date().toISOString(),
    summary: summary,
  };

  UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  Logger.log(`[SENDER] Webhook sent to ${webhookUrl}`);
};
