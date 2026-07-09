/**
 * GreenArt SW SC_V1 schedule reminder
 *
 * Expected sheet columns on "통합 타임라인":
 * A No, B 차시, C 일자, D D-Day, E 상태, F 구분, G 유형, H 일정,
 * I 세부내용, J 중요도, K 원본
 *
 * Recommended extra columns:
 * L 담당자메일, M 참조메일, N 발송기준
 *
 * N 발송기준 example: D-14,D-7,D-1,D-day
 */

const CONFIG = {
  timezone: "Asia/Seoul",
  scheduleSheetName: "통합 타임라인",
  logSheetName: "메일발송이력",
  defaultOffsets: [14, 7, 1, 0],
  allowedPriority: ["필수", "중요"],
  fallbackRecipients: ["your-team@example.com"],
  senderName: "그린수원 통합 스케줄 알림"
};

function sendDailyScheduleReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName(CONFIG.scheduleSheetName);
  if (!scheduleSheet) {
    throw new Error(`시트를 찾을 수 없습니다: ${CONFIG.scheduleSheetName}`);
  }

  const logSheet = getOrCreateLogSheet_(ss);
  const sentKeys = getSentKeys_(logSheet);
  const today = toDateOnly_(new Date());
  const values = scheduleSheet.getDataRange().getValues();
  const rows = values.slice(1);
  const pendingLogs = [];

  rows.forEach((row, index) => {
    const item = parseScheduleRow_(row, index + 2);
    if (!item || !item.date) return;
    if (!CONFIG.allowedPriority.includes(item.priority)) return;

    const daysLeft = diffDays_(today, item.date);
    const offsets = item.offsets.length ? item.offsets : CONFIG.defaultOffsets;
    if (!offsets.includes(daysLeft)) return;

    const recipients = item.to.length ? item.to : CONFIG.fallbackRecipients;
    const cc = item.cc.join(",");
    const stage = daysLeft === 0 ? "D-day" : `D-${daysLeft}`;

    recipients.forEach((recipient) => {
      const key = makeLogKey_(item, stage, recipient);
      if (sentKeys.has(key)) return;

      MailApp.sendEmail({
        to: recipient,
        cc,
        name: CONFIG.senderName,
        subject: `[${stage}] ${item.course} ${item.title}`,
        htmlBody: buildMailBody_(item, stage)
      });

      pendingLogs.push([
        new Date(),
        key,
        item.course,
        formatDate_(item.date),
        stage,
        item.type,
        item.title,
        recipient,
        cc,
        item.rowNumber
      ]);
      sentKeys.add(key);
    });
  });

  if (pendingLogs.length) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, pendingLogs.length, pendingLogs[0].length).setValues(pendingLogs);
  }
}

function createDailyReminderTrigger() {
  ScriptApp.newTrigger("sendDailyScheduleReminders")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .nearMinute(30)
    .create();
}

function getOrCreateLogSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.logSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.logSheetName);
    sheet.appendRow(["발송시각", "키", "과정", "일자", "기준", "유형", "일정", "수신자", "참조", "원본행"]);
  }
  return sheet;
}

function getSentKeys_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();
  return new Set(sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat().filter(String));
}

function parseScheduleRow_(row, rowNumber) {
  const rawDate = row[2];
  const date = normalizeDate_(rawDate);
  if (!date) return null;

  return {
    rowNumber,
    course: String(row[1] || "").trim(),
    date,
    group: String(row[5] || "").trim(),
    type: String(row[6] || "").trim(),
    title: String(row[7] || "").trim(),
    detail: String(row[8] || "").trim(),
    priority: String(row[9] || "").trim(),
    source: String(row[10] || "").trim(),
    to: splitMails_(row[11]),
    cc: splitMails_(row[12]),
    offsets: parseOffsets_(row[13])
  };
}

function normalizeDate_(value) {
  if (value instanceof Date) return toDateOnly_(value);
  const text = String(value || "").trim();
  const match = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (!match) return null;
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function splitMails_(value) {
  return String(value || "")
    .split(/[,\n;]/)
    .map((mail) => mail.trim())
    .filter(Boolean);
}

function parseOffsets_(value) {
  const text = String(value || "").trim();
  if (!text) return [];
  return text.split(",")
    .map((item) => item.trim())
    .map((item) => item === "D-day" ? 0 : Number(item.replace(/^D-/, "")))
    .filter((item) => Number.isFinite(item));
}

function diffDays_(fromDate, targetDate) {
  return Math.round((toDateOnly_(targetDate) - toDateOnly_(fromDate)) / 86400000);
}

function toDateOnly_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatDate_(date) {
  return Utilities.formatDate(date, CONFIG.timezone, "yyyy-MM-dd");
}

function makeLogKey_(item, stage, recipient) {
  return [item.course, formatDate_(item.date), item.title, stage, recipient].join("|");
}

function buildMailBody_(item, stage) {
  return `
    <div style="font-family:Arial,'Noto Sans KR',sans-serif;line-height:1.6;color:#172033">
      <h2 style="margin:0 0 12px">[${stage}] ${item.title}</h2>
      <p><b>과정</b>: ${item.course}</p>
      <p><b>일자</b>: ${formatDate_(item.date)}</p>
      <p><b>구분</b>: ${item.group} / ${item.type}</p>
      <p><b>중요도</b>: ${item.priority}</p>
      <p><b>세부내용</b>: ${item.detail || "-"}</p>
      <p><b>원본</b>: ${item.source || "-"}</p>
      <hr>
      <p style="color:#667085;font-size:12px">이 메일은 SC_V1 통합 스케줄 리마인드 기준으로 자동 생성되었습니다.</p>
    </div>
  `;
}
