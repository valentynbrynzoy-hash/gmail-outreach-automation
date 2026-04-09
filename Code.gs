// ============================================================
//  НАСТРОЙКИ — меняй только этот блок
// ============================================================

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
const MY_ACCOUNT     = 'your_account@gmail.com';

// Строки аккаунтов в листе 📋 Меню (раздел аккаунтов, начиная с 6)
const ACCOUNT_ROW = {
  'account1@gmail.com': 6,
  'account2@gmail.com': 7,
  'account3@gmail.com': 8,
  'account4@gmail.com': 9,
  'account5@gmail.com': 10,
};

const ACCOUNT_BATCH_CELLS = Object.fromEntries(
  Object.entries(ACCOUNT_ROW).map(([k, r]) => [k, `G${r}`])
);

const ACCOUNT_FU_BATCH_CELLS = Object.fromEntries(
  Object.entries(ACCOUNT_ROW).map(([k, r]) => [k, `H${r}`])
);

const ACCOUNT_STATUS_CELLS = Object.fromEntries(
  Object.entries(ACCOUNT_ROW).map(([k, r]) => [k, `E${r}`])
);

// ============================================================
//  ЛИСТЫ
// ============================================================
const SHEET_STATUS    = '📋 Меню';
const CELL_STATUS     = ACCOUNT_STATUS_CELLS[MY_ACCOUNT] || 'E6';
const SHEET_DATA      = '📮 Почты';
const SHEET_TEMPLATES = '📧 Шаблоны';
const SHEET_FOLLOWUP  = '🔁 Фолоу-апы';
const SHEET_LOG       = '📊 Логи';

// ============================================================
//  НАСТРОЙКИ
// ============================================================
const SETTINGS_START_ROW = 19;
const SETTINGS_END_ROW   = 32;

// ============================================================
//  КОЛОНКИ ЛИСТА 📮 Почты
// ============================================================
const COL_NAME        = 1;
const COL_PLATFORM    = 2;
const COL_GEO         = 3;
const COL_STATUS      = 4;
const COL_DATE        = 5;
const COL_ACCOUNT     = 6;
const COL_TIME_LOC    = 7;
const COL_TIME_UKR    = 8;
const COL_EMAIL       = 9;
const COL_TEMPLATE    = 10;
const COL_REPLY       = 11;
const COL_REPLY_DATE  = 12;
const COL_FU_NUM      = 13;
const COL_FU_STATUS   = 14;
const COL_FU_NEXT     = 15;
const COL_MODE        = 16;

// ============================================================
//  КОЛОНКИ ЛИСТА 📧 Шаблоны (раздел отправителей)
// ============================================================
const SENDERS_START_ROW = 3;
const SENDERS_END_ROW   = 12;
const SCOL_EMAIL    = 1;
const SCOL_NAME     = 2;
const SCOL_SURNAME  = 3;
const SCOL_TITLE    = 4;
const SCOL_SIGN     = 5;
const SCOL_WHATSAPP = 6;
const SCOL_ACTIVE   = 7;

// ============================================================
//  КОЛОНКИ ЛИСТА 📧 Шаблоны (шаблоны писем)
// ============================================================
const TEMPLATES_START_ROW = 14;
const TCOL_ID      = 1;
const TCOL_STYLE   = 2;
const TCOL_SUBJECT = 3;
const TCOL_BODY    = 4;
const TCOL_ACTIVE  = 6;

// ============================================================
//  КОЛОНКИ ЛИСТА 🔁 Фолоу-апы
// ============================================================
const FOLLOWUP_START_ROW = 4;
const FCOL_ID      = 1;
const FCOL_NUM     = 2;
const FCOL_STYLE   = 3;
const FCOL_SUBJECT = 4;
const FCOL_BODY    = 5;
const FCOL_ACTIVE  = 6;

// ============================================================
//  КОЛОНКИ ЛИСТА 📊 Логи
// ============================================================
const LOG_START_ROW  = 3;
const LCOL_DATE      = 1;
const LCOL_TYPE      = 2;
const LCOL_ACCOUNT   = 3;
const LCOL_RECIPIENT = 4;
const LCOL_BLOGGER   = 5;
const LCOL_TPL_ID    = 6;
const LCOL_STYLE     = 7;
const LCOL_SUBJECT   = 8;
const LCOL_RESULT    = 9;
const LCOL_NOTES     = 10;

// ============================================================
//  ГЕО-СМЕЩЕНИЯ UTC
// ============================================================
const GEO_OFFSETS = {
  'ПЕРУ':      -5,
  'ЧИЛИ':      -3,
  'АРГЕНТИНА': -3,
  'КОЛУМБИЯ':  -5,
  'МЕКСИКА':   -6,
  'БРАЗИЛИЯ':  -3,
  'УКРАИНА':    3,
  'ТАИЛАНД':    7,
  'МОНГОЛИЯ':   8,
  'ПАКИСТАН':   5,
  'ШРИ-ЛАНКА':  5,
};

const PUBLIC_DOMAINS = [
  'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com',
  'icloud.com', 'mail.ru', 'yandex.ru', 'yandex.com',
  'live.com', 'msn.com', 'protonmail.com', 'aol.com',
];

// ============================================================
//  УМНЫЙ РАНДОМАЙЗЕР — настройки
// ============================================================
const FU_MIN_GAP_MINUTES = 3;
const FU_WINDOW_MINUTES  = 45;

// ============================================================
//  ЛИМИТ GMAIL API CALLS / ЗАПУСК
// ============================================================
const GMAIL_CALL_LIMIT = 200;
let gmailCallsUsed = 0;

function canSearch() { return gmailCallsUsed < GMAIL_CALL_LIMIT; }

function gmailSearch(query, start, max) {
  if (!canSearch()) return [];
  gmailCallsUsed++;
  return GmailApp.search(query, start, max);
}

// ============================================================
//  БАТЧ-ПОЗИЦИИ
// ============================================================
function getBatchCell()   { return ACCOUNT_BATCH_CELLS[MY_ACCOUNT]    || 'G6'; }
function getFuBatchCell() { return ACCOUNT_FU_BATCH_CELLS[MY_ACCOUNT] || 'H6'; }

function getBatchPos(ss, cell) {
  const sheet = ss.getSheetByName(SHEET_STATUS);
  if (!sheet) return 0;
  const val = sheet.getRange(cell).getValue();
  const pos = parseInt(val, 10);
  return isNaN(pos) || pos < 0 ? 0 : pos;
}

function setBatchPos(ss, cell, pos) {
  const sheet = ss.getSheetByName(SHEET_STATUS);
  if (sheet) sheet.getRange(cell).setValue(pos);
}

function resetBatchPos(ss, cell) { setBatchPos(ss, cell, 0); }

// ============================================================
//  ЗАГРУЗКА НАСТРОЕК
// ============================================================
function loadSettings(ss) {
  const sheet = ss.getSheetByName(SHEET_STATUS);
  if (!sheet) return {};
  const cfg  = {};
  const data = sheet.getDataRange().getValues();
  for (let i = SETTINGS_START_ROW - 1; i <= SETTINGS_END_ROW - 1; i++) {
    if (i >= data.length) break;
    const key = String(data[i][0]).trim();
    const val = String(data[i][1]).trim();
    if (key && val) cfg[key] = val;
  }
  return cfg;
}

function cfgBool(cfg, key, def = true) {
  if (!(key in cfg)) return def;
  return cfg[key].toUpperCase() === 'TRUE';
}

function cfgInt(cfg, key, def = 0) {
  const v = parseInt(cfg[key], 10);
  return isNaN(v) ? def : v;
}

function getFuInterval(cfg, fuNum) {
  if (fuNum === 1) return cfgInt(cfg, 'FOLLOWUP_INTERVAL_1', 1);
  if (fuNum === 2) return cfgInt(cfg, 'FOLLOWUP_INTERVAL_2', 5);
  return                    cfgInt(cfg, 'FOLLOWUP_INTERVAL_3', 7);
}

// ============================================================
//  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ============================================================
function formatUkrTime(utcDate) {
  try {
    const parts = new Intl.DateTimeFormat('uk-UA', {
      timeZone: 'Europe/Kiev',
      year:     'numeric',
      month:    '2-digit',
      day:      '2-digit',
      hour:     '2-digit',
      minute:   '2-digit',
      hour12:   false,
    }).formatToParts(utcDate);
    const p = {};
    parts.forEach(({ type, value }) => { p[type] = value; });
    return `${p.day}.${p.month}.${p.year} ${p.hour}:${p.minute}`;
  } catch (e) {
    const ukr = new Date(utcDate.getTime() + 3 * 60 * 60 * 1000);
    return [
      String(ukr.getUTCDate()).padStart(2, '0'),
      String(ukr.getUTCMonth() + 1).padStart(2, '0'),
      ukr.getUTCFullYear(),
    ].join('.') + ' ' + [
      String(ukr.getUTCHours()).padStart(2, '0'),
      String(ukr.getUTCMinutes()).padStart(2, '0'),
    ].join(':');
  }
}

function parseLogDate(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date) return dateVal;
  const str   = String(dateVal).trim();
  const match = str.match(/^(\d{2})\.(\d{2})\.(\d{4})(?:\s+(\d{2}):(\d{2}))?/);
  if (match) {
    const [, d, m, y, h = '0', mi = '0'] = match;
    return new Date(Date.UTC(+y, +m - 1, +d, +h - 3, +mi, 0));
  }
  return new Date(dateVal);
}

function parseEmails(raw) {
  return raw.split(/[\n,;]+/).map(e => e.trim()).filter(e => e.length > 0);
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function fillVars(text, bloggerName, sender, origSubject) {
  return String(text)
    .replace(/{{блогер}}/gi,          bloggerName      || '')
    .replace(/{{имя_отправителя}}/gi, sender.name      || '')
    .replace(/{{подпись}}/gi,         sender.signature || '')
    .replace(/{{ватсап}}/gi,          sender.whatsapp  || '')
    .replace(/{{тема_оригинала}}/gi,  origSubject      || '');
}

function getEmailDomain(email) {
  const parts = email.split('@');
  return parts.length === 2 ? parts[1].toLowerCase() : '';
}

function isPublicDomain(domain) {
  return PUBLIC_DOMAINS.includes(domain.toLowerCase());
}

function isAutoreply(msg) {
  const sub = msg.getSubject().toLowerCase();
  const bdy = msg.getPlainBody().toLowerCase().slice(0, 500);
  return [
    'auto-reply', 'out of office', 'automated', 'do not reply',
    'noreply', 'no-reply', 'автоответ', 'отсутствую',
  ].some(k => sub.includes(k) || bdy.includes(k));
}

// ============================================================
//  УМНЫЙ РАНДОМАЙЗЕР ВРЕМЕНИ ДЛЯ ФОЛОУ-АПОВ
// ============================================================
function getOccupiedSlotsForDay(dataRows, logRows, targetDayUTC, sessionSlots) {
  const occupied = sessionSlots ? [...sessionSlots] : [];

  const ty = targetDayUTC.getUTCFullYear();
  const tm = targetDayUTC.getUTCMonth();
  const td = targetDayUTC.getUTCDate();

  function sameDay(d) {
    return d && !isNaN(d)
      && d.getUTCFullYear() === ty
      && d.getUTCMonth()   === tm
      && d.getUTCDate()    === td;
  }

  function add(d) {
    if (sameDay(d)) occupied.push(d.getUTCHours() * 60 + d.getUTCMinutes());
  }

  for (let i = LOG_START_ROW - 1; i < logRows.length; i++) {
    const row  = logRows[i];
    const acc  = String(row[LCOL_ACCOUNT - 1]).trim();
    const type = String(row[LCOL_TYPE    - 1]).trim();
    if (acc !== MY_ACCOUNT) continue;
    if (type !== 'РАССЫЛКА' && type !== 'ФОЛОУ-АП') continue;
    add(parseLogDate(row[LCOL_DATE - 1]));
  }

  for (let i = 1; i < dataRows.length; i++) {
    const row = dataRows[i];
    if (String(row[COL_ACCOUNT - 1]).trim() !== MY_ACCOUNT) continue;
    add(parseLogDate(row[COL_FU_NEXT  - 1]));
    add(parseLogDate(row[COL_TIME_UKR - 1]));
  }

  return occupied;
}

function pickFreeSlot(baseMinutes, occupied) {
  const candidates = [];
  for (let delta = -FU_WINDOW_MINUTES; delta <= FU_WINDOW_MINUTES; delta++) {
    const slot = baseMinutes + delta;
    if (slot >= 0 && slot < 24 * 60) candidates.push(slot);
  }

  for (let i = candidates.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [candidates[i], candidates[j]] = [candidates[j], candidates[i]];
  }

  for (const slot of candidates) {
    if (occupied.every(occ => Math.abs(slot - occ) >= FU_MIN_GAP_MINUTES)) {
      return slot;
    }
  }

  Logger.log(`⚠️ pickFreeSlot: все слоты в окне заняты — берём рандом`);
  return candidates[Math.floor(Math.random() * candidates.length)];
}

function pickSmartFuTime(baseSentUTC, intervalDays, dataRows, logRows, sessionSlots) {
  const targetDayUTC = new Date(Date.UTC(
    baseSentUTC.getUTCFullYear(),
    baseSentUTC.getUTCMonth(),
    baseSentUTC.getUTCDate() + intervalDays,
    0, 0, 0, 0
  ));

  const _todayUTC = new Date();
  const _todayMidnight = new Date(Date.UTC(
    _todayUTC.getUTCFullYear(),
    _todayUTC.getUTCMonth(),
    _todayUTC.getUTCDate()
  ));
  if (targetDayUTC < _todayMidnight) {
    targetDayUTC.setUTCFullYear(_todayMidnight.getUTCFullYear());
    targetDayUTC.setUTCMonth(_todayMidnight.getUTCMonth());
    targetDayUTC.setUTCDate(_todayMidnight.getUTCDate() + 1);
    Logger.log(`⚠️ pickSmartFuTime: таргет в прошлом → снаппим на завтра`);
  }

  const dayKey      = `${targetDayUTC.getUTCFullYear()}-${targetDayUTC.getUTCMonth()}-${targetDayUTC.getUTCDate()}`;
  const sessForDay  = (sessionSlots && sessionSlots[dayKey]) ? sessionSlots[dayKey] : [];
  const baseMinutes = baseSentUTC.getUTCHours() * 60 + baseSentUTC.getUTCMinutes();

  const occupied   = getOccupiedSlotsForDay(dataRows, logRows, targetDayUTC, sessForDay);
  const chosenMins = pickFreeSlot(baseMinutes, occupied);

  if (sessionSlots) {
    if (!sessionSlots[dayKey]) sessionSlots[dayKey] = [];
    sessionSlots[dayKey].push(chosenMins);
  }

  const h  = Math.floor(chosenMins / 60);
  const mi = chosenMins % 60;

  const result = new Date(Date.UTC(
    targetDayUTC.getUTCFullYear(),
    targetDayUTC.getUTCMonth(),
    targetDayUTC.getUTCDate(),
    h, mi, 0, 0
  ));

  Logger.log(`📅 pickSmartFuTime: база=${baseMinutes}мин → выбрано ${String(h).padStart(2,'0')}:${String(mi).padStart(2,'0')} UTC | занято слотов: ${occupied.length}`);
  return result;
}

// ============================================================
//  ОТПРАВКА ФОЛОУ-АПА В ТРЕДЕ
// ============================================================
function sendReplyInThread(thread, toEmail, origSubject, htmlBody) {
  const subject = (origSubject && origSubject.startsWith('Re:'))
    ? origSubject
    : 'Re: ' + (origSubject || '');

  const messages = thread.getMessages();
  const lastMsg  = messages[messages.length - 1];
  let msgId      = '';
  try {
    const raw   = lastMsg.getRawContent();
    const match = raw.match(/^Message-ID:\s*(<[^>]+>)/im);
    if (match) msgId = match[1];
  } catch(e) {
    Logger.log(`⚠️ Не удалось получить Message-ID: ${e.message}`);
  }

  const headers = [
    `To: ${toEmail}`,
    `Subject: ${subject}`,
    `MIME-Version: 1.0`,
    `Content-Type: text/html; charset=UTF-8`,
  ];
  if (msgId) {
    headers.push(`In-Reply-To: ${msgId}`);
    headers.push(`References: ${msgId}`);
  }

  const emailRaw = [...headers, '', htmlBody].join('\r\n');
  const encoded  = Utilities.base64EncodeWebSafe(emailRaw, Utilities.Charset.UTF_8);

  Gmail.Users.Messages.send(
    { threadId: thread.getId(), raw: encoded },
    'me'
  );
}

// ============================================================
//  ЗАГРУЗКА ОТПРАВИТЕЛЕЙ И ШАБЛОНОВ
// ============================================================
function loadSender(sheet) {
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = SENDERS_START_ROW - 1; i <= SENDERS_END_ROW - 1; i++) {
    if (i >= data.length) break;
    const row    = data[i];
    const email  = String(row[SCOL_EMAIL  - 1]).trim();
    const active = String(row[SCOL_ACTIVE - 1]).trim();
    if (email === MY_ACCOUNT && (active === '✅' || active.toLowerCase() === 'true')) {
      return {
        email:     email,
        name:      String(row[SCOL_NAME     - 1]).trim(),
        surname:   String(row[SCOL_SURNAME  - 1]).trim(),
        title:     String(row[SCOL_TITLE    - 1]).trim(),
        signature: String(row[SCOL_SIGN     - 1]).trim(),
        whatsapp:  String(row[SCOL_WHATSAPP - 1]).trim(),
      };
    }
  }
  return null;
}

function loadActiveTemplates(sheet) {
  const data      = sheet.getDataRange().getValues();
  const templates = [];
  for (let i = TEMPLATES_START_ROW - 1; i < data.length; i++) {
    const row    = data[i];
    const id     = row[TCOL_ID     - 1];
    const style  = row[TCOL_STYLE  - 1];
    const subj   = row[TCOL_SUBJECT - 1];
    const body   = row[TCOL_BODY   - 1];
    const active = String(row[TCOL_ACTIVE - 1]).trim();
    if (id && subj && body && (active === '✅' || active.toLowerCase() === 'true')) {
      templates.push({ id: String(id), style: String(style), subject: String(subj), body: String(body) });
    }
  }
  return templates;
}

function getUsedTemplateIds(logSheet, daysUnique) {
  if (!logSheet) return [];
  const data    = logSheet.getDataRange().getValues();
  const usedIds = [];
  const cutoff  = new Date();
  cutoff.setDate(cutoff.getDate() - daysUnique);
  for (let i = LOG_START_ROW - 1; i < data.length; i++) {
    const row     = data[i];
    const account = String(row[LCOL_ACCOUNT - 1]).trim();
    const type    = String(row[LCOL_TYPE    - 1]).trim();
    const tplId   = String(row[LCOL_TPL_ID  - 1]).trim();
    if (account !== MY_ACCOUNT || type !== 'РАССЫЛКА') continue;
    const rowDate = parseLogDate(row[LCOL_DATE - 1]);
    if (!rowDate || isNaN(rowDate)) continue;
    if (rowDate >= cutoff && tplId && !usedIds.includes(tplId)) usedIds.push(tplId);
  }
  return usedIds;
}

function loadFollowUpTemplates(fuSheet) {
  if (!fuSheet) return [];
  const data = fuSheet.getDataRange().getValues();
  const tpls = [];
  for (let i = FOLLOWUP_START_ROW - 1; i < data.length; i++) {
    const row    = data[i];
    const id     = String(row[FCOL_ID     - 1]).trim();
    const num    = parseInt(row[FCOL_NUM  - 1], 10) || 1;
    const style  = String(row[FCOL_STYLE  - 1]).trim();
    const subj   = String(row[FCOL_SUBJECT - 1]).trim();
    const body   = String(row[FCOL_BODY   - 1]).trim();
    const active = String(row[FCOL_ACTIVE - 1]).trim();
    if (!id || !subj || !body) continue;
    if (active !== '✅' && active.toLowerCase() !== 'true') continue;
    tpls.push({ id, num, style, subject: subj, body });
  }
  return tpls;
}

function getUsedFollowUpIds(logSheet, daysUniqueFollowup) {
  if (!logSheet) return [];
  const data    = logSheet.getDataRange().getValues();
  const usedIds = [];
  const cutoff  = new Date();
  cutoff.setDate(cutoff.getDate() - daysUniqueFollowup);
  for (let i = LOG_START_ROW - 1; i < data.length; i++) {
    const row     = data[i];
    const account = String(row[LCOL_ACCOUNT - 1]).trim();
    const type    = String(row[LCOL_TYPE    - 1]).trim();
    const tplId   = String(row[LCOL_TPL_ID  - 1]).trim();
    if (account !== MY_ACCOUNT || type !== 'ФОЛОУ-АП') continue;
    const rowDate = parseLogDate(row[LCOL_DATE - 1]);
    if (!rowDate || isNaN(rowDate)) continue;
    if (rowDate >= cutoff && tplId && !usedIds.includes(tplId)) usedIds.push(tplId);
  }
  return usedIds;
}

function pickFollowUpTemplate(templates, targetNum, style, logUsedIds, sessionUsedIds) {
  const logUsed     = logUsedIds     || [];
  const sessionUsed = sessionUsedIds || [];

  const byNumStyle = templates.filter(
    t => t.num === targetNum && t.style.toLowerCase() === style.toLowerCase()
  );
  const pool = byNumStyle.length > 0
    ? byNumStyle
    : templates.filter(t => t.num === targetNum);

  if (pool.length === 0) return null;

  const fresh = pool.filter(t => !logUsed.includes(t.id) && !sessionUsed.includes(t.id));
  if (fresh.length > 0) return fresh[Math.floor(Math.random() * fresh.length)];

  const notSession = pool.filter(t => !sessionUsed.includes(t.id));
  if (notSession.length > 0) return notSession[Math.floor(Math.random() * notSession.length)];

  return pool[Math.floor(Math.random() * pool.length)];
}

// ============================================================
//  ЗАПИСЬ В ЛОГ
// ============================================================
function writeLog(logSheet, nowUTC, type, recipient, blogger, tplId, style, subject, result, notes) {
  if (!logSheet) return;
  logSheet.appendRow([
    formatUkrTime(nowUTC), type, MY_ACCOUNT,
    recipient, blogger || '—', tplId || '—',
    style || '—', subject || '—', result, notes || '',
  ]);
}

// ============================================================
//  ПОИСК ТРЕДА В SENT
// ============================================================
function findSentThreadWithSubject(emails, sentDateRaw) {
  for (const email of emails) {
    if (!isValidEmail(email)) continue;
    if (!canSearch()) break;
    try {
      const sentDate   = sentDateRaw instanceof Date ? sentDateRaw : new Date(sentDateRaw);
      const searchDate = new Date(sentDate.getTime() - 24 * 60 * 60 * 1000);
      const yyyy = searchDate.getFullYear();
      const mm   = String(searchDate.getMonth() + 1).padStart(2, '0');
      const dd   = String(searchDate.getDate()).padStart(2, '0');
      const query   = `to:(${email}) after:${yyyy}/${mm}/${dd} in:sent`;
      const threads = gmailSearch(query, 0, 5);
      if (!threads || threads.length === 0) continue;
      for (const thread of threads) {
        for (const msg of thread.getMessages()) {
          if (msg.getFrom().includes(MY_ACCOUNT)) {
            return { thread, subject: msg.getSubject() };
          }
        }
      }
    } catch (e) {
      Logger.log(`⚠️ findSentThread (${email}): ${e.message}`);
    }
  }
  return { thread: null, subject: null };
}

// ============================================================
//  БЫСТРАЯ ПРОВЕРКА ОТВЕТА
// ============================================================
function quickCheckReply(emails, sentDateRaw) {
  if (!canSearch()) return { found: false, thread: null, origSubject: null };
  const sentDate = sentDateRaw instanceof Date ? sentDateRaw : new Date(sentDateRaw);
  const { thread, subject: origSubject } = findSentThreadWithSubject(emails, sentDateRaw);
  if (!thread) return { found: false, thread: null, origSubject };
  for (const msg of thread.getMessages()) {
    if (isAutoreply(msg)) continue;
    if (msg.getDate() > sentDate && !msg.getFrom().includes(MY_ACCOUNT)) {
      return { found: true, date: formatUkrTime(msg.getDate()), thread, origSubject };
    }
  }
  return { found: false, thread, origSubject };
}

// ============================================================
//  ФУНКЦИЯ 1: ОСНОВНАЯ ХОЛОДНАЯ РАССЫЛКА
// ============================================================
function sendScheduledEmails() {
  const ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cfg        = loadSettings(ss);
  const daysUnique = cfgInt(cfg, 'DAYS_UNIQUE', 5);
  const fuEnabled  = cfgBool(cfg, 'FOLLOWUP_ENABLED', true);

  const dataSheet  = ss.getSheetByName(SHEET_DATA);
  const templSheet = ss.getSheetByName(SHEET_TEMPLATES);
  const logSheet   = ss.getSheetByName(SHEET_LOG);

  const sender = loadSender(templSheet);
  if (!sender) {
    Logger.log(`❌ Аккаунт ${MY_ACCOUNT} не найден в Шаблонах`);
    updateTriggerStatus(ss);
    return;
  }

  const allTemplates = loadActiveTemplates(templSheet);
  if (allTemplates.length === 0) {
    Logger.log('❌ Нет активных шаблонов');
    updateTriggerStatus(ss);
    return;
  }

  const usedIds = getUsedTemplateIds(logSheet, daysUnique);
  let availableTemplates = allTemplates.filter(t => !usedIds.includes(t.id));
  if (availableTemplates.length === 0) {
    Logger.log('⚠️ Все шаблоны использованы — сбрасываем цикл');
    availableTemplates = allTemplates;
  }

  const nowUTC       = new Date();
  const dataRows     = dataSheet.getDataRange().getValues();
  const logRows      = logSheet ? logSheet.getDataRange().getValues() : [];
  const sessionSlots = {};

  for (let i = 1; i < dataRows.length; i++) {
    const row      = dataRows[i];
    const account  = row[COL_ACCOUNT  - 1];
    const status   = row[COL_STATUS   - 1];
    const dateRaw  = row[COL_DATE     - 1];
    const timeRaw  = row[COL_TIME_LOC - 1];
    const emailRaw = row[COL_EMAIL    - 1];
    const name     = row[COL_NAME     - 1];
    const geo      = String(row[COL_GEO      - 1] || '').trim().toUpperCase();
    const styleRaw = String(row[COL_TEMPLATE - 1] || '').trim();
    const mode     = String(row[COL_MODE     - 1] || '').trim().toLowerCase();
    const fuStatus = String(row[COL_FU_STATUS - 1] || '').trim();

    if (account !== MY_ACCOUNT)     continue;
    if (status  === 'да')           continue;
    if (!dateRaw || !emailRaw)      continue;
    if (mode === 'только проверка') continue;
    if (fuStatus === 'bounce')      continue;

    if (!(geo in GEO_OFFSETS)) {
      Logger.log(`Строка ${i + 1}: неизвестное Гео — "${geo}"`);
      continue;
    }

    const offsetHours = GEO_OFFSETS[geo];
    let rowHour = 0, rowMinute = 0;
    if (timeRaw instanceof Date) {
      rowHour = timeRaw.getHours(); rowMinute = timeRaw.getMinutes();
    } else if (timeRaw) {
      const parts = String(timeRaw).replace(',', '.').replace(':', '.').split('.');
      rowHour   = parseInt(parts[0], 10) || 0;
      rowMinute = parts[1] ? parseInt(parts[1], 10) : 0;
    }

    const sendDate = new Date(dateRaw);
    const sendUTC  = new Date(Date.UTC(
      sendDate.getFullYear(), sendDate.getMonth(), sendDate.getDate(),
      rowHour - offsetHours, rowMinute, 0
    ));

    const nowDateOnly  = Date.UTC(nowUTC.getUTCFullYear(),  nowUTC.getUTCMonth(),  nowUTC.getUTCDate());
    const sendDateOnly = Date.UTC(sendUTC.getUTCFullYear(), sendUTC.getUTCMonth(), sendUTC.getUTCDate());

    if (nowDateOnly !== sendDateOnly) continue;
    if (nowUTC < sendUTC)            continue;

    let pool = availableTemplates;
    if (styleRaw) {
      const byStyle = availableTemplates.filter(t => t.style.toLowerCase() === styleRaw.toLowerCase());
      if (byStyle.length > 0) {
        pool = byStyle;
      } else {
        const allByStyle = allTemplates.filter(t => t.style.toLowerCase() === styleRaw.toLowerCase());
        pool = allByStyle.length > 0 ? allByStyle : availableTemplates;
        Logger.log(`Строка ${i + 1} (${name}): шаблоны стиля «${styleRaw}» исчерпаны — сброс`);
      }
    }
    const tpl = pool[Math.floor(Math.random() * pool.length)];

    const subject = fillVars(tpl.subject, name, sender, '');
    const body    = fillVars(tpl.body,    name, sender, '').replace(/\n/g, '<br>');
    const emails  = parseEmails(String(emailRaw));

    let sent = false;
    for (const email of emails) {
      if (!isValidEmail(email)) {
        Logger.log(`Невалидный email: ${email}`);
        continue;
      }
      try {
        GmailApp.sendEmail(email, subject, '', { htmlBody: body });
        Logger.log(`✅ РАССЫЛКА ${name} | ${email} | шаблон ${tpl.id} (${tpl.style}) | ${formatUkrTime(nowUTC)}`);
        writeLog(logSheet, nowUTC, 'РАССЫЛКА', email, name, tpl.id, tpl.style, subject, '✅ Отправлено', '');
        sent = true;
        availableTemplates = availableTemplates.filter(t => t.id !== tpl.id);
        if (availableTemplates.length === 0) availableTemplates = allTemplates;
      } catch (e) {
        Logger.log(`❌ Ошибка отправки ${email}: ${e.message}`);
        writeLog(logSheet, nowUTC, 'РАССЫЛКА', email, name, tpl.id, tpl.style, subject, `❌ ${e.message}`, '');
      }
    }

    if (sent) {
      dataSheet.getRange(i + 1, COL_STATUS).setValue('да');
      dataSheet.getRange(i + 1, COL_TIME_UKR).setValue(formatUkrTime(nowUTC));
      dataSheet.getRange(i + 1, COL_FU_NUM).setValue(0);
      if (!mode) dataSheet.getRange(i + 1, COL_MODE).setValue('полный');

      dataRows[i][COL_TIME_UKR - 1] = formatUkrTime(nowUTC);

      if (fuEnabled) {
        const interval1 = getFuInterval(cfg, 1);
        const fu1UTC = pickSmartFuTime(sendUTC, interval1, dataRows, logRows, sessionSlots);
        dataSheet.getRange(i + 1, COL_FU_STATUS).setValue('ожидание');
        dataSheet.getRange(i + 1, COL_FU_NEXT).setValue(formatUkrTime(fu1UTC));
        dataRows[i][COL_FU_NEXT - 1] = formatUkrTime(fu1UTC);
      }
    }
  }

  updateTriggerStatus(ss);
}

// ============================================================
//  ФУНКЦИЯ 2: ПРОВЕРКА ОТВЕТОВ
// ============================================================
function checkReplies() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cfg = loadSettings(ss);

  if (!cfgBool(cfg, 'CHECK_REPLIES_ENABLED')) {
    Logger.log('⏸ Проверка ответов отключена');
    return;
  }

  const maxDays   = cfgInt(cfg, 'CHECK_REPLIES_DAYS', 30);
  const batchSize = cfgInt(cfg, 'CHECK_BATCH_SIZE',   150);

  const dataSheet = ss.getSheetByName(SHEET_DATA);
  const logSheet  = ss.getSheetByName(SHEET_LOG);
  const nowUTC    = new Date();
  const rows      = dataSheet.getDataRange().getValues();

  const candidates = [];
  for (let i = 1; i < rows.length; i++) {
    const row         = rows[i];
    const account     = row[COL_ACCOUNT   - 1];
    const status      = row[COL_STATUS    - 1];
    const replyStatus = String(row[COL_REPLY     - 1] || '').trim().toLowerCase();
    const fuStatus    = String(row[COL_FU_STATUS  - 1] || '').trim();
    const mode        = String(row[COL_MODE       - 1] || '').trim().toLowerCase();
    const emailRaw    = row[COL_EMAIL     - 1];
    const dateSent    = row[COL_DATE      - 1];

    if (account !== MY_ACCOUNT)                               continue;
    if (mode !== 'полный' && mode !== 'только проверка')      continue;
    if (status      !== 'да')                                 continue;
    if (replyStatus === 'да')                                 continue;
    if (fuStatus    === 'ответил')                            continue;
    if (!emailRaw   || !dateSent)                             continue;

    const diffDays = (nowUTC - new Date(dateSent)) / (1000 * 60 * 60 * 24);
    if (diffDays > maxDays) continue;

    candidates.push({ rowIndex: i + 1, name: row[COL_NAME - 1], emailRaw: String(emailRaw), dateSent });
  }

  if (candidates.length === 0) {
    Logger.log(`ℹ️ checkReplies (${MY_ACCOUNT}): нет строк для проверки`);
    return;
  }

  const batchCell = getBatchCell();
  let batchStart  = getBatchPos(ss, batchCell);
  if (batchStart >= candidates.length) {
    resetBatchPos(ss, batchCell);
    batchStart = 0;
  }

  const batchEnd = Math.min(batchStart + batchSize, candidates.length);
  const batch    = candidates.slice(batchStart, batchEnd);
  Logger.log(`📋 checkReplies (${MY_ACCOUNT}): [${batchStart}–${batchEnd - 1}] из ${candidates.length}`);

  let checked = 0, found = 0;

  for (let b = 0; b < batch.length; b++) {
    if (!canSearch()) {
      Logger.log(`⚠️ Лимит Gmail API (${gmailCallsUsed}) — стоп на индексе ${batchStart + b}`);
      setBatchPos(ss, batchCell, batchStart + b);
      return;
    }

    const { rowIndex, name, emailRaw, dateSent } = batch[b];
    const emails    = parseEmails(emailRaw);
    const sentDate  = new Date(dateSent);
    const sd        = new Date(sentDate.getTime() - 24 * 60 * 60 * 1000);
    const afterDate = `${sd.getFullYear()}/${String(sd.getMonth() + 1).padStart(2, '0')}/${String(sd.getDate()).padStart(2, '0')}`;

    checked++;
    let replyFound = false;
    let replyDate  = '';

    const { thread, subject: origSubject } = findSentThreadWithSubject(emails, sentDate);
    if (thread) {
      for (const msg of thread.getMessages()) {
        if (isAutoreply(msg)) continue;
        if (msg.getDate() > sentDate && !msg.getFrom().includes(MY_ACCOUNT)) {
          replyFound = true;
          replyDate  = formatUkrTime(msg.getDate());
          break;
        }
      }
    }

    if (!replyFound && origSubject && canSearch()) {
      const reSub    = origSubject.startsWith('Re:') ? origSubject : `Re: ${origSubject}`;
      const threads2 = gmailSearch(`subject:("${reSub}") after:${afterDate} in:inbox`, 0, 5);
      outer2: for (const t of threads2) {
        for (const msg of t.getMessages()) {
          if (isAutoreply(msg)) continue;
          if (msg.getDate() > sentDate && !msg.getFrom().includes(MY_ACCOUNT)) {
            replyFound = true; replyDate = formatUkrTime(msg.getDate());
            break outer2;
          }
        }
      }
    }

    if (!replyFound && canSearch()) {
      for (const email of emails) {
        if (!isValidEmail(email)) continue;
        const domain = getEmailDomain(email);
        if (!domain || isPublicDomain(domain)) continue;
        const threads3 = gmailSearch(`from:(*@${domain}) after:${afterDate} in:inbox`, 0, 5);
        let hit = false;
        outer3: for (const t of threads3) {
          for (const msg of t.getMessages()) {
            if (isAutoreply(msg)) continue;
            if (msg.getDate() > sentDate && !msg.getFrom().includes(MY_ACCOUNT)) {
              replyFound = true; replyDate = formatUkrTime(msg.getDate());
              hit = true; break outer3;
            }
          }
        }
        if (hit) break;
      }
    }

    if (replyFound) {
      found++;
      dataSheet.getRange(rowIndex, COL_REPLY).setValue('да');
      dataSheet.getRange(rowIndex, COL_REPLY_DATE).setValue(replyDate);
      dataSheet.getRange(rowIndex, COL_FU_STATUS).setValue('ответил');
      dataSheet.getRange(rowIndex, COL_FU_NEXT).setValue('');
      writeLog(logSheet, nowUTC, 'ПРОВЕРКА', emailRaw, name, '—', '—', '—', '✅ Ответ получен', replyDate);
    }
  }

  const nextStart = (batchStart + batch.length) >= candidates.length ? 0 : batchStart + batch.length;
  setBatchPos(ss, batchCell, nextStart);
  Logger.log(`✅ checkReplies: проверено ${checked}, найдено ${found}, следующий: ${nextStart}`);
}

// ============================================================
//  ФУНКЦИЯ 3: ФОЛОУ-АПЫ
// ============================================================
function sendFollowUps() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cfg = loadSettings(ss);

  if (!cfgBool(cfg, 'FOLLOWUP_ENABLED')) {
    Logger.log('⏸ Фолоу-апы отключены');
    return;
  }

  const maxFU              = cfgInt(cfg,  'FOLLOWUP_COUNT',         3);
  const daysUniqueFollowup = cfgInt(cfg,  'DAYS_UNIQUE_FOLLOWUP',   5);
  const stopOnReply        = cfgBool(cfg, 'FOLLOWUP_STOP_ON_REPLY', true);

  const dataSheet  = ss.getSheetByName(SHEET_DATA);
  const fuSheet    = ss.getSheetByName(SHEET_FOLLOWUP);
  const templSheet = ss.getSheetByName(SHEET_TEMPLATES);
  const logSheet   = ss.getSheetByName(SHEET_LOG);

  const sender = loadSender(templSheet);
  if (!sender) { Logger.log(`❌ Отправитель ${MY_ACCOUNT} не найден`); return; }

  const fuTemplates    = loadFollowUpTemplates(fuSheet);
  const logUsedIds     = getUsedFollowUpIds(logSheet, daysUniqueFollowup);
  const sessionUsedIds = [];

  const nowUTC      = new Date();
  const nowDateOnly = Date.UTC(nowUTC.getUTCFullYear(), nowUTC.getUTCMonth(), nowUTC.getUTCDate());

  const dataRows     = dataSheet.getDataRange().getValues();
  const logRows      = logSheet ? logSheet.getDataRange().getValues() : [];
  const sessionSlots = {};

  const candidates = [];
  for (let i = 1; i < dataRows.length; i++) {
    const row         = dataRows[i];
    const account     = row[COL_ACCOUNT   - 1];
    const status      = row[COL_STATUS    - 1];
    const emailRaw    = row[COL_EMAIL     - 1];
    const replyStatus = String(row[COL_REPLY     - 1] || '').trim().toLowerCase();
    const fuStatus    = String(row[COL_FU_STATUS  - 1] || '').trim();
    const fuNum       = parseInt(row[COL_FU_NUM   - 1], 10) || 0;
    const fuNextRaw   = row[COL_FU_NEXT   - 1];
    const mode        = String(row[COL_MODE       - 1] || '').trim().toLowerCase();

    if (account !== MY_ACCOUNT)   continue;
    if (status  !== 'да')         continue;
    if (mode    !== 'полный')     continue;
    if (!emailRaw)                continue;
    if (fuStatus === 'ответил')   continue;
    if (fuStatus === 'завершён')  continue;
    if (fuStatus === 'нет треда') continue;
    if (fuStatus === 'bounce')    continue;
    if (stopOnReply && replyStatus === 'да') continue;
    if (fuNum >= maxFU)           continue;
    if (!fuNextRaw)               continue;

    const fuNextDate = parseLogDate(fuNextRaw);
    if (!fuNextDate || isNaN(fuNextDate)) continue;

    const fuDateOnly = Date.UTC(fuNextDate.getUTCFullYear(), fuNextDate.getUTCMonth(), fuNextDate.getUTCDate());
    if (nowDateOnly !== fuDateOnly) continue;
    if (nowUTC < fuNextDate)        continue;

    candidates.push({ rowIndex: i + 1, rowArrayIndex: i, fuNextDate });
  }

  if (candidates.length === 0) {
    Logger.log(`ℹ️ sendFollowUps (${MY_ACCOUNT}): нет кандидатов`);
    return;
  }

  Logger.log(`📋 sendFollowUps (${MY_ACCOUNT}): кандидатов ${candidates.length}`);

  let sentCount = 0;

  for (let b = 0; b < candidates.length; b++) {
    const { rowIndex, rowArrayIndex, fuNextDate } = candidates[b];
    const row = dataRows[rowArrayIndex];

    const emailRaw  = row[COL_EMAIL    - 1];
    const name      = row[COL_NAME     - 1];
    const dateSent  = row[COL_DATE     - 1];
    const styleRaw  = String(row[COL_TEMPLATE - 1] || '').trim();
    const fuNum     = parseInt(row[COL_FU_NUM  - 1], 10) || 0;
    const nextFuNum = fuNum + 1;
    const emails    = parseEmails(String(emailRaw));

    let cachedThread   = null;
    let cachedOrigSubj = null;
    if (canSearch()) {
      const found    = findSentThreadWithSubject(emails, dateSent);
      cachedThread   = found.thread;
      cachedOrigSubj = found.subject;
    }

    if (!cachedThread) {
      Logger.log(`⛔ ФА №${nextFuNum} отменён — тред не найден (${name})`);
      dataSheet.getRange(rowIndex, COL_FU_STATUS).setValue('нет треда');
      dataSheet.getRange(rowIndex, COL_FU_NEXT).setValue('');
      writeLog(logSheet, nowUTC, 'ФОЛОУ-АП', String(emailRaw), name,
               '—', '—', '—', '⛔ Тред не найден', `ФА №${nextFuNum} — рассылка остановлена`);
      continue;
    }

    const tpl = pickFollowUpTemplate(fuTemplates, nextFuNum, styleRaw, logUsedIds, sessionUsedIds);
    if (!tpl) {
      Logger.log(`⚠️ Нет шаблона ФА №${nextFuNum}, стиль "${styleRaw}" — строка ${rowIndex}`);
      continue;
    }

    const fuBody = fillVars(tpl.body, name, sender, cachedOrigSubj || '').replace(/\n/g, '<br>');

    let wasSent = false;
    try {
      sendReplyInThread(cachedThread, emails[0], cachedOrigSubj || '', fuBody);
      Logger.log(`🔁 ФА №${nextFuNum} (reply-in-thread) → ${name} | ${emails[0]} | ${tpl.id}`);
      writeLog(logSheet, nowUTC, 'ФОЛОУ-АП', String(emailRaw), name,
               tpl.id, tpl.style, '[reply in thread]', '✅ Отправлено', `ФА №${nextFuNum}`);
      if (!sessionUsedIds.includes(tpl.id)) sessionUsedIds.push(tpl.id);
      wasSent = true;
    } catch (e) {
      Logger.log(`❌ ФА ошибка: ${e.message}`);
      writeLog(logSheet, nowUTC, 'ФОЛОУ-АП', String(emailRaw), name,
               tpl.id, tpl.style, '', `❌ ${e.message}`, `ФА №${nextFuNum}`);
    }

    if (wasSent) {
      sentCount++;
      dataSheet.getRange(rowIndex, COL_FU_NUM).setValue(nextFuNum);

      const isDone = nextFuNum >= maxFU;
      if (isDone) {
        dataSheet.getRange(rowIndex, COL_FU_STATUS).setValue('завершён');
        dataSheet.getRange(rowIndex, COL_FU_NEXT).setValue('');
        dataRows[rowArrayIndex][COL_FU_NEXT - 1] = '';
      } else {
        const baseSentUTC  = nowUTC;
        const nextInterval = getFuInterval(cfg, nextFuNum + 1);
        const nextFuDate   = pickSmartFuTime(baseSentUTC, nextInterval, dataRows, logRows, sessionSlots);

        dataSheet.getRange(rowIndex, COL_FU_STATUS).setValue('ожидание');
        dataSheet.getRange(rowIndex, COL_FU_NEXT).setValue(formatUkrTime(nextFuDate));
        dataRows[rowArrayIndex][COL_FU_NEXT - 1] = formatUkrTime(nextFuDate);
      }
    }
  }

  Logger.log(`✅ sendFollowUps: отправлено ${sentCount}`);
}

// ============================================================
//  ФУНКЦИЯ 4: ПРОВЕРКА BOUNCE
// ============================================================
function checkBounces() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cfg = loadSettings(ss);

  if (!cfgBool(cfg, 'BOUNCE_CHECK_ENABLED', true)) {
    Logger.log('⏸ Bounce-проверка отключена');
    return;
  }

  const dataSheet = ss.getSheetByName(SHEET_DATA);
  const logSheet  = ss.getSheetByName(SHEET_LOG);
  const nowUTC    = new Date();

  if (!canSearch()) return;

  const bounceQuery =
    'from:(mailer-daemon OR "Mail Delivery Subsystem" OR postmaster@) ' +
    'subject:(Delivery OR Undeliverable OR "returned to sender" OR "failed to deliver" OR "address not found") ' +
    'newer_than:14d in:inbox';

  const threads = gmailSearch(bounceQuery, 0, 30);

  if (!threads || threads.length === 0) {
    Logger.log('✅ Bounce-проверка: уведомлений нет');
    return;
  }

  const bounced = new Set();
  const patterns = [
    /Final-Recipient:\s*rfc822;\s*<?([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})>?/gi,
    /Original-Recipient:\s*rfc822;\s*<?([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})>?/gi,
    /The email address\s+([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})\s+(does not|doesn't|wasn't)/gi,
    /550[^\n]*?([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})/gi,
  ];

  for (const thread of threads) {
    for (const msg of thread.getMessages()) {
      const body = msg.getPlainBody();
      for (const re of patterns) {
        re.lastIndex = 0;
        let m;
        while ((m = re.exec(body)) !== null) {
          const candidate = m[1].toLowerCase().trim();
          if (isValidEmail(candidate) && !candidate.includes('mailer-daemon')) {
            bounced.add(candidate);
          }
        }
      }
    }
  }

  if (bounced.size === 0) {
    Logger.log(`ℹ️ Bounce: ${threads.length} писем, адреса не распознаны — проверь вручную`);
    return;
  }

  Logger.log(`⚠️ Bounce: адреса — ${[...bounced].join(', ')}`);

  const rows   = dataSheet.getDataRange().getValues();
  let   marked = 0;
  for (let i = 1; i < rows.length; i++) {
    const emailRaw = String(rows[i][COL_EMAIL - 1] || '').toLowerCase();
    const name     = rows[i][COL_NAME - 1];
    for (const addr of bounced) {
      if (emailRaw.includes(addr)) {
        dataSheet.getRange(i + 1, COL_FU_STATUS).setValue('bounce');
        dataSheet.getRange(i + 1, COL_FU_NEXT).setValue('');
        writeLog(logSheet, nowUTC, 'BOUNCE', emailRaw, name,
                 '—', '—', '—', '⚠️ Bounce detected', addr);
        marked++;
        break;
      }
    }
  }
  Logger.log(`⚠️ Bounce: помечено строк — ${marked}`);
}

// ============================================================
//  ТРИГГЕРЫ И СТАТУС
// ============================================================
function updateTriggerStatus(ss) {
  const ssObj    = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
  const isActive = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'sendScheduledEmails');
  setStatus(isActive ? 'active' : 'inactive', ssObj);
}

function setStatus(type, ss) {
  const ssObj = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ssObj.getSheetByName(SHEET_STATUS);
  if (!sheet) return;
  const cell = sheet.getRange(CELL_STATUS);
  if (type === 'active') {
    cell.setValue('✅ Триггер активен');
    cell.setBackground('#E8F5E9').setFontColor('#1B5E20');
  } else {
    cell.setValue('❌ Триггер не активен');
    cell.setBackground('#FFEBEE').setFontColor('#B71C1C');
  }
  cell.setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
}

function createTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendScheduledEmails').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('checkReplies').timeBased().everyHours(3).create();
  ScriptApp.newTrigger('sendFollowUps').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('checkBounces').timeBased().everyHours(6).create();
  ScriptApp.newTrigger('updateTriggerStatus').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('cleanLog').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
  setStatus('active');
  Logger.log('✅ Триггеры созданы');
}

function deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  setStatus('inactive');
  Logger.log('Все триггеры удалены');
}

function cleanLog() {
  const ss            = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cfg           = loadSettings(ss);
  const retentionDays = cfgInt(cfg, 'LOG_RETENTION_DAYS', 30);
  const logSheet      = ss.getSheetByName(SHEET_LOG);
  if (!logSheet) return;
  const data   = logSheet.getDataRange().getValues();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - retentionDays);
  let deleted = 0;
  for (let i = data.length - 1; i >= LOG_START_ROW - 1; i--) {
    const rowDate = parseLogDate(data[i][LCOL_DATE - 1]);
    if (rowDate && !isNaN(rowDate) && rowDate < cutoff) {
      logSheet.deleteRow(i + 1);
      deleted++;
    }
  }
  Logger.log(`🧹 cleanLog: удалено ${deleted} строк (старше ${retentionDays} дней)`);
}

function resetStatus() {
  const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dataSheet = ss.getSheetByName(SHEET_DATA);
  const rows      = dataSheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][COL_STATUS - 1] === 'да') {
      dataSheet.getRange(i + 1, COL_STATUS).setValue('');
      dataSheet.getRange(i + 1, COL_TIME_UKR).setValue('');
      dataSheet.getRange(i + 1, COL_REPLY).setValue('');
      dataSheet.getRange(i + 1, COL_REPLY_DATE).setValue('');
      dataSheet.getRange(i + 1, COL_FU_NUM).setValue('');
      dataSheet.getRange(i + 1, COL_FU_STATUS).setValue('');
      dataSheet.getRange(i + 1, COL_FU_NEXT).setValue('');
      count++;
    }
  }
  Logger.log(`Сброшено строк: ${count}`);
}
