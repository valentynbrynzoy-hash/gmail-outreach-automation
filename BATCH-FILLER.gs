// ============================================================
//  BATCH FILLER v4
//  - Лимит 10 писем на аккаунт в день (жёстко)
//  - Каждый аккаунт стартует параллельно с 09:00
//  - Диалог: только дата + скип выходных
// ============================================================

// ── Настройки ─────────────────────────────────────────────────
const BF_INTERVAL_MIN  = 5;        // минут между письмами одного аккаунта
const BF_START_TIME    = '09:00';  // старт (украинское время)
const BF_MAX_PER_ACC   = 10;       // жёсткий лимит писем на аккаунт в день

// ── Листы ─────────────────────────────────────────────────────
const BF_SHEET_DATA      = '📮 Почты';
const BF_SHEET_TEMPLATES = '📧 Шаблоны';

// ── Колонки листа 📮 Почты ────────────────────────────────────
const BF_COL_NAME     = 1;
const BF_COL_GEO      = 3;
const BF_COL_STATUS   = 4;
const BF_COL_DATE     = 5;
const BF_COL_ACCOUNT  = 6;
const BF_COL_TIME_LOC = 7;

// ── Строки аккаунтов в листе 📧 Шаблоны ──────────────────────
const BF_SENDERS_START = 3;
const BF_SENDERS_END   = 12;
const BF_SCOL_EMAIL    = 1;
const BF_SCOL_NAME     = 2;
const BF_SCOL_ACTIVE   = 7;



// ── Меню ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📬 Рассылка')
    .addItem('🚀 Заполнить новые каналы', 'bfShowDialog')
    .addToUi();
}

// ── Серверные функции ─────────────────────────────────────────

function bfGetInitData() {
  const accounts  = bfGetActiveAccounts();
  const emptyRows = bfFindEmptyRows();
  const accCount  = accounts.length;
  const dailyCap  = accCount * BF_MAX_PER_ACC;
  const daysNeeded = dailyCap > 0 ? Math.ceil(emptyRows.length / dailyCap) : 0;

  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const defaultDate = Utilities.formatDate(tomorrow, 'Europe/Kiev', 'yyyy-MM-dd');

  return {
    emptyCount:  emptyRows.length,
    accCount,
    dailyCap,
    daysNeeded,
    defaultDate,
  };
}

// Вычисляем сколько рабочих дней нужно начиная с даты (с учётом скипа выходных)
function bfCalcDaysNeeded(startDateStr, daysNeeded, skipWeekends) {
  if (!skipWeekends) return daysNeeded;
  // Считаем сколько календарных дней займут daysNeeded рабочих дней
  const [y, m, d] = startDateStr.split('-').map(Number);
  const cur = new Date(y, m - 1, d);
  let workDays = 0;
  let calDays  = 0;
  const limit  = 365;
  while (workDays < daysNeeded && calDays < limit) {
    const dow = cur.getDay();
    if (dow !== 0 && dow !== 6) workDays++;
    cur.setDate(cur.getDate() + 1);
    calDays++;
  }
  // Возвращаем дату последнего дня
  const end = new Date(y, m - 1, d);
  end.setDate(end.getDate() + calDays - 1);
  return Utilities.formatDate(end, 'Europe/Kiev', 'yyyy-MM-dd');
}

function bfGetActiveAccounts() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BF_SHEET_TEMPLATES);
  if (!sheet) return [];
  const data   = sheet.getDataRange().getValues();
  const result = [];
  for (let i = BF_SENDERS_START - 1; i <= BF_SENDERS_END - 1 && i < data.length; i++) {
    const email  = String(data[i][BF_SCOL_EMAIL  - 1] || '').trim();
    const active = String(data[i][BF_SCOL_ACTIVE - 1] || '').trim();
    if (email && email.includes('@') && (active === '✅' || active.toLowerCase() === 'true')) {
      result.push({ email });
    }
  }
  return result;
}

function bfFindEmptyRows() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BF_SHEET_DATA);
  if (!sheet) return [];
  const data  = sheet.getDataRange().getValues();
  const empty = [];
  for (let i = 1; i < data.length; i++) {
    const name    = String(data[i][BF_COL_NAME    - 1] || '').trim();
    const geo     = String(data[i][BF_COL_GEO     - 1] || '').trim();
    const account = String(data[i][BF_COL_ACCOUNT - 1] || '').trim();
    const status  = String(data[i][BF_COL_STATUS  - 1] || '').trim();
    if (name && geo && !account && status !== 'да') {
      empty.push({ rowIdx: i + 1, geo: geo.toUpperCase() });
    }
  }
  return empty;
}

/**
 * config = {
 *   startDate:    'yyyy-MM-dd',
 *   skipWeekends: bool
 * }
 */
function bfFillRows(config) {
  const accounts  = bfGetActiveAccounts();
  const emptyRows = bfFindEmptyRows();

  if (emptyRows.length === 0) return { ok: true, filled: 0 };
  if (accounts.length === 0)  return { ok: false, error: 'Нет активных аккаунтов в листе Шаблоны' };

  const dailyCap = accounts.length * BF_MAX_PER_ACC;
  const daysNeeded = Math.ceil(emptyRows.length / dailyCap);
  const dates = bfBuildDates(config.startDate, daysNeeded, config.skipWeekends);

  if (dates.length === 0) return { ok: false, error: 'Нет рабочих дней в выбранном диапазоне' };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(BF_SHEET_DATA);
  const [startH, startM] = BF_START_TIME.split(':').map(Number);

  let cursor      = 0;
  let totalFilled = 0;

  for (const date of dates) {
    if (cursor >= emptyRows.length) break;

    // Берём chunk на этот день (не больше dailyCap)
    const chunk = emptyRows.slice(cursor, cursor + dailyCap);

    // Блоки: сначала все строки акк0, потом акк1, и т.д.
    accounts.forEach((acc, accIdx) => {
      const block = chunk.slice(accIdx * BF_MAX_PER_ACC, (accIdx + 1) * BF_MAX_PER_ACC);
      block.forEach((row, posInAcc) => {
        const totalMin = startH * 60 + startM + posInAcc * BF_INTERVAL_MIN;
        const h = Math.floor(totalMin / 60) % 24;
        const m = totalMin % 60;
        sheet.getRange(row.rowIdx, BF_COL_DATE).setValue(date);
        sheet.getRange(row.rowIdx, BF_COL_ACCOUNT).setValue(acc.email);
        sheet.getRange(row.rowIdx, BF_COL_TIME_LOC).setValue(bfPad(h) + ':' + bfPad(m) + ':00');
      });
    });

    cursor      += chunk.length;
    totalFilled += chunk.length;
  }

  return { ok: true, filled: totalFilled, days: dates.length };
}

// ── Утилиты ───────────────────────────────────────────────────

function bfBuildDates(startDateStr, numDays, skipWeekends) {
  const [y, m, d] = startDateStr.split('-').map(Number);
  const dates = [];
  const cur   = new Date(y, m - 1, d);
  const limit = new Date(y, m - 1, d);
  limit.setDate(limit.getDate() + 120);
  while (dates.length < numDays && cur <= limit) {
    const dow = cur.getDay();
    if (!skipWeekends || (dow !== 0 && dow !== 6)) {
      dates.push(new Date(cur));
    }
    cur.setDate(cur.getDate() + 1);
  }
  return dates;
}

function bfPad(n) { return String(n).padStart(2, '0'); }

// ── Диалог ────────────────────────────────────────────────────

function bfShowDialog() {
  const html = HtmlService
    .createHtmlOutput(BF_HTML)
    .setWidth(400)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, '🚀 Заполнить новые каналы');
}

const BF_HTML = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * { box-sizing: border-box; font-family: 'Google Sans', Arial, sans-serif; }
  body { margin: 0; padding: 16px; background: #f8f9fa; font-size: 13px; color: #202124; }
  h3   { margin: 0 0 12px; font-size: 15px; }
  .card { background: #fff; border-radius: 8px; padding: 14px 16px; margin-bottom: 10px;
          box-shadow: 0 1px 3px rgba(0,0,0,.1); }
  label { display: block; font-weight: 600; margin-bottom: 5px; color: #444; font-size: 12px; }
  input[type=date] { width: 100%; padding: 8px 10px; border: 1px solid #ddd;
                     border-radius: 6px; font-size: 13px; outline: none; }
  input[type=date]:focus { border-color: #1a73e8; }
  .weekend-row { display: flex; align-items: center; gap: 8px; margin-top: 10px; font-size: 12px; }
  .weekend-row input { width: auto; }
  .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; }
  .info-item { background: #f1f3f4; border-radius: 6px; padding: 8px 10px; }
  .info-item .val { font-size: 17px; font-weight: 700; color: #1a73e8; line-height: 1.2; }
  .info-item .lbl { font-size: 11px; color: #5f6368; margin-top: 2px; }
  .info-item.warn .val { color: #e37400; }
  .btn { width: 100%; padding: 10px; background: #1a73e8; color: #fff; font-size: 14px;
         font-weight: 600; border: none; border-radius: 8px; cursor: pointer; margin-top: 4px; }
  .btn:hover:not(:disabled) { background: #1558b0; }
  .btn:disabled { background: #9aa0a6; cursor: not-allowed; }
  #status { text-align: center; padding: 6px 0 2px; font-size: 12px; min-height: 22px; }
  .ok  { color: #188038; font-weight: 600; }
  .err { color: #c5221f; font-weight: 600; }
</style>
</head>
<body>
<h3>🚀 Заполнить новые каналы</h3>

<div class="card">
  <label>📅 Дата старта</label>
  <input type="date" id="startDate" onchange="updateInfo()">
  <div class="weekend-row">
    <input type="checkbox" id="skipWeekends" checked onchange="updateInfo()">
    <span>Пропускать выходные (сб/вс)</span>
  </div>
</div>

<div class="card">
  <div class="info-grid" id="info-grid">
    <div class="info-item"><div class="val" id="v-rows">—</div><div class="lbl">Строк найдено</div></div>
    <div class="info-item"><div class="val" id="v-accs">—</div><div class="lbl">Активных аккаунтов</div></div>
    <div class="info-item"><div class="val" id="v-cap">—</div><div class="lbl">Ёмкость в день</div></div>
    <div class="info-item"><div class="val" id="v-days">—</div><div class="lbl">Рабочих дней</div></div>
  </div>
</div>

<button class="btn" id="btn" onclick="doFill()" disabled>⚡ Заполнить</button>
<div id="status"></div>

<script>
  var D = null;

  google.script.run
    .withSuccessHandler(function(d) {
      D = d;
      document.getElementById('startDate').value = d.defaultDate;
      document.getElementById('v-rows').textContent = d.emptyCount;
      document.getElementById('v-accs').textContent = d.accCount;
      document.getElementById('v-cap').textContent  = d.dailyCap + '/день';
      document.getElementById('v-days').textContent = d.daysNeeded;
      if (d.emptyCount > 0 && d.accCount > 0) {
        document.getElementById('btn').disabled = false;
      } else if (d.emptyCount === 0) {
        document.getElementById('status').innerHTML = '<span class="ok">✅ Всё уже заполнено</span>';
      } else {
        document.getElementById('status').innerHTML = '<span class="err">❌ Нет активных аккаунтов</span>';
      }
    })
    .withFailureHandler(function(e) {
      document.getElementById('status').innerHTML = '<span class="err">Ошибка: ' + e.message + '</span>';
    })
    .bfGetInitData();

  function updateInfo() {
    if (!D) return;
    var skip = document.getElementById('skipWeekends').checked;
    // Пересчитываем дни с учётом выходных (приближённо на клиенте)
    var days = D.daysNeeded;
    if (skip && days > 0) {
      // Добавляем ~40% за выходные (2/7 дней выходные)
      // Точный расчёт делает сервер при реальном заполнении
      days = Math.ceil(days * 7 / 5);
    }
    document.getElementById('v-days').textContent = days + (skip ? ' раб.' : '');
  }

  function doFill() {
    var btn = document.getElementById('btn');
    var st  = document.getElementById('status');
    btn.disabled = true;
    st.innerHTML = '⏳ Заполняю...';

    google.script.run
      .withSuccessHandler(function(r) {
        if (r.ok) {
          google.script.host.close();
        } else {
          st.innerHTML = '<span class="err">❌ ' + (r.error || 'Ошибка') + '</span>';
          btn.disabled = false;
        }
      })
      .withFailureHandler(function(e) {
        st.innerHTML = '<span class="err">❌ ' + e.message + '</span>';
        btn.disabled = false;
      })
      .bfFillRows({
        startDate:    document.getElementById('startDate').value,
        skipWeekends: document.getElementById('skipWeekends').checked,
      });
  }
</script>
</body>
</html>`;
