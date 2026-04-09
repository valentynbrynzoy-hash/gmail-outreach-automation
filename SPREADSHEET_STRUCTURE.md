# Spreadsheet Structure

Reference for recreating the Google Sheet used by this script.

---

## 📋 Меню (Dashboard)

**Row 1–2:** Title / subtitle (decorative)

**Rows 4–12: Account table**

| Col A | Col B | Col C | Col D | Col E | Col F | Col G | Col H |
|-------|-------|-------|-------|-------|-------|-------|-------|
| # | Gmail account | Name | Script link | Trigger status | Description | Batch checkReplies | Batch FU |

- Script reads `ACCOUNT_STATUS_CELLS` → writes trigger status to column E of each account's row
- Columns G and H store batch cursor positions (integers, managed by script)

**Rows 19–32: Settings (key-value)**

| Col A (key) | Col B (value) |
|-------------|---------------|
| FOLLOWUP_ENABLED | TRUE |
| FOLLOWUP_COUNT | 3 |
| FOLLOWUP_INTERVAL_1 | 1 |
| FOLLOWUP_INTERVAL_2 | 5 |
| FOLLOWUP_INTERVAL_3 | 7 |
| CHECK_REPLIES_ENABLED | TRUE |
| CHECK_REPLIES_DAYS | 30 |
| CHECK_BATCH_SIZE | 150 |
| DAYS_UNIQUE | 5 |
| DAYS_UNIQUE_FOLLOWUP | 5 |
| BOUNCE_CHECK_ENABLED | TRUE |
| LOG_RETENTION_DAYS | 30 |
| FOLLOWUP_STOP_ON_REPLY | TRUE |

---

## 📮 Почты (Blogger data)

**Row 1:** Header row

| Col | Name | Notes |
|-----|------|-------|
| A (1) | Имя блогера | Blogger username / name |
| B (2) | Тип платформы | e.g. ИНСТАГРАМ |
| C (3) | Гео | Must match a key in `GEO_OFFSETS` (e.g. УКРАИНА, ПАКИСТАН) |
| D (4) | Статус отправки | Set to `да` by script after send |
| E (5) | Дата отправки | Date to send (DATE type) |
| F (6) | С какого аккаунта | Gmail account — must match `MY_ACCOUNT` |
| G (7) | Время (местное) | Local send time (TIME type) |
| H (8) | Время (укр) | Actual send time in Ukraine TZ (written by script) |
| I (9) | Почты | Email address(es), comma/newline separated |
| J (10) | Шаблон | Template style name (must match TCOL_STYLE / FCOL_STYLE) |
| K (11) | Ответ получен | Set to `да` by script |
| L (12) | Дата ответа | Written by script |
| M (13) | Фолоу-ап № | Current FU count (written by script) |
| N (14) | Статус фолоу-апа | `ожидание` / `ответил` / `завершён` / `нет треда` / `bounce` |
| O (15) | Дата след. фолоу-апа | Next FU datetime (written by script) |
| P (16) | Режим | `полный` / `только проверка` (set automatically if empty) |

---

## 📧 Шаблоны (Senders + Templates)

### Section 1: Sender profiles (rows 3–12)

| Col | Name |
|-----|------|
| A (1) | Email (must match MY_ACCOUNT) |
| B (2) | First name |
| C (3) | Last name |
| D (4) | Title / position |
| E (5) | Signature block (multi-line OK) |
| F (6) | WhatsApp line |
| G (7) | Active (`✅` or `TRUE`) |

### Section 2: Email templates (rows 14+)

| Col | Name |
|-----|------|
| A (1) | Template ID (e.g. `T1`, `T2`) |
| B (2) | Style (e.g. `Инста`, `Ютуб`) — matches column J of 📮 Почты |
| C (3) | Subject (supports `{{блогер}}` etc.) |
| D (4) | Body HTML (supports template variables) |
| E | *(unused)* |
| F (6) | Active (`✅` or `TRUE`) |

---

## 🔁 Фолоу-апы (Follow-up templates)

**Row 1:** Title  
**Row 2:** Description  
**Row 3:** Header  
**Rows 4+:** Templates

| Col | Name |
|-----|------|
| A (1) | ID (e.g. `FU1-1`, `FU2-3`) |
| B (2) | FU number (1, 2, or 3) |
| C (3) | Style (must match TCOL_STYLE) |
| D (4) | Subject (use `{{тема_оригинала}}` for `Re:` threads) |
| E (5) | Body (supports all template variables) |
| F (6) | Active (`✅` or `TRUE`) |

Multiple rows with the same FU number + style = variants; script picks randomly, avoiding recently used IDs.

---

## 📊 Логи (Log)

**Row 1:** Title  
**Row 2:** Header  
**Rows 3+:** Log entries (appended by script)

| Col | Name |
|-----|------|
| A (1) | Date (formatted `dd.mm.yyyy hh:mm`) |
| B (2) | Type (`РАССЫЛКА` / `ФОЛОУ-АП` / `ПРОВЕРКА` / `BOUNCE`) |
| C (3) | Account |
| D (4) | Recipient email |
| E (5) | Blogger name |
| F (6) | Template ID |
| G (7) | Style |
| H (8) | Subject |
| I (9) | Result |
| J (10) | Notes |
