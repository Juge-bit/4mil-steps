/**
 * Canonical dashboard endpoint for this project.
 * Keep only ONE active dashboard doGet in production deployment.
 */
const DASHBOARD_SHEET_NAME = 'daily_steps';
const GOAL_STEPS = 4000000;

function doGet(e) {
  try {
    const mode = (e && e.parameter && e.parameter.mode) ? String(e.parameter.mode) : 'dashboard';
    if (mode !== 'dashboard') {
      return jsonOut_(200, { ok: true, message: 'Use ?mode=dashboard' });
    }

    const year = (e && e.parameter && e.parameter.year) ? Number(e.parameter.year) : 2026;
    if (!Number.isFinite(year) || year < 2000 || year > 2100) {
      return jsonOut_(400, { ok: false, error: 'Invalid year' });
    }

    refreshRecentDailyIngestForDashboard_(3);
    const payload = buildDashboardPayload_(year);
    return jsonOut_(200, payload);
  } catch (err) {
    return jsonOut_(500, { ok: false, error: String(err) });
  }
}


function refreshRecentDailyIngestForDashboard_(daysBack) {
  try {
    if (typeof ingestOneDay_ !== 'function' || typeof INGEST_FOLDER_ID === 'undefined') {
      return;
    }

    const folder = DriveApp.getFolderById(INGEST_FOLDER_ID);
    const today = new Date();
    const safeDaysBack = Math.max(1, Number(daysBack) || 1);

    for (let back = safeDaysBack; back >= 1; back--) {
      const d = new Date(today);
      d.setDate(today.getDate() - back);
      ingestOneDay_(folder, d);
    }
  } catch (err) {
    console.log(`Dashboard ingest refresh skipped: ${err}`);
  }
}

function buildDashboardPayload_(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sh) return { ok: false, error: `Sheet "${DASHBOARD_SHEET_NAME}" not found` };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return {
      ok: true,
      year,
      goal: GOAL_STEPS,
      totalSteps: 0,
      percent: 0,
      lastUpdatedAt: null,
      today: isoToday_(),
      todaySteps: 0,
      yesterday: isoYesterday_(),
      yesterdaySteps: 0,
      last14Days: [],
      topDays: [],
      missingDates: []
    };
  }

  const rows = sh.getRange(2, 1, lastRow - 1, 4).getValues();
  const resolvedByDate = resolveBestRowsByDate_(rows);

  let lastUpdatedAt = null;
  for (const obj of resolvedByDate.values()) {
    if (obj.updatedAtDate && (!lastUpdatedAt || obj.updatedAtDate > lastUpdatedAt)) lastUpdatedAt = obj.updatedAtDate;
  }

  const start = `${year}-01-01`;
  const end = `${year}-12-31`;

  const stepsByDay = [];
  let total = 0;
  for (const [date, obj] of resolvedByDate.entries()) {
    if (date >= start && date <= end) {
      total += obj.steps;
      stepsByDay.push({ date, steps: obj.steps });
    }
  }
  stepsByDay.sort((a, b) => a.date.localeCompare(b.date));

  const topDays = stepsByDay
    .slice()
    .sort((a, b) => (b.steps - a.steps) || a.date.localeCompare(b.date))
    .slice(0, 5);

  const percent = Math.max(0, Math.min(100, (total / GOAL_STEPS) * 100));
  const today = isoToday_();
  const yesterday = isoYesterday_();

  const todaySteps = resolvedByDate.has(today) ? resolvedByDate.get(today).steps : 0;
  const yesterdaySteps = resolvedByDate.has(yesterday) ? resolvedByDate.get(yesterday).steps : 0;

  const last14Days = [];
  for (let i = 13; i >= 0; i--) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    const iso = toISODate_(d);
    const s = resolvedByDate.has(iso) ? resolvedByDate.get(iso).steps : 0;
    last14Days.push({ date: iso, steps: s });
  }

  const missingDates = [];
  const now = new Date();
  const yearStart = new Date(`${year}-01-01T00:00:00`);
  const yearEndCap = new Date(Math.min(now.getTime(), new Date(`${year}-12-31T00:00:00`).getTime()));
  for (let d = new Date(yearStart); d <= yearEndCap; d.setDate(d.getDate() + 1)) {
    const iso = toISODate_(d);
    if (!resolvedByDate.has(iso)) missingDates.push(iso);
  }

  return {
    ok: true,
    year,
    goal: GOAL_STEPS,
    totalSteps: Math.round(total),
    percent: Number(percent.toFixed(4)),
    lastUpdatedAt: lastUpdatedAt ? lastUpdatedAt.toISOString() : null,
    today,
    todaySteps,
    yesterday,
    yesterdaySteps,
    last14Days,
    topDays,
    missingDates
  };
}

function resolveBestRowsByDate_(rows) {
  const byDate = new Map();

  for (const r of rows) {
    const date = normalizeDateKey_(r[0]);
    if (!date) continue;

    const steps = Number(r[1]) || 0;
    const sourceFile = String(r[2] || '');
    const updatedAtRaw = r[3];
    const updatedAtDate = parseUpdatedAt_(updatedAtRaw);

    const candidate = {
      date,
      steps,
      sourceFile,
      sourceType: classifySource_(sourceFile),
      updatedAtRaw,
      updatedAtDate
    };

    const existing = byDate.get(date);
    if (!existing || isBetterCandidate_(candidate, existing)) byDate.set(date, candidate);
  }

  return byDate;
}

function isBetterCandidate_(a, b) {
  const pa = sourcePriority_(a.sourceType);
  const pb = sourcePriority_(b.sourceType);
  if (pa !== pb) return pa > pb;

  const ta = a.updatedAtDate ? a.updatedAtDate.getTime() : -1;
  const tb = b.updatedAtDate ? b.updatedAtDate.getTime() : -1;
  if (ta !== tb) return ta > tb;

  if (a.steps !== b.steps) return a.steps > b.steps;
  return String(a.sourceFile) > String(b.sourceFile);
}

function sourcePriority_(sourceType) {
  if (sourceType === 'manual') return 400;
  if (sourceType === 'daily') return 300;
  if (sourceType === 'rolling') return 200;
  return 100;
}

function classifySource_(sourceFile) {
  const s = String(sourceFile || '').trim();
  if (!s) return 'unknown';
  if (/^manual$/i.test(s)) return 'manual';
  if (/^Steps\s+\d{4}\.\d{2}\.\d{2}-\d{4}\.\d{2}\.\d{2}\s+Health Connect\.csv$/i.test(s)) return 'rolling';
  if (/^Steps\s+\d{4}\.\d{2}\.\d{2}\s+Health Connect\.csv$/i.test(s)) return 'daily';
  return 'unknown';
}

function parseUpdatedAt_(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;

  const s = String(v).trim();
  const d0 = new Date(s);
  if (!isNaN(d0.getTime())) return d0;

  const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+klo)?\s+(\d{1,2})\.(\d{1,2})(?:\.(\d{1,2}))?$/i);
  if (m) {
    return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4]), Number(m[5]), m[6] ? Number(m[6]) : 0);
  }

  return null;
}

function jsonOut_(statusCode, obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function isoToday_() { return toISODate_(new Date()); }
function isoYesterday_() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return toISODate_(d);
}

function toISODate_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function normalizeDateKey_(v) {
  if (!v) return '';
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);

  const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;

  return '';
}

function rebuildFactSheet_() {
  const FACT_SHEET_NAME = 'steps_fact';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!source) throw new Error(`Sheet "${DASHBOARD_SHEET_NAME}" not found`);

  const lastRow = source.getLastRow();
  if (lastRow < 2) throw new Error('Source sheet has no data rows');

  const rows = source.getRange(2, 1, lastRow - 1, 4).getValues();
  const resolved = resolveBestRowsByDate_(rows);

  let fact = ss.getSheetByName(FACT_SHEET_NAME);
  if (!fact) fact = ss.insertSheet(FACT_SHEET_NAME);

  fact.clearContents();
  fact.getRange(1, 1, 1, 5).setValues([['date', 'steps', 'source_type', 'source_file', 'updated_at']]);

  const out = Array.from(resolved.values())
    .sort((a, b) => a.date.localeCompare(b.date))
    .map((r) => [r.date, r.steps, r.sourceType, r.sourceFile, r.updatedAtDate || r.updatedAtRaw || '']);

  if (out.length) fact.getRange(2, 1, out.length, 5).setValues(out);
}
