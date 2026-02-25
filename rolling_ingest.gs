const ROLLING_RAW_SHEET_NAME = 'rolling_steps_raw';

/**
 * Ingest rolling 30-day CSV to a separate raw sheet.
 * This must NOT overwrite daily_steps directly.
 */
function ingestRolling30DaysSteps() {
  const folder = DriveApp.getFolderById(INGEST_FOLDER_ID);

  const file = findLatestRollingStepsFile_(folder);
  if (!file) {
    console.log('No rolling 30 days steps file found in folder.');
    return;
  }

  console.log(
    `Using rolling file: ${file.getName()} | created=${file.getDateCreated().toISOString()} | updated=${file.getLastUpdated().toISOString()}`
  );

  const rows = parseDailyStepsFromRollingCsv_(file);
  console.log(`Parsed ${rows.length} daily rows from rolling file.`);

  for (const r of rows) {
    upsertRollingRaw_(r.date, r.steps, file.getName(), file.getLastUpdated());
  }

  console.log('Rolling ingest done (raw sheet only).');
}

function findLatestRollingStepsFile_(folder) {
  const it = folder.getFiles();
  const re = /^Steps \d{4}\.\d{2}\.\d{2}-\d{4}\.\d{2}\.\d{2} Health Connect\.csv$/;

  let best = null;
  let bestUpdated = null;

  while (it.hasNext()) {
    const f = it.next();
    const name = f.getName();
    if (!re.test(name)) continue;

    const updated = f.getLastUpdated();
    if (!best || updated > bestUpdated) {
      best = f;
      bestUpdated = updated;
    }
  }

  return best;
}

function parseDailyStepsFromRollingCsv_(file) {
  const text = file.getBlob().getDataAsString('UTF-8');
  const rows = Utilities.parseCsv(text);
  if (!rows || rows.length < 2) return [];

  const header = rows[0].map(c => String(c || '').trim().toLowerCase());
  const dateIdx = header.indexOf('date');
  const stepsIdx = header.indexOf('steps');

  if (dateIdx < 0 || stepsIdx < 0) {
    console.log(`Unexpected rolling header: ${JSON.stringify(header)}`);
    return [];
  }

  const sumByDate = new Map();

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length <= Math.max(dateIdx, stepsIdx)) continue;

    const dayKey = normalizeRollingDateTime_(String(r[dateIdx] || '').trim());
    if (!dayKey) continue;

    const steps = rollingNormalizeSteps_(r[stepsIdx]);
    if (steps <= 0) continue;

    sumByDate.set(dayKey, (sumByDate.get(dayKey) || 0) + steps);
  }

  return Array.from(sumByDate.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([date, steps]) => ({ date, steps: Math.round(steps) }));
}

function normalizeRollingDateTime_(s) {
  const m = String(s || '').trim().match(/^(\d{4})\.(\d{2})\.(\d{2})/);
  if (!m) return '';
  return `${m[1]}-${m[2]}-${m[3]}`;
}

function rollingNormalizeSteps_(v) {
  const s = String(v ?? '').trim().replace(/\s/g, '').replace(',', '.');
  const n = Number(s);
  return Number.isFinite(n) ? Math.round(n) : 0;
}

function upsertRollingRaw_(date, steps, sourceFile, updatedAt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(ROLLING_RAW_SHEET_NAME) || ss.insertSheet(ROLLING_RAW_SHEET_NAME);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 4).setValues([['date', 'steps', 'source_file', 'updated_at']]);
  }

  const lastRow = sh.getLastRow();
  const existing = {};

  if (lastRow >= 2) {
    const values = sh.getRange(2, 1, lastRow - 1, 3).getValues();
    values.forEach((r, i) => {
      const d = normalizeDateKey_(r[0]);
      if (!d) return;
      existing[d] = i + 2;
    });
  }

  const rowIndex = existing[date];
  const stamp = updatedAt instanceof Date ? updatedAt : new Date();

  if (rowIndex) {
    sh.getRange(rowIndex, 2, 1, 3).setValues([[Math.round(steps), sourceFile, stamp]]);
  } else {
    sh.appendRow([date, Math.round(steps), sourceFile, stamp]);
  }
}
