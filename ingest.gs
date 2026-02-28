const INGEST_FOLDER_ID = '1PowR9TrxAGvYnx2omjFObTH98AtREyw9';
const INGEST_SHEET_NAME = 'daily_steps';

function ingestHealthSyncSteps() {
  const folder = DriveApp.getFolderById(INGEST_FOLDER_ID);
  const today = new Date();

  // Ingest last 3 completed days (self-healing for late sync)
  for (let back = 3; back >= 1; back--) {
    const d = new Date(today);
    d.setDate(today.getDate() - back);
    ingestOneDay_(folder, d);
  }
}

function ingestOneDay_(folder, dateObj) {
  const iso = ingestToISODate_(dateObj);
  const ymdDots = iso.replace(/-/g, '.');

  const targetName = `Steps ${ymdDots} Health Connect.csv`;
  const file = findFileByExactName_(folder, targetName);

  if (!file) {
    console.log(`No file found for ${iso}: ${targetName}`);
    return;
  }

  const steps = parseDailyStepsFromCsv_(file);
  upsertDailySteps_(iso, steps, file.getName(), file.getLastUpdated());
  console.log(`Imported ${steps} steps for ${iso} from ${file.getName()}`);
}

function findFileByExactName_(folder, exactName) {
  const it = folder.getFilesByName(exactName);
  let latest = null;

  while (it.hasNext()) {
    const candidate = it.next();
    if (!latest || candidate.getLastUpdated().getTime() > latest.getLastUpdated().getTime()) {
      latest = candidate;
    }
  }

  return latest;
}

function parseDailyStepsFromCsv_(file) {
  const text = file.getBlob().getDataAsString('UTF-8');
  const rows = Utilities.parseCsv(text);
  if (!rows || rows.length < 2) return 0;

  const header = rows[0].map(c => String(c || '').trim().toLowerCase());

  let stepsCol = -1;
  for (let i = 0; i < header.length; i++) {
    if (header[i].includes('step')) {
      stepsCol = i;
      break;
    }
  }

  let sumRaw = 0;
  const maxBy10MinBucket = new Map();

  const dateIdx = header.indexOf('date');
  const timeIdx = header.indexOf('time');

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    let n = 0;

    if (stepsCol >= 0 && stepsCol < row.length) {
      n = ingestNormalizeSteps_(row[stepsCol]);
    } else {
      for (const cell of row) {
        n = ingestNormalizeSteps_(cell);
        if (n > 0) break;
      }
    }

    sumRaw += n;

    const bucket = parseTenMinuteBucket_(row, dateIdx, timeIdx);
    if (bucket) {
      const prev = maxBy10MinBucket.get(bucket) || 0;
      if (n > prev) maxBy10MinBucket.set(bucket, n);
    }
  }

  const sumBucketMax = Array.from(maxBy10MinBucket.values()).reduce((a, b) => a + b, 0);

  // Health Sync can contain both sub-events + 10-minute summary rows in the same daily file.
  // When raw sum is clearly inflated versus bucket-max sum, trust bucket-max.
  if (sumBucketMax > 0 && sumRaw > sumBucketMax * 1.6) {
    console.log(`Using bucket-max parse for ${file.getName()} (raw=${Math.round(sumRaw)}, bucketMax=${Math.round(sumBucketMax)})`);
    return Math.round(sumBucketMax);
  }

  return Math.round(sumRaw);
}

function parseTenMinuteBucket_(row, dateIdx, timeIdx) {
  const dateCell = dateIdx >= 0 && dateIdx < row.length ? String(row[dateIdx] || '').trim() : '';
  const timeCell = timeIdx >= 0 && timeIdx < row.length ? String(row[timeIdx] || '').trim() : '';

  // Prefer Date column like "2026.02.24 19:40:52"
  let m = dateCell.match(/^(\d{4})\.(\d{2})\.(\d{2})\s+(\d{2}):(\d{2})(?::\d{2})?$/);
  if (!m && timeCell) {
    // Fallback if Date column not parseable
    const dm = dateCell.match(/^(\d{4})\.(\d{2})\.(\d{2})/);
    const tm = timeCell.match(/^(\d{2}):(\d{2})(?::\d{2})?$/);
    if (dm && tm) {
      m = [dateCell, dm[1], dm[2], dm[3], tm[1], tm[2]];
    }
  }

  if (!m) return '';

  const yyyy = m[1];
  const mm = m[2];
  const dd = m[3];
  const hh = m[4];
  const min = Number(m[5]);
  if (!Number.isFinite(min)) return '';

  const bucketMin = String(Math.floor(min / 10) * 10).padStart(2, '0');
  return `${yyyy}-${mm}-${dd} ${hh}:${bucketMin}`;
}

function ingestToISODate_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function ingestNormalizeSteps_(v) {
  const s = String(v ?? '').trim().replace(/\s/g, '').replace(',', '.');
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function upsertDailySteps_(date, steps, sourceFile, updatedAtOverride) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INGEST_SHEET_NAME) || ss.insertSheet(INGEST_SHEET_NAME);

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
      const src = String(r[2] || '').trim().toLowerCase();
      existing[d] = { row: i + 2, source: src };
    });
  }

  const now = updatedAtOverride instanceof Date ? updatedAtOverride : new Date();
  const info = existing[date];

  if (info && info.source === 'manual') {
    console.log(`Skip overwrite for ${date} (manual). Incoming: ${steps} from ${sourceFile}`);
    return;
  }

  if (info && info.row) {
    sh.getRange(info.row, 2, 1, 3).setValues([[Math.round(steps), sourceFile, now]]);
  } else {
    sh.appendRow([date, Math.round(steps), sourceFile, now]);
  }
}
