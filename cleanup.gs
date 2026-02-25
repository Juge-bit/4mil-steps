function cleanupDailyStepsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(DASHBOARD_SHEET_NAME || 'daily_steps');
  if (!sh) throw new Error(`Sheet "${DASHBOARD_SHEET_NAME || 'daily_steps'}" not found`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) {
    console.log('Nothing to clean (no data rows).');
    return;
  }

  const values = sh.getRange(1, 1, lastRow, Math.max(4, lastCol)).getValues();
  const rows = values.slice(1);

  const bestByDate = resolveBestRowsByDate_(rows);

  const out = Array.from(bestByDate.values())
    .sort((a, b) => a.date.localeCompare(b.date))
    .map((r) => [r.date, Math.round(r.steps), r.sourceFile || '', r.updatedAtDate || r.updatedAtRaw || '']);

  sh.getRange(1, 1, sh.getMaxRows(), 4).clearContent();
  sh.getRange(1, 1, 1, 4).setValues([['date', 'steps', 'source_file', 'updated_at']]);
  if (out.length) sh.getRange(2, 1, out.length, 4).setValues(out);

  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange('B:B').setNumberFormat('0');
  sh.getRange('D:D').setNumberFormat('yyyy-mm-dd hh:mm:ss');

  console.log(`Cleanup complete. Kept ${out.length} unique dates.`);
}
