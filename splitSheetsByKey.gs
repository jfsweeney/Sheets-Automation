/**
 * splitSheetsByKey.gs
 *
 * Moves rows from the "Companies" and "Censuses" worksheets into new
 * "Companies Extract" and "Censuses Extract" worksheets, based on a list
 * of Company IDs stored in a "Keys" sheet.
 *
 * Prerequisites
 * -------------
 *   Keys      – Row 1 is a header. Company IDs to extract start in row 2,
 *               column A. Duplicates in this list are handled automatically.
 *   Companies – Row 1 is a header. Column A is Company ID (one row per ID).
 *   Censuses  – Row 1 is a header. Column A is Company ID (may repeat).
 *
 * Output
 * ------
 *   "Companies Extract" – Created (or replaced) with the Companies header
 *                         plus all rows whose Company ID was in Keys.
 *   "Censuses Extract"  – Created (or replaced) with the Censuses header
 *                         plus all rows whose Company ID was in Keys.
 *   The matching rows are deleted from Companies and Censuses.
 *
 * Quality check
 * -------------
 *   An alert reports total data-row counts before and after; they must match.
 *
 * How to run
 * ----------
 *   1. Open the Google Sheet.
 *   2. Extensions > Apps Script, paste this file, save.
 *   3. Select the splitSheetsByKey function and click Run.
 *      (You will be asked to grant permissions the first time.)
 */

// ─────────────────────────────────────────────────────────────────────────────
// Entry point
// ─────────────────────────────────────────────────────────────────────────────

function splitSheetsByKey() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ── Confirmation ────────────────────────────────────────────────────────────
  const confirm = ui.alert(
    'Split sheets by key',
    'This will move rows from "Companies" and "Censuses" into new extract ' +
    'sheets and delete them from the originals.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) {
    Logger.log('User cancelled.');
    return;
  }

  // ── 1. Load and deduplicate keys ────────────────────────────────────────────
  const keysSheet = ss.getSheetByName('Keys');
  if (!keysSheet) {
    ui.alert('Error: Sheet named "Keys" not found.');
    return;
  }

  const keysRaw = keysSheet.getDataRange().getValues();
  const keySet  = new Set();

  // Row 0 is the header; keys start at row 1 (index).
  for (let i = 1; i < keysRaw.length; i++) {
    const k = normaliseKey(keysRaw[i][0]);
    if (k !== '') keySet.add(k);
  }

  if (keySet.size === 0) {
    ui.alert('Error: No keys found in the Keys sheet.\n' +
             'Expected a header in row 1 and Company IDs from row 2 onwards.');
    return;
  }
  Logger.log('Unique keys loaded: ' + keySet.size);

  // ── 2. Locate source sheets ─────────────────────────────────────────────────
  const companiesSheet = ss.getSheetByName('Companies');
  const censusesSheet  = ss.getSheetByName('Censuses');

  if (!companiesSheet) { ui.alert('Error: Sheet named "Companies" not found.'); return; }
  if (!censusesSheet)  { ui.alert('Error: Sheet named "Censuses" not found.');  return; }

  // ── 3. Process each source sheet ────────────────────────────────────────────
  const companiesStats = extractRows(ss, companiesSheet, keySet, 'Companies Extract');
  const censusesStats  = extractRows(ss, censusesSheet,  keySet, 'Censuses Extract');

  // ── 4. Quality check and summary ────────────────────────────────────────────
  const totalBefore = companiesStats.before + censusesStats.before;
  const totalAfter  = companiesStats.kept   + companiesStats.extracted +
                      censusesStats.kept    + censusesStats.extracted;
  const pass = totalBefore === totalAfter;

  const summary =
    'Split complete.\n\n' +
    'Companies\n' +
    '  Remaining in sheet : ' + companiesStats.kept      + '\n' +
    '  Moved to extract   : ' + companiesStats.extracted + '\n\n' +
    'Censuses\n' +
    '  Remaining in sheet : ' + censusesStats.kept       + '\n' +
    '  Moved to extract   : ' + censusesStats.extracted  + '\n\n' +
    'Quality check\n' +
    '  Total rows before  : ' + totalBefore + '\n' +
    '  Total rows after   : ' + totalAfter  + '\n' +
    '  Result             : ' + (pass ? 'PASS' : 'FAIL – do NOT save; investigate first!');

  Logger.log(summary);
  ui.alert(summary);
}

// ─────────────────────────────────────────────────────────────────────────────
// Core helper
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Reads all data from sourceSheet, partitions it into rows to keep and rows
 * to extract (based on keySet matching column A), writes the extract rows to
 * a new sheet, and overwrites the source sheet with only the kept rows.
 *
 * All work is done in memory with bulk reads/writes for performance.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet}       sourceSheet
 * @param {Set<string>}                              keySet
 * @param {string}                                   extractName
 * @returns {{ before: number, kept: number, extracted: number }}
 */
function extractRows(ss, sourceSheet, keySet, extractName) {
  const allData = sourceSheet.getDataRange().getValues();

  if (allData.length === 0) {
    Logger.log(sourceSheet.getName() + ': empty, nothing to do.');
    return { before: 0, kept: 0, extracted: 0 };
  }

  const header   = allData[0];
  const keepRows = [header];   // will go back into the source sheet
  const extRows  = [header];   // will go into the extract sheet

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (keySet.has(normaliseKey(row[0]))) {
      extRows.push(row);
    } else {
      keepRows.push(row);
    }
  }

  const before    = allData.length - 1;
  const extracted = extRows.length  - 1;
  const kept      = keepRows.length - 1;

  Logger.log(sourceSheet.getName() + ': before=' + before +
             ', kept=' + kept + ', extracted=' + extracted);

  // ── Write extract sheet ──────────────────────────────────────────────────
  // Delete any pre-existing extract sheet so we start clean.
  const existing = ss.getSheetByName(extractName);
  if (existing) ss.deleteSheet(existing);

  const extSheet = ss.insertSheet(extractName);
  extSheet.getRange(1, 1, extRows.length, header.length).setValues(extRows);

  // ── Overwrite source sheet with kept rows ────────────────────────────────
  // clear() wipes content and formatting so no ghost data remains.
  sourceSheet.clear();
  if (keepRows.length > 0) {
    sourceSheet.getRange(1, 1, keepRows.length, header.length).setValues(keepRows);
  }

  return { before, kept, extracted };
}

// ─────────────────────────────────────────────────────────────────────────────
// Utility
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Converts a cell value to a trimmed string for consistent key comparison.
 * Numeric IDs (e.g. 12345) and string IDs (e.g. "12345") will match.
 *
 * @param  {*}      value
 * @returns {string}
 */
function normaliseKey(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}
