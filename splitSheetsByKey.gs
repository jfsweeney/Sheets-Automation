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
 * IMPORTANT: Make a backup copy of the spreadsheet before running.
 * The script clears and rewrites source sheets; there is no undo.
 *
 * Filters and hidden rows/columns
 * --------------------------------
 * Before processing each sheet the script removes any active filter and
 * unhides all rows and columns. This ensures the full dataset is read and
 * written — not just what is currently visible. Filters will not be present
 * on the rewritten sheets after the script completes.
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
 *   1. Make a backup copy of the spreadsheet.
 *   2. Open the Google Sheet.
 *   3. Extensions > Apps Script, paste this file, save.
 *   4. Select the splitSheetsByKey function and click Run.
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
    'IMPORTANT: Make a backup copy of this spreadsheet before continuing.\n\n' +
    'This script clears and rewrites the "Companies" and "Censuses" sheets. ' +
    'If it is interrupted after clearing but before writing back, data in ' +
    'those sheets will be lost. There is no undo.\n\n' +
    'Have you made a backup? Continue?',
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

  // Reveal all rows in the Keys sheet so hidden keys are not silently missed.
  fullyRevealSheet(keysSheet);

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

  // ── 3. Remove filters and unhide everything in source sheets ─────────────────
  // Filters or hidden rows/columns cause getDataRange() to return an incomplete
  // dataset. Revealing everything now means the full data is read, partitioned,
  // and written back — nothing is silently skipped or lost.
  fullyRevealSheet(companiesSheet);
  fullyRevealSheet(censusesSheet);

  // ── 4. Create extract sheets (replace if they already exist) ─────────────────
  const companiesExtract = prepareExtractSheet(ss, 'Companies Extract');
  const censusesExtract  = prepareExtractSheet(ss, 'Censuses Extract');

  // ── 5. Process each source sheet ────────────────────────────────────────────
  const companiesStats = extractRows(companiesSheet, keySet, companiesExtract);
  const censusesStats  = extractRows(censusesSheet,  keySet, censusesExtract);

  // ── 6. Quality check and summary ────────────────────────────────────────────
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
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Removes any active filter and shows all hidden rows and columns on a sheet.
 *
 * Filters and hidden rows/columns cause getDataRange().getValues() to return
 * an incomplete dataset on some Google Sheets versions. Revealing everything
 * before reading guarantees the script works on the full data.
 *
 * Note: filters are not restored after the script runs. The rewritten source
 * sheets will have no filter applied.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function fullyRevealSheet(sheet) {
  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
    Logger.log(sheet.getName() + ': filter removed.');
  }
  sheet.showRows(1, sheet.getMaxRows());
  sheet.showColumns(1, sheet.getMaxColumns());
  Logger.log(sheet.getName() + ': all rows and columns unhidden.');
}

/**
 * Deletes any existing sheet with the given name and creates a fresh one.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function prepareExtractSheet(ss, name) {
  const existing = ss.getSheetByName(name);
  if (existing) ss.deleteSheet(existing);
  return ss.insertSheet(name);
}

/**
 * Reads all data from sourceSheet, partitions it into rows to keep and rows
 * to extract (based on keySet matching column A), writes the extract rows to
 * extractSheet, and overwrites sourceSheet with only the kept rows.
 *
 * All work is done in memory with bulk reads/writes for performance.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet
 * @param {Set<string>}                        keySet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} extractSheet
 * @returns {{ before: number, kept: number, extracted: number }}
 */
function extractRows(sourceSheet, keySet, extractSheet) {
  const allData = sourceSheet.getDataRange().getValues();

  // A sheet with no data rows (empty or header-only) has nothing to partition.
  if (allData.length <= 1) {
    Logger.log(sourceSheet.getName() + ': no data rows, nothing to do.');
    return { before: 0, kept: 0, extracted: 0 };
  }

  const header   = allData[0];
  const keepRows = [header];   // will go back into the source sheet
  const extRows  = [header];   // will go into the extract sheet

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const key = normaliseKey(row[0]);

    // Rows with no Company ID are always kept in the source sheet.
    if (key === '') {
      keepRows.push(row);
      continue;
    }

    if (keySet.has(key)) {
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

  // Write extract data (header is always row 1).
  extractSheet.getRange(1, 1, extRows.length, header.length).setValues(extRows);

  // Overwrite source sheet with kept rows.
  // clear() wipes content and formatting so no ghost data remains after
  // the row count shrinks.
  sourceSheet.clear();
  sourceSheet.getRange(1, 1, keepRows.length, header.length).setValues(keepRows);

  return { before, kept, extracted };
}

// ─────────────────────────────────────────────────────────────────────────────
// Utility
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Converts a cell value to a trimmed string for consistent key comparison.
 * Alphanumeric IDs stored as text or small integers are handled correctly.
 *
 * @param  {*}      value
 * @returns {string}
 */
function normaliseKey(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}
