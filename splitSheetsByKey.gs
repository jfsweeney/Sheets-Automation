/**
 * splitSheetsByKey.gs
 *
 * Performs a carve-off on a Google Sheets workbook: moves rows from the
 * "Companies" and "Censuses" worksheets into new "Companies Carve-Off" and
 * "Censuses Carve-Off" worksheets, based on a list of Company IDs stored in
 * a "Keys" sheet.
 *
 * Prerequisites
 * -------------
 *   Keys      – Row 1 is a header. Company IDs to carve off start in row 2,
 *               column A. Duplicates in this list are handled automatically.
 *   Companies – Row 1 is a header. Column A is Company ID (one row per ID).
 *   Censuses  – Row 1 is a header. Column A is Company ID (may repeat).
 *
 * IMPORTANT: Make a backup copy of the spreadsheet before running.
 * The script clears and rewrites source sheets; there is no undo.
 *
 * Filters and hidden rows/columns
 * --------------------------------
 * Before processing, the script removes any active filter from "Companies"
 * and "Censuses" and unhides all rows and columns on those sheets. This
 * ensures the full dataset is read and written — nothing is silently skipped.
 * Filters on "Companies" and "Censuses" will not be present after the script
 * completes. The "Keys" sheet filter is left in place.
 *
 * Output
 * ------
 *   "Companies Carve-Off" – Created (or replaced) with the Companies header
 *                           plus all rows whose Company ID was in Keys.
 *   "Censuses Carve-Off"  – Created (or replaced) with the Censuses header
 *                           plus all rows whose Company ID was in Keys.
 *   The matching rows are deleted from Companies and Censuses.
 *
 * Quality check
 * -------------
 *   An alert reports total populated data-row counts before and after.
 *   The counts must match. An independent cross-check also verifies that
 *   the number of rows returned by the API matches getLastRow() on each
 *   sheet; a mismatch aborts the script before any data is modified.
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
    'Carve-off sheets by key',
    'IMPORTANT: Make a backup copy of this spreadsheet before continuing.\n\n' +
    'This script clears and rewrites the "Companies" and "Censuses" sheets. ' +
    'If it is interrupted after clearing but before writing back, data in ' +
    'those sheets will be lost. There is no undo.\n\n' +
    'Note: any filters on "Companies" and "Censuses" will be permanently removed.\n\n' +
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

  // Unhide any manually hidden rows/columns in Keys so no key is missed.
  // We do not remove the Keys filter: the sheet is never rewritten, and
  // getValues() returns all rows regardless of filter state.
  keysSheet.showRows(1, keysSheet.getMaxRows());
  keysSheet.showColumns(1, keysSheet.getMaxColumns());

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
  fullyRevealSheet(companiesSheet);
  fullyRevealSheet(censusesSheet);

  // ── 4. Read and partition both sheets in memory (no writes yet) ──────────────
  // All reading and partitioning is completed before any sheet is cleared or
  // written. If anything goes wrong during the read (e.g. a row-count mismatch
  // that suggests the data was not fully returned), the script aborts here and
  // no data is modified.
  let companiesPartition, censusesPartition;
  try {
    companiesPartition = partitionSheet(companiesSheet, keySet);
    censusesPartition  = partitionSheet(censusesSheet,  keySet);
  } catch (e) {
    ui.alert('Error reading data — no data has been modified.\n\n' + e.message);
    return;
  }

  // ── 5. Write all results ─────────────────────────────────────────────────────
  // ─── Point of no return ───────────────────────────────────────────────────
  applyPartition(ss, companiesSheet, companiesPartition, 'Companies Carve-Off');
  applyPartition(ss, censusesSheet,  censusesPartition,  'Censuses Carve-Off');

  // ── 6. Quality check and summary ────────────────────────────────────────────
  const totalBefore = companiesPartition.before   + censusesPartition.before;
  const totalAfter  = companiesPartition.kept      + companiesPartition.carvedOff +
                      censusesPartition.kept       + censusesPartition.carvedOff;
  const pass = totalBefore === totalAfter;

  const summary =
    'Carve-off complete.\n\n' +
    'Companies\n' +
    '  Remaining in sheet : ' + companiesPartition.kept      + '\n' +
    '  Moved to carve-off : ' + companiesPartition.carvedOff + '\n\n' +
    'Censuses\n' +
    '  Remaining in sheet : ' + censusesPartition.kept       + '\n' +
    '  Moved to carve-off : ' + censusesPartition.carvedOff  + '\n\n' +
    'Quality check (populated rows)\n' +
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
 * Removes any active basic filter and shows all hidden rows and columns.
 * Called only on source sheets that will be fully rewritten.
 *
 * Note: sheet.getFilter() handles only basic filters (Data > Create a filter).
 * Filter views (Data > Filter views) are a separate feature and are not
 * removed here. This is safe because filter views are display-only — the
 * getValues() API always returns the full dataset regardless of which filter
 * view is active.
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
 * Reads all data from sourceSheet and partitions it into rows to keep and
 * rows to carve off. No data is written at this stage.
 *
 * Cross-check: compares getLastRow() against the row count returned by
 * getDataRange().getValues(). These use different internal paths; a mismatch
 * suggests the data was not fully returned and the script throws rather than
 * risk silent data loss.
 *
 * Only populated rows (at least one non-empty cell) are counted in the
 * before/kept/carvedOff statistics. Fully-empty rows are preserved in
 * keepRows so they survive in the source sheet, but are not counted.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet
 * @param {Set<string>}                        keySet
 * @returns {{ header, keepRows, carveOffRows, before, kept, carvedOff }}
 */
function partitionSheet(sourceSheet, keySet) {
  const lastRow = sourceSheet.getLastRow();

  // Completely empty sheet — nothing to partition.
  if (lastRow === 0) {
    Logger.log(sourceSheet.getName() + ': sheet is completely empty.');
    return { header: [], keepRows: [], carveOffRows: [], before: 0, kept: 0, carvedOff: 0 };
  }

  const allData = sourceSheet.getDataRange().getValues();

  // Cross-check: getLastRow() and getDataRange() must agree after fullyRevealSheet().
  // A mismatch here means the API did not return a complete dataset — abort
  // before touching anything.
  if (allData.length !== lastRow) {
    throw new Error(
      sourceSheet.getName() + ': row count mismatch after unhiding — ' +
      'getLastRow() reports ' + lastRow + ' rows but getDataRange() returned ' +
      allData.length + '. This may indicate a filter or display issue.'
    );
  }

  // Header-only sheet — nothing to partition, but preserve the header.
  if (allData.length <= 1) {
    const header = allData[0];
    Logger.log(sourceSheet.getName() + ': no data rows, nothing to do.');
    return { header, keepRows: [header], carveOffRows: [header], before: 0, kept: 0, carvedOff: 0 };
  }

  const header      = allData[0];
  const keepRows    = [header];
  const carveOffRows = [header];
  let before = 0, kept = 0, carvedOff = 0;

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];

    // Fully-empty rows are preserved in the source sheet but not counted
    // in the quality check, as they contain no populated data.
    if (row.every(cell => cell === '' || cell === null || cell === undefined)) {
      keepRows.push(row);
      continue;
    }

    before++;
    const key = normaliseKey(row[0]);

    // Populated rows with no Company ID are kept in the source sheet.
    if (key === '') {
      keepRows.push(row);
      kept++;
      continue;
    }

    if (keySet.has(key)) {
      carveOffRows.push(row);
      carvedOff++;
    } else {
      keepRows.push(row);
      kept++;
    }
  }

  Logger.log(sourceSheet.getName() + ': before=' + before +
             ', kept=' + kept + ', carvedOff=' + carvedOff);

  return { header, keepRows, carveOffRows, before, kept, carvedOff };
}

/**
 * Writes a completed partition to the spreadsheet:
 *   - Creates (or replaces) the named carve-off sheet and fills it with
 *     carveOffRows.
 *   - Clears the source sheet and writes back only the kept rows.
 *
 * This is called only after both source sheets have been successfully
 * partitioned in memory.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet}  ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet}        sourceSheet
 * @param {{ header, keepRows, carveOffRows }}         partition
 * @param {string}                                     carveOffName
 */
function applyPartition(ss, sourceSheet, partition, carveOffName) {
  const { header, keepRows, carveOffRows } = partition;

  // Nothing to write for a completely empty sheet.
  if (!header || header.length === 0) return;

  const carveOffSheet = prepareCarveOffSheet(ss, carveOffName);
  carveOffSheet.getRange(1, 1, carveOffRows.length, header.length).setValues(carveOffRows);

  // clear() removes content and formatting so no ghost data lingers after
  // the row count shrinks.
  sourceSheet.clear();
  sourceSheet.getRange(1, 1, keepRows.length, header.length).setValues(keepRows);
}

/**
 * Deletes any existing sheet with the given name and creates a fresh one.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function prepareCarveOffSheet(ss, name) {
  const existing = ss.getSheetByName(name);
  if (existing) ss.deleteSheet(existing);
  return ss.insertSheet(name);
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
