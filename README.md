# Sheets Automation — Carve-Off by Key

This script performs a **carve-off** on a Google Sheets workbook — moving a specified set of company records out of the main data sheets and into separate carve-off sheets.

Specifically, it reads a list of Company IDs you provide and moves all matching rows from the **Companies** and **Censuses** worksheets into two new worksheets called **Companies Carve-Off** and **Censuses Carve-Off**. Those rows are then deleted from the originals. Everything else is left untouched.

---

## Background

In group benefits renewal quoting for a Professional Employer Organization (PEO), the full block of business often needs to be segmented into tranches — based on criteria such as group size, industry, renewal date, or claims history — for separate quoting treatment. Separating a subset of groups from the main block in this way is called a **carve-off** in the industry.

This script automates that step. Once it completes, the workbook contains two clean, separated datasets: the carved-off groups and the remainder. Each can then be exported to its own file and run independently through the quoting decision intelligence engine.

---

## Before you start

### Your file must be in Google Sheets format — not XLSX

Google Sheets does not allow scripts to run on `.xlsx` files stored in Drive. The **Extensions** menu (where you install and run the script) will not appear unless the file has been converted to native Google Sheets format first.

**To convert:**

1. Open the `.xlsx` file in Google Drive (double-click it).
2. Go to **File > Save as Google Sheets**.
3. A new, converted copy opens in a new tab. Work from this copy — your original `.xlsx` file is left unchanged in Drive.

### Make a backup

The script rewrites the Companies and Censuses sheets in place. Before running it, make a copy of the converted file as a backup:

**File > Make a copy**

Keep this backup until you are satisfied with the results.

### Filters and hidden rows or columns

If the **Companies** or **Censuses** sheets have a filter applied or have hidden rows or columns, **the script will remove the filter and unhide everything automatically** before reading the data. This is necessary to ensure no rows are silently skipped.

Be aware of two consequences:
- **Filters on Companies and Censuses will not be restored** on the rewritten sheets after the script runs.
- **All previously hidden rows will be visible** in the resulting sheets.

If you need your filters preserved, note them down before running so you can reapply them afterwards.

The **Keys** sheet filter is left in place — the script reads Keys without removing its filter, and never rewrites it.

---

## What your workbook needs to contain

The script expects these sheet tabs to exist (the names must match exactly, including capitalisation):

| Sheet name | What it contains |
|---|---|
| **Keys** | The list of Company IDs you want to carve off. Row 1 should be a header (e.g. "Company ID"). The IDs themselves start in row 2, one per row in column A. Duplicates in this list are fine — they are handled automatically. |
| **Companies** | Your companies data. Row 1 must be a header row. Column A must be the Company ID. |
| **Censuses** | Your census data. Row 1 must be a header row. Column A must be the Company ID. |

Any other sheets in the workbook are ignored.

---

## Step-by-step instructions

### 1. Add the Keys sheet

In your converted Google Sheets file:

1. Click the **+** button at the bottom left to add a new sheet.
2. Rename it `Keys` (right-click the tab > **Rename**).
3. In cell A1, type a header such as `Company ID`.
4. Paste your list of Company IDs into column A starting at row 2, one ID per row.

### 2. Open the script editor

1. In the menu bar, click **Extensions > Apps Script**.
2. A new browser tab opens showing the script editor. You may see a default file called `Code.gs` with some placeholder text — that is fine.

### 3. Paste the script

1. Open the file `splitSheetsByKey.gs` from this repository and copy its entire contents.
2. Back in the Apps Script tab, select all the existing text in the editor and delete it.
3. Paste the copied script in its place.
4. Click the **Save** icon (floppy disk) or press **Ctrl+S** / **Cmd+S**.

### 4. Run the script

1. In the function dropdown at the top of the editor (it may say `myFunction` or show a function name), select **splitSheetsByKey**.
2. Click the **Run** button (the triangle/play icon).

**First run only:** Google will ask you to review and grant permissions. Click **Review permissions**, choose your Google account, and click **Allow**. This lets the script read and modify the spreadsheet.

### 5. Confirm the backup warning

A dialog box will appear reminding you to have a backup and explaining that the operation cannot be undone. Click **Yes** to proceed.

### 6. Wait for it to finish

The script may take up to a minute or two to process large sheets. Do not close the browser tab while it is running. When it is done, a results dialog will appear.

---

## Reading the results

When the script completes, a summary dialog reports how many rows were moved and includes a quality check:

```
Carve-off complete.

Companies
  Remaining in sheet : 1842
  Moved to carve-off : 190

Censuses
  Remaining in sheet : 34284
  Moved to carve-off : 716

Quality check (populated rows)
  Total rows before  : 36332
  Total rows after   : 36332
  Result             : PASS
```

- **PASS** means the total populated row count is consistent — no data rows were created or lost.
- **FAIL** means something unexpected happened. Do not save or close the file. Check the original backup and contact whoever set up the script.

The script also runs an independent integrity check before touching any data. If the number of rows returned by the API does not match what the sheet reports, the script stops immediately and displays an error. In that case nothing will have been modified.

---

## What the output looks like

After a successful run your workbook will have (at minimum) these sheets:

| Sheet | Contents |
|---|---|
| **Companies** | Original data minus the carved-off rows |
| **Censuses** | Original data minus the carved-off rows |
| **Companies Carve-Off** | Header row from Companies + all carved-off company rows |
| **Censuses Carve-Off** | Header row from Censuses + all carved-off census rows |

If **Companies Carve-Off** or **Censuses Carve-Off** sheets already existed when you ran the script, they were replaced entirely.

---

## Troubleshooting

**The Extensions menu is greyed out or missing.**
The file is still in XLSX format. Follow the conversion steps at the top of this guide.

**"Sheet named 'Keys' not found" / "Sheet named 'Companies' not found"**
The sheet tab name does not match exactly. Check for extra spaces, different capitalisation, or a typo. The names must be `Keys`, `Companies`, and `Censuses`.

**"No keys found in the Keys sheet."**
The Keys sheet is empty, or the IDs start in row 1 instead of row 2. Row 1 is treated as a header and is skipped — your IDs must start in row 2.

**The script asked for permissions but I clicked Deny.**
Close the Apps Script tab, reopen it via **Extensions > Apps Script**, and run the function again. You must grant permissions for the script to access the spreadsheet.

**The quality check says FAIL.**
Do not save the file. Close it without saving and reopen your backup copy. Then contact whoever set up the script with a description of what you saw.
