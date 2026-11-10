function followUp() {
  // Get the active spreadsheet and select the sheet by name
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Follow-Up");
  if (!sheet) return; // Exit if the sheet is not found

  // Find the last row with data
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // Exit if there are no data rows beyond the header

  // Get all values in column D (starting from row 2 to the last row)
  const columnD = sheet.getRange("D2:D" + lastRow).getValues();
  const columnE = []; // Prepare an array for the values to write into column E

  // Loop through each row in column D
  for (let i = 0; i < columnD.length; i++) {
    const rawText = columnD[i][0]; // Current cell value in column D
    const text = normalizeText(rawText || ""); // Normalize text (lowercase, no accents)

    // Check the normalized text against known keywords/phrases
    if (text.includes("contract closed")) {
      columnE.push(["Contract Closed"]);
    } else if (text.includes("no response")) {
      columnE.push(["Contact made - no response"]);
    } else if (text.includes("doing it themselves")) {
      columnE.push(["Doing Independently"]);
    } else if (text.includes("contract sent")) {
      columnE.push(["Contract Sent"]);
    } else if (text.includes("other")) {
      columnE.push(["Other"]);
    } else if (text.includes("contract cancelled")) {
      columnE.push(["Contract Cancelled"]);
    } else if (text.includes("business plus client, not starter")) {
      columnE.push(["Business Plus Client but not Starter Client"]);
    } else {
      // If none of the keywords match, leave cell blank
      columnE.push([""]);
    }
  }

  // Write the results into column E (row 2 through last row)
  sheet.getRange("E2:E" + lastRow).setValues(columnE);
}

function normalizeText(text) {
  return text
    .toString()             // Ensure the value is a string
    .toLowerCase()          // Convert to lowercase
    .normalize("NFD")       // Normalize characters (separate base + accent)
    .replace(/[̀-ͯ]/g, ""); // Remove accents/diacritics
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function followUp2() {
  // Get the active spreadsheet and select the sheet by name
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Follow-Up 2");
  if (!sheet) return; // Exit if the sheet is not found

  // Get all values from column E starting at row 2
  const colE = sheet.getRange("E2:E").getValues();

  // Find the first empty row in column E (rowCount = index of first blank + 1)
  // If no empty cell is found, default to the full column length
  const rowCount = colE.findIndex(row => row[0] === "") + 1 || colE.length;

  if (rowCount < 2) return; // Exit if there are no valid rows

  // Get all values from column E, limited to the rows with data
  const columnE = sheet.getRange("E2:E" + (rowCount + 1)).getValues();
  const columnF = []; // Prepare an array to hold the results for column F

  // Loop through each row in column E
  for (let i = 0; i < columnE.length; i++) {
    const rawText = columnE[i][0]; // Current cell value in column E
    const text = normalizeText(rawText || ""); // Normalize text (lowercase, no accents)

    // Match normalized text to specific categories and push results into columnF
    if (text.includes("contract closed")) {
      columnF.push(["Contract Closed"]);
    } else if (text.includes("no response")) {
      columnF.push(["Contact made - no response"]);
    } else if (text.includes("doing it themselves")) {
      columnF.push(["Doing Independently"]);
    } else if (text.includes("contract sent")) {
      columnF.push(["Contract Sent"]);
    } else if (text.includes("other")) {
      columnF.push(["Other"]);
    } else if (text.includes("contract cancelled")) {
      columnF.push(["Contract Cancelled"]);
    } else if (text.includes("business plus client, not starter")) {
      columnF.push(["Business Plus Client but not Starter Client"]);
    } else {
      // If no match, leave the corresponding F cell blank
      columnF.push([""]);
    }
  }

  // Write the results into column F (row 2 down to the last processed row)
  sheet.getRange("F2:F" + (columnF.length + 1)).setValues(columnF);
}

function normalizeText(text) {
  return text
    .toString()             // Ensure value is treated as string
    .toLowerCase()          // Convert all characters to lowercase
    .normalize("NFD")       // Decompose accented characters into base + accent
    .replace(/[̀-ͯ]/g, ""); // Remove accents
}


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function updateBusinessPlus() {
  try {
    // IDs and sheet names for source and destination
    const SRC_ID = 'YOUR_SOURCE_SPREADSHEET_ID';
    const SRC_SHEET = 'Business Plus';
    const DST_SHEET = 'Business Plus Updated';

    // Open source and destination sheets
    const src = SpreadsheetApp.openById(SRC_ID).getSheetByName(SRC_SHEET);
    const dst = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DST_SHEET);
    if (!src) throw new Error(`Source tab "${SRC_SHEET}" not found.`);
    if (!dst) throw new Error(`Destination tab "${DST_SHEET}" not found.`);

    // --- Read data from source sheet ---

    // Last row with data in source sheet
    const lastSrcRow = src.getLastRow();
    if (lastSrcRow < 2) return; // Exit if no data rows beyond the header

    const numRows = lastSrcRow - 1; // Number of rows of actual data

    // Get values from column A (names)
    const names = src.getRange(2, 1, numRows, 1).getValues();

    // Get values from column O (dates)
    const oVals = src.getRange(2, 15, numRows, 1).getValues();          // raw values
    const oDisp = src.getRange(2, 15, numRows, 1).getDisplayValues();   // displayed text (mm/dd/yyyy)

    // --- Build a set of existing keys in destination to avoid duplicates ---
    const existing = buildExistingKeySet(dst);

    const toAppend = [];
    for (let i = 0; i < numRows; i++) {
      const name = String(names[i][0] || '').trim(); // Current name
      if (!name) continue; // Skip blank names

      // Parse the date from column O (raw value or displayed string)
      const parsed = coerceDateMMDDYYYY(oVals[i][0], oDisp[i][0]);
      if (!parsed) continue; // Skip invalid or missing dates

      const day = localDateOnly(parsed); // Normalize to date-only (no time)
      const key = dedupeKey(name, day);  // Unique key for name+date
      if (!existing.has(key)) {
        // Only append if this name+date is not already in destination
        existing.add(key);
        toAppend.push([name, day]);
      }
    }

    if (!toAppend.length) return; // Nothing new to add

    // --- Append to destination sheet ---

    // Find the last filled row in column A of destination (ignores formatting/empty formulas)
    const lastFilled = getLastFilledRow(dst, 1, 1);
    const startRow = (lastFilled || 1) + 1;

    // Write new data to destination (columns A:B)
    dst.getRange(startRow, 1, toAppend.length, 2).setValues(toAppend);

    // Format column B as mm/dd/yyyy
    dst.getRange(startRow, 2, toAppend.length, 1).setNumberFormat('mm/dd/yyyy');
    SpreadsheetApp.flush(); // Apply all pending changes
  } catch (err) {
    // Show error as toast notification in the spreadsheet
    SpreadsheetApp.getActive().toast(`Error: ${err.message}`);
    throw err; // Rethrow for debugging/logging
  }
}

/** Helpers **/

// Return a date with time zeroed out (date-only)
function localDateOnly(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
}

// Parse a date from raw value or display text in mm/dd/yyyy format
function coerceDateMMDDYYYY(val, displayText) {
  if (val instanceof Date && !isNaN(val)) return val; // Use actual Date if valid
  const s = String(displayText || '').trim();
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/); // Match mm/dd/yyyy or mm-dd-yyyy
  if (!m) return null;
  const mm = +m[1], dd = +m[2], yy = +(m[3].length === 2 ? '20' + m[3] : m[3]);
  const d = new Date(yy, mm - 1, dd, 0, 0, 0, 0);
  return isNaN(d) ? null : d;
}

// Find the last row in the given column range that actually has content.
// Ignores formatting, borders, validations, and formulas that return "".
function getLastFilledRow(sheet, startCol, endCol) {
  const lr = sheet.getLastRow(); // Upper bound
  if (lr < 2) return 1;          // Only header exists
  const width = (endCol - startCol + 1);
  const values = sheet.getRange(2, startCol, lr - 1, width).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const rowHasData = values[i].some(v => String(v).trim() !== '');
    if (rowHasData) return i + 2; // Offset since data starts at row 2
  }
  return 1; // No data found beyond header
}

// Build a set of existing keys (name + date) from destination sheet to prevent duplicates.
function buildExistingKeySet(sheet) {
  const last = getLastFilledRow(sheet, 1, 2); // Look at columns A:B
  const set = new Set();
  if (last < 2) return set; // No data

  const vals = sheet.getRange(2, 1, last - 1, 2).getValues();
  const disp = sheet.getRange(2, 1, last - 1, 2).getDisplayValues();
  for (let i = 0; i < vals.length; i++) {
    const name = String(vals[i][0] || '').trim();
    if (!name) continue;

    let date = vals[i][1];
    if (!(date instanceof Date) || isNaN(date)) {
      // Fallback: try parsing displayed text if date is not valid
      date = coerceDateMMDDYYYY(vals[i][1], disp[i][1]);
      if (!date) continue;
    }
    set.add(dedupeKey(name, localDateOnly(date)));
  }
  return set;
}

// Create a unique key by combining name and date (YYYY-MM-DD)
function dedupeKey(name, date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${name}|${y}-${m}-${d}`;
}


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function syncBusinessPlusFields() {
  const SRC_SHEET = 'Business Plus Details'; // Source sheet name
  const DST_SHEET = 'Business Plus Updated'; // Destination sheet name

  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SRC_SHEET);
  const dst = ss.getSheetByName(DST_SHEET);
  if (!src) throw new Error(`Source sheet "${SRC_SHEET}" not found`);
  if (!dst) throw new Error(`Destination sheet "${DST_SHEET}" not found`);

  // Find the last filled rows in column A of both sheets
  const srcLast = getLastFilledRowInColumn(src, 1); // Col A in source
  const dstLast = getLastFilledRowInColumn(dst, 1); // Col A in destination
  if (srcLast < 2 || dstLast < 2) return; // Exit if no usable data

  // --- Read source data (columns A..E) ---
  const srcRows = srcLast - 1;
  const srcData = src.getRange(2, 1, srcRows, 5).getValues(); // A..E

  // Preprocess source data:
  // - Tokenize names (col A)
  // - Store tokens + columns C, D, E values
  const srcItems = [];
  for (const row of srcData) {
    const name = String(row[0] || '').trim();
    if (!name) continue;
    const tokens = new Set(tokenizeName(name));
    if (tokens.size === 0) continue;
    srcItems.push({ tokens, cde: [row[2], row[3], row[4]], tokenCount: tokens.size });
  }

  // --- Read destination data ---
  const dstRows = dstLast - 1;
  const dstNames = dst.getRange(2, 1, dstRows, 1).getValues();  // Col A (names)
  const dstCDE   = dst.getRange(2, 3, dstRows, 3).getValues();  // Cols C..E (to update)

  let updates = 0;
  for (let i = 0; i < dstRows; i++) {
    const raw = String(dstNames[i][0] || '').trim();
    if (!raw) continue;

    const tokensArr = tokenizeName(raw);
    if (tokensArr.length === 0) continue;
    const destTokens = new Set(tokensArr);

    // Try to find the "best match" source row:
    // Destination tokens must be a subset of source tokens
    // Choose the source row with the largest token count (most specific match)
    let best = null, bestSize = -1;
    for (const item of srcItems) {
      if (isSubset(destTokens, item.tokens)) {
        if (item.tokenCount > bestSize) {
          best = item;
          bestSize = item.tokenCount;
        }
      }
    }

    // If a best match is found, copy C..E values from source into destination
    if (best) {
      dstCDE[i][0] = best.cde[0];
      dstCDE[i][1] = best.cde[1];
      dstCDE[i][2] = best.cde[2];
      updates++;
    }
  }

  // Write updates back to destination if changes were made
  if (updates > 0) {
    dst.getRange(2, 3, dstRows, 3).setValues(dstCDE);
  }
  // Optional toast notification: SpreadsheetApp.getActive().toast(`Updated ${updates} row(s).`);
}

/** Helpers **/

// Find the last visually non-empty row in a column (ignores formatting/blank formulas)
function getLastFilledRowInColumn(sheet, col) {
  const lr = sheet.getLastRow();
  if (lr < 2) return 1;
  const disp = sheet.getRange(2, col, lr - 1, 1).getDisplayValues();
  for (let i = disp.length - 1; i >= 0; i--) {
    if (String(disp[i][0]).trim() !== '') return i + 2;
  }
  return 1;
}

// Tokenize a name:
// - Remove accents
// - Convert to lowercase
// - Split into words (letters only)
// - Drop short tokens and common name particles/connectors
function tokenizeName(s) {
  if (!s) return [];
  const STOP = new Set(['de','la','van','von','der','and']); // common connectors
  const ascii = String(s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '') // strip accents
    .toLowerCase();
  const words = ascii.match(/[a-z]+/g) || []; // extract only a–z sequences
  return words.filter(w => w.length > 1 && !STOP.has(w));
}

// Return true if every token in setA is present in setB
function isSubset(setA, setB) {
  for (const t of setA) if (!setB.has(t)) return false;
  return true;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function syncAnnualReviewFromBusinessPlusUpdated() {
  const SRC_SHEET = 'Business Plus Updated'; // source
  const DST_SHEET = 'Annual Review Updated'; // destination

  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SRC_SHEET);
  const dst = ss.getSheetByName(DST_SHEET);
  if (!src) throw new Error(`Source sheet "${SRC_SHEET}" not found`);
  if (!dst) throw new Error(`Destination sheet "${DST_SHEET}" not found`);

  // Last actually-filled rows (scan col A, ignore formatting/blank formulas)
  const srcLast = getLastFilledRowInColumn(src, 1);
  if (srcLast < 2) return; // nothing to sync
  const dstLast = getLastFilledRowInColumn(dst, 1);

  // Source block A..E
  const srcRows = srcLast - 1;
  const srcData = src.getRange(2, 1, srcRows, 5).getValues(); // [A,B,C,D,E]

  // Map: key(name) -> full source row
  const srcByName = new Map();
  for (let i = 0; i < srcRows; i++) {
    const row = srcData[i];
    const key = normalizeName(row[0]);
    if (!key) continue;
    srcByName.set(key, row); // last wins
  }

  // Destination block A..I (so we can read C flag and write A,B,F,G,I)
  let dstRows = 0, dstBlock = [];
  if (dstLast >= 2) {
    dstRows  = dstLast - 1;
    dstBlock = dst.getRange(2, 1, dstRows, 9).getValues(); // A..I
  }

  // Existing dest keys (names)
  const dstKeys = new Set();
  for (let i = 0; i < dstRows; i++) {
    const key = normalizeName(dstBlock[i][0]); // A
    if (key) dstKeys.add(key);
  }

  // Prepare column-wise arrays (copy existing as baseline)
  const newA = Array.from({length: dstRows}, (_, i) => [dstBlock[i][0]]); // A
  const newB = Array.from({length: dstRows}, (_, i) => [dstBlock[i][1]]); // B
  const newF = Array.from({length: dstRows}, (_, i) => [dstBlock[i][5]]); // F
  const newG = Array.from({length: dstRows}, (_, i) => [dstBlock[i][6]]); // G
  const newI = Array.from({length: dstRows}, (_, i) => [dstBlock[i][8]]); // I

  let changed = 0;

  // UPDATE existing rows where C (flag) != "Yes"
  for (let i = 0; i < dstRows; i++) {
    const key = normalizeName(dstBlock[i][0]);         // name in A
    if (!key) continue;

    const cFlag = String(dstBlock[i][2] ?? '').trim().toLowerCase(); // C
    if (cFlag === 'yes') continue; // do not touch

    const srcRow = srcByName.get(key);
    if (!srcRow) continue;

    // Mapping: src [A,B,C,D,E] -> dst [A,B,F,G,I]
    if (!valuesEqual(newA[i][0], srcRow[0])) { newA[i][0] = srcRow[0]; changed++; }
    if (!valuesEqual(newB[i][0], srcRow[1])) { newB[i][0] = srcRow[1]; changed++; }
    if (!valuesEqual(newF[i][0], srcRow[2])) { newF[i][0] = srcRow[2]; changed++; }
    if (!valuesEqual(newG[i][0], srcRow[3])) { newG[i][0] = srcRow[3]; changed++; }
    if (!valuesEqual(newI[i][0], srcRow[4])) { newI[i][0] = srcRow[4]; changed++; }
  }

  // Write column batches if anything changed
  if (dstRows > 0 && changed > 0) {
    dst.getRange(2, 1, dstRows, 1).setValues(newA); // A
    dst.getRange(2, 2, dstRows, 1).setValues(newB); // B
    dst.getRange(2, 6, dstRows, 1).setValues(newF); // F
    dst.getRange(2, 7, dstRows, 1).setValues(newG); // G
    dst.getRange(2, 9, dstRows, 1).setValues(newI); // I
  }

  // APPEND any new names that don't yet exist in destination
  const toAppend = [];
  for (const [key, row] of srcByName.entries()) {
    if (!dstKeys.has(key)) {
      // Build a 9-cell row A..I
      // Fill A,B,F,G,I from src A,B,C,D,E; leave C,D,E,H blank
      toAppend.push([
        row[0],      // A
        row[1],      // B
        '',          // C (flag; left blank)
        '',          // D
        '',          // E
        row[2],      // F (src C)
        row[3],      // G (src D)
        '',          // H
        row[4]       // I (src E)
      ]);
    }
  }

  if (toAppend.length > 0) {
    const startRow = (dstLast < 2 ? 2 : dstLast + 1);
    dst.getRange(startRow, 1, toAppend.length, 9).setValues(toAppend);
  }

  // Optional toast:
  // SpreadsheetApp.getActive().toast(`Updated ${changed} cell(s), appended ${toAppend.length} row(s).`);
}

/** Helpers **/

// Last visually non-empty row in column (ignores formatting/blank formulas)
function getLastFilledRowInColumn(sheet, col) {
  const lr = sheet.getLastRow();
  if (lr < 2) return 1;
  const disp = sheet.getRange(2, col, lr - 1, 1).getDisplayValues();
  for (let i = disp.length - 1; i >= 0; i--) {
    if (String(disp[i][0]).trim() !== '') return i + 2;
  }
  return 1;
}

// Normalize name key (case-insensitive, trimmed)
function normalizeName(v) {
  return String(v || '').trim().toLowerCase();
}

// Compare values incl. Date equality by epoch
function valuesEqual(a, b) {
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  return String(a) === String(b);
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function importNoFromBusinessPlus() {
  const SRC_ID = 'YOUR_SOURCE_SPREADSHEET_ID'; // external spreadsheet
  const SRC_SHEET = 'Business Plus';
  const DST_SHEET = 'Annual Review Updated';

  const srcSS = SpreadsheetApp.openById(SRC_ID);
  const src = srcSS.getSheetByName(SRC_SHEET);
  const dst = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DST_SHEET);
  if (!src) throw new Error(`Source sheet "${SRC_SHEET}" not found in ${SRC_ID}`);
  if (!dst) throw new Error(`Destination sheet "${DST_SHEET}" not found in active spreadsheet`);

  // ---- Read source (Business Plus) ----
  const srcLast = getLastFilledRowInColumn(src, 1); // scan col A
  if (srcLast < 2) return; // nothing to do
  const srcRows = srcLast - 1;

  // Read A..I so we can grab A, B, C, E, I
  const srcData = src.getRange(2, 1, srcRows, 9).getValues(); // [A..I]

  // ---- Destination names to avoid duplicates ----
  const dstLast = getLastFilledRowInColumn(dst, 1); // scan col A
  const dstRows = Math.max(0, dstLast - 1);
  const destNameSet = new Set();
  if (dstRows > 0) {
    const dstNames = dst.getRange(2, 1, dstRows, 1).getValues();
    for (const [name] of dstNames) {
      const key = norm(name);
      if (key) destNameSet.add(key);
    }
  }

  // ---- Build rows to append (A->A, C->B, E->C, "Yes"->D, I->E) when B == "No" ----
  const toAppend = [];
  for (const row of srcData) {
    const name = row[0];     // Business Plus A
    const flag = row[1];     // Business Plus B
    const colC = row[2];     // Business Plus C
    const colE = row[4];     // Business Plus E
    const colI = row[8];     // Business Plus I

    if (norm(flag) !== 'no') continue;         // only B == "No"
    const key = norm(name);
    if (!key) continue;                         // skip blank names
    if (destNameSet.has(key)) continue;         // already present

    // Map into Annual Review Updated A..E
    toAppend.push([name, colC, colE, 'Yes', colI]);
    destNameSet.add(key); // avoid dupes within this run
  }

  if (toAppend.length === 0) return;

  // ---- Append below last actually filled row (ignores formatting/"" formulas) ----
  const startRow = (dstLast < 2 ? 2 : dstLast + 1);
  dst.getRange(startRow, 1, toAppend.length, 5).setValues(toAppend);
}

/* ================= Helpers ================= */

// Last visually non-empty row in a column (ignores formatting/blank formulas)
function getLastFilledRowInColumn(sheet, col) {
  const lr = sheet.getLastRow();
  if (lr < 2) return 1;
  const disp = sheet.getRange(2, col, lr - 1, 1).getDisplayValues();
  for (let i = disp.length - 1; i >= 0; i--) {
    if (String(disp[i][0]).trim() !== '') return i + 2;
  }
  return 1;
}

// Normalize text: trim, lowercase, remove accents (so "Café" -> "cafe")
function norm(v) {
  return String(v || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().trim();
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/**
 * Nightly sync: Source file "Quick Check" -> Destination tab "Quick Check Updated"
 * Criteria:
 *   - Source column F date >= 2024 OR
 *   - Source column E date >= 2025
 *
 * Mapping:  A -> Dest A,  F -> Dest B,  E -> Dest C
 * - Appends after the last non-empty row in A..C (ignores borders/formatting)
 * - First-ever write begins at row 2
 * - Skips duplicates based on A|F|E(Date)
 * - Formats only the newly written date cells (MM/dd/yyyy)
 */
function syncQuickCheck_Conditional() {
  // === CONFIG ===
  var SOURCE_URL = 'https://docs.google.com/spreadsheets/d/YOUR_SOURCE_SPREADSHEET_ID/edit'; // Source spreadsheet URL
  var SOURCE_TAB = 'Quick Check';         // Source sheet/tab name
  var DEST_TAB   = 'Quick Check Updated'; // Destination sheet/tab name

  var HEADER_ROW = 1;                  // Row index for header row
  var FIRST_DATA_ROW = HEADER_ROW + 1; // First row of actual data (row 2)

  // === Open sheets ===
  var destSS = SpreadsheetApp.getActive();                        // Active spreadsheet
  var destSh = destSS.getSheetByName(DEST_TAB) || destSS.insertSheet(DEST_TAB); // Destination sheet (create if missing)

  var srcSS = SpreadsheetApp.openByUrl(SOURCE_URL);               // Source spreadsheet by URL
  var srcSh = srcSS.getSheetByName(SOURCE_TAB);                   // Source sheet
  if (!srcSh) throw new Error('Source tab "' + SOURCE_TAB + '" not found.');

  var srcValues = srcSh.getDataRange().getValues();               // All values from source
  if (srcValues.length < 2) return;                               // Exit if only header/no data

  // === Helpers ===

  // Convert a value into a Date object if possible (handles multiple formats)
  function asDate(val) {
    if (val instanceof Date && !isNaN(val)) return val;
    if (val == null) return null;
    var s = String(val).trim();
    if (!s) return null;

    // Match mm/dd/yyyy or mm-dd-yyyy
    var m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
    if (m) {
      var d = parseInt(m[1], 10), mo = parseInt(m[2], 10) - 1, y = parseInt(m[3], 10);
      return new Date(y, mo, d);
    }

    // Match yyyy/mm/dd or yyyy-mm-dd
    m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
    if (m) {
      var y2 = parseInt(m[1], 10), mo2 = parseInt(m[2], 10) - 1, d2 = parseInt(m[3], 10);
      return new Date(y2, mo2, d2);
    }

    // Match year only (yyyy)
    var yOnly = s.match(/\b(\d{4})\b/);
    if (yOnly) return new Date(parseInt(yOnly[1], 10), 0, 1);

    return null;
  }

  // Format a Date object as yyyy-mm-dd string
  function yyyymmdd(dateObj) {
    var y = dateObj.getFullYear();
    var m = ('0' + (dateObj.getMonth() + 1)).slice(-2);
    var d = ('0' + dateObj.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }

  // Find last row in columns A–C with content (ignores formatting and blank formulas)
  function getLastContentRowInABC_(sheet) {
    var maxRows = sheet.getMaxRows();
    if (maxRows < FIRST_DATA_ROW) return HEADER_ROW;
    var rng = sheet.getRange(FIRST_DATA_ROW, 1, maxRows - FIRST_DATA_ROW + 1, 3);
    var values = rng.getValues();
    for (var i = values.length - 1; i >= 0; i--) {
      var row = values[i];
      if (
        (row[0] !== '' && row[0] != null) ||
        (row[1] !== '' && row[1] != null) ||
        (row[2] !== '' && row[2] != null)
      ) {
        return FIRST_DATA_ROW + i; // Return index of last row with data
      }
    }
    return HEADER_ROW; // No content beyond header
  }

  // === Build set of existing keys to skip duplicates ===
  var lastContentRow = getLastContentRowInABC_(destSh);
  var existingKeys = new Set();
  if (lastContentRow >= FIRST_DATA_ROW) {
    var existRange = destSh.getRange(FIRST_DATA_ROW, 1, lastContentRow - FIRST_DATA_ROW + 1, 3).getValues();
    for (var i = 0; i < existRange.length; i++) {
      var a = existRange[i][0]; // Column A
      var f = existRange[i][1]; // Column B (mapped from F in source)
      var e = existRange[i][2]; // Column C (mapped from E in source)
      var eDate = asDate(e);
      var eNorm = eDate ? yyyymmdd(eDate) : String(e).trim();
      var key = String(a).trim() + '|' + String(f).trim() + '|' + eNorm;
      existingKeys.add(key);
    }
  }

  // === Collect new rows from source ===
  var out = [];
  for (var r = 1; r < srcValues.length; r++) { // Start at row 2 (skip header)
    var row = srcValues[r];
    var colA = row[0]; // Source column A
    var colE = row[4]; // Source column E
    var colF = row[5]; // Source column F

    var dE = asDate(colE);
    var dF = asDate(colF);

    // Conditions: include if E ≥ 2025 or F ≥ 2024
    var condE = dE && dE.getFullYear() >= 2025;
    var condF = dF && dF.getFullYear() >= 2024;
    if (!(condE || condF)) continue;

    // Build unique key (A|F|E)
    var keyNew = String(colA).trim() + '|' + String(colF).trim() + '|' + (dE ? yyyymmdd(dE) : String(colE).trim());
    if (existingKeys.has(keyNew)) continue; // Skip if already exists

    // Prepare new row for output: A → col A, F → col B, E → col C
    out.push([colA, colF, dE || colE]);
    existingKeys.add(keyNew);
  }

  if (!out.length) return; // Nothing new to write

  // === Append to destination ===
  var writeRow = Math.max(FIRST_DATA_ROW, lastContentRow + 1);
  destSh.getRange(writeRow, 1, out.length, 3).setValues(out);          // Write values into A–C
  destSh.getRange(writeRow, 3, out.length, 1).setNumberFormat('MM/dd/yyyy'); // Format column C as date
}


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/**
 * Compare column A of 'Quick Check' and 'Quick Check Updated'.
 * If there's a match, copy columns D–G from 'Quick Check' to 'Quick Check Updated'.
 */

function updateQuickCheckUpdatedWithExtraColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var srcSh = ss.getSheetByName('Quick Check');          // Source sheet
  var destSh = ss.getSheetByName('Quick Check Updated'); // Destination sheet
  if (!srcSh || !destSh) {
    throw new Error("One or both tabs 'Quick Check' or 'Quick Check Updated' not found.");
  }

  // Get all values from both sheets
  var srcValues = srcSh.getDataRange().getValues();   // Source: entire sheet
  var destValues = destSh.getDataRange().getValues(); // Destination: entire sheet

  // Exit if either sheet has only headers or is empty
  if (srcValues.length < 2 || destValues.length < 2) return;

  // Build a lookup map from source column A -> columns D-G
  var srcMap = {};
  for (var i = 1; i < srcValues.length; i++) { // Start at row 2, skip header
    var key = srcValues[i][0]; // Column A of source
    if (key !== "" && key != null) {
      // Save columns D-G (slice(3,7) gives indices 3,4,5,6)
      srcMap[String(key).trim()] = srcValues[i].slice(3, 7);
    }
  }

  // Arrays to track updates and row numbers in destination
  var updates = [];
  var updateRows = [];

  // Iterate through destination rows
  for (var j = 1; j < destValues.length; j++) { // Skip header row
    var keyDest = destValues[j][0]; // Column A of destination
    if (keyDest !== "" && keyDest != null) {
      var srcData = srcMap[String(keyDest).trim()];
      if (srcData) {
        // Collect data to update in D-G
        updates.push(srcData);
        updateRows.push(j + 1); // Actual row number in sheet (+1 for header offset)
      }
    }
  }

  // Write updates back into destination sheet
  for (var k = 0; k < updates.length; k++) {
    var rowNum = updateRows[k];
    destSh.getRange(rowNum, 4, 1, 4).setValues([updates[k]]); // Write to cols D-G
  }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/**
 * Merges Starter and Business Plus source tabs into "Annual Review Updated",
 * combining a fill pass for existing rows with two append passes for new
 * ones. Unlike syncAnnualReviewFromBusinessPlusUpdated (which only reads an
 * already-processed "Business Plus Updated" tab), this reads the raw
 * Starter and Business Plus source tabs directly.
 *
 * Source file: an external spreadsheet holding both "Starter" and
 * "Business Plus" tabs (see SOURCE_URL below — set this to your own sheet).
 *   - Starter tab: data starts at row 7, columns:
 *       A (name), O (date, mm/dd/yyyy)
 *   - Business Plus tab: data starts at row 8, columns:
 *       A (name), C (date or "-"), E (date or "-"), I (string: "Yes" | "No" | "-" | "")
 *
 * Destination tab: "Annual Review Updated", columns:
 *   A = Name
 *   B = Date from Starter!O or Business Plus!C
 *   C = Date from Business Plus!E
 *   D = String from Business Plus!I (kept as text, not a date)
 *
 * Rules, in order:
 *   0. Fill pass — for each existing destination row, look up its name in
 *      the Business Plus tab:
 *        - Only proceeds if the name appears EXACTLY ONCE there (an
 *          ambiguous or missing match is left untouched).
 *        - If dest C is empty and Business Plus!E is a valid date, copy it in.
 *        - If dest D is empty and Business Plus!I is non-empty, copy it in.
 *   1. Append pass (Starter) — compare dest (A+B) against Starter
 *      (A+O); append any Starter rows not already present. If the same
 *      name+date also has an exact match in Business Plus (A+C), pull
 *      Business Plus!E and Business Plus!I into the new row's C/D.
 *   2. Append pass (Business Plus) — compare dest (A) against Business Plus (A);
 *      append any Business Plus names not already in the destination
 *      (B/C/D taken directly from Business Plus's C/E/I columns).
 *
 * Intended to run once per day via a time-driven trigger.
 */
function syncAnnualReviewFromStarterAndBusinessPlus() {
  const SOURCE_URL = 'https://docs.google.com/spreadsheets/d/YOUR_SOURCE_SPREADSHEET_ID/edit'; // Source spreadsheet URL
  const STARTER_TAB = 'Starter';
  const BUSINESS_PLUS_TAB = 'Business Plus';
  const DEST_TAB = 'Annual Review Updated';

  // === Helpers ===

  /** True when val is a valid Sheets Date object */
  function isValidDate_(val) {
    return val instanceof Date && !isNaN(val);
  }

  /** Normalize date to YYYY-MM-DD for stable keying */
  function normDate(val) {
    return isValidDate_(val)
      ? `${val.getFullYear()}-${('0' + (val.getMonth() + 1)).slice(-2)}-${('0' + val.getDate()).slice(-2)}`
      : '';
  }

  /** Normalize names for matching (trim + lowercase) */
  function normName(val) {
    return String(val || '').trim().toLowerCase();
  }

  /** Composite key: name|date */
  function makeKey(name, dateVal) {
    return normName(name) + '|' + normDate(dateVal);
  }

  /** Treat empty as true only for genuinely blank cells */
  function isEmptyCell_(val) {
    return val === '' || val === null || typeof val === 'undefined';
  }

  // Find the last real row in column A (ignores formatting + ghost rows)
  function getLastNonEmptyRowA_(sheet) {
    const colA = sheet.getRange("A:A").getValues();
    for (let i = colA.length - 1; i >= 0; i--) {
      if (colA[i][0] !== "" && colA[i][0] !== null) {
        return i + 1; // Convert index to row number
      }
    }
    return 1; // Only header exists
  }

  // === Open sheets ===
  const destSS = SpreadsheetApp.getActive();
  const destSh = destSS.getSheetByName(DEST_TAB) || destSS.insertSheet(DEST_TAB);

  const srcSS = SpreadsheetApp.openByUrl(SOURCE_URL);
  const starterSh = srcSS.getSheetByName(STARTER_TAB);
  const businessPlusSh = srcSS.getSheetByName(BUSINESS_PLUS_TAB);

  // === Read existing destination ===
  const destVals = destSh.getDataRange().getValues();
  const destKeys = new Set();   // set of "name|date(B)" already present
  const destNames = new Set();  // set of normalized names already present

  for (let r = 1; r < destVals.length; r++) {
    const a = destVals[r][0]; // A (name)
    const b = destVals[r][1]; // B (date)
    if (a) destNames.add(normName(a));
    if (a && isValidDate_(b)) {
      destKeys.add(makeKey(a, b));
    }
  }

  // === Read sources ===
  // Starter rows: from row 7, columns A..O (15 columns)
  let starterVals = [];
  const starterCount = starterSh.getLastRow() - 6;
  if (starterCount > 0) {
    starterVals = starterSh.getRange(7, 1, starterCount, 15).getValues(); // A..O
  }

  // Business Plus rows: from row 8, columns A..I (9 columns)
  let businessPlusVals = [];
  const businessPlusCount = businessPlusSh.getLastRow() - 7;
  if (businessPlusCount > 0) {
    businessPlusVals = businessPlusSh.getRange(8, 1, businessPlusCount, 9).getValues(); // A..I
  }

  // --- Build Business Plus indices ---

  // 1) Exact (A+C) map for the append-pass join (only when C is a valid date)
  const businessPlusByKey = {};

  // 2) Name → {count, E, I} for the fill pass (no selection when multiple matches)
  const businessPlusNameInfo = {};

  for (let r = 0; r < businessPlusVals.length; r++) {
    const name = businessPlusVals[r][0];   // A
    const dateC = businessPlusVals[r][2];  // C (date)
    const dateE = businessPlusVals[r][4];  // E (date)
    const strI  = businessPlusVals[r][8];  // I (string: Yes/No/-/"")
    if (!name) continue;

    // (A+C) → for exact key joins
    if (isValidDate_(dateC)) {
      businessPlusByKey[makeKey(name, dateC)] = {
        E: isValidDate_(dateE) ? new Date(dateE) : null,
        I: (strI !== null && strI !== undefined) ? String(strI) : ''
      };
    }

    // Name counts (no "most recent" picking; we only use values when count === 1)
    const n = normName(name);
    if (!businessPlusNameInfo[n]) {
      businessPlusNameInfo[n] = {
        count: 1,
        E: isValidDate_(dateE) ? new Date(dateE) : null,
        I: (strI !== null && strI !== undefined) ? String(strI) : ''
      };
    } else {
      businessPlusNameInfo[n].count += 1;
      // Keep first-seen E/I but they will be ignored if count > 1.
    }
  }

  // === Rule 0: Fill missing C and D in existing destination rows by matching names to Business Plus ===
  let needCDUpdate = false;
  if (destVals.length > 1) {
    for (let r = 1; r < destVals.length; r++) {
      const name = destVals[r][0]; // A
      if (!name) continue;
      const rn = normName(name);
      const info = businessPlusNameInfo[rn];
      if (!info || info.count !== 1) continue; // ambiguous or not found → do nothing

      // C: from Business Plus!E if dest C is empty and Business Plus!E is a valid date
      if (isEmptyCell_(destVals[r][2]) && isValidDate_(info.E)) {
        destVals[r][2] = new Date(info.E);
        needCDUpdate = true;
      }

      // D: from Business Plus!I (string) if dest D is empty and Business Plus!I is non-empty
      if (isEmptyCell_(destVals[r][3]) && String(info.I).trim() !== '') {
        destVals[r][3] = String(info.I);
        needCDUpdate = true;
      }
    }
  }

  // === Collect rows to append ===
  const out = [];

  // Rule 1: Append Starter rows that aren't already in dest (match on A+B vs A+O)
  for (let r = 0; r < starterVals.length; r++) {
    const name = starterVals[r][0];       // A
    const dateO = starterVals[r][14];     // O
    if (!name || !isValidDate_(dateO)) continue;

    const key = makeKey(name, dateO);
    if (!destKeys.has(key)) {
      // newRow: [A=name, B=date from Starter!O, C=from Business Plus!E (if exact match), D=from Business Plus!I (string, if exact match)]
      const newRow = [name, new Date(dateO), '', ''];

      if (businessPlusByKey[key]) {
        const bp = businessPlusByKey[key];
        if (isValidDate_(bp.E)) newRow[2] = new Date(bp.E);
        else if (bp.E) newRow[2] = '-';

        // D is a string; copy as-is (could be "Yes", "No", "-", or "")
        if (bp.I !== undefined && bp.I !== null) newRow[3] = String(bp.I);
      }

      out.push(newRow);
      destKeys.add(key);
      destNames.add(normName(name));
    }
  }

  // Rule 2: Append Business Plus names not already in dest
  for (let r = 0; r < businessPlusVals.length; r++) {
    const name = businessPlusVals[r][0];     // A
    if (!name) continue;
    if (destNames.has(normName(name))) continue; // skip if name already in dest

    const dateC = businessPlusVals[r][2];    // C (date)
    const dateE = businessPlusVals[r][4];    // E (date)
    const strI  = businessPlusVals[r][8];    // I (string)

    const newRow = [
      name,
      isValidDate_(dateC) ? new Date(dateC) : '-', // B
      isValidDate_(dateE) ? new Date(dateE) : '-', // C
      (strI !== null && strI !== undefined) ? String(strI) : '' // D (string)
    ];

    out.push(newRow);
    destNames.add(normName(name));
  }

  // === Write updates ===

  // 0) If we filled any existing C/D cells, write those back in one batch
  if (needCDUpdate && destVals.length > 1) {
    const cdValues = destVals.slice(1).map(row => [row[2], row[3]]); // C & D only, excluding header
    destSh.getRange(2, 3, cdValues.length, 2).setValues(cdValues);
  }

  // 1) Append any new rows collected from Starter/Business Plus
  if (out.length) {
    const appendStartRow = getLastNonEmptyRowA_(destSh) + 1;
    destSh.getRange(appendStartRow, 1, out.length, out[0].length).setValues(out);
  }

  // 2) Date formatting for B and C ONLY (D is a string and should remain unformatted)
  const lastRow = destSh.getLastRow();
  if (lastRow >= 2) {
    destSh.getRange(2, 2, lastRow - 1, 2).setNumberFormat('MM/dd/yyyy'); // Columns B..C from row 2
  }
}
