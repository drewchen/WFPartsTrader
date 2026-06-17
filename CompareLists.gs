/**
 * Compares "List A" and "List B" inventory tables on the Inventory sheet,
 * using a normalized Prime-item lookup table on the "Prime Data" sheet,
 * and writes four result tables to the Trades sheet:
 *   1. Extras from A wanted by B
 *   2. Extras from B wanted by A
 *   3. Extras from A to sell (grouped by Ducats, highest first)
 *   4. Extras from B to sell (grouped by Ducats, highest first)
 *
 * --- Data model ---
 *
 * "Prime Data" sheet (the immutable lookup table):
 *   Columns: Set | Part | Ducats | Qty Required
 *   Row 1 = headers, data from row 2.
 *   One row per part of a Prime set, e.g.:
 *     Acceltra | Barrel     | 100 | 1
 *     Acceltra | Receiver   | 45  | 1
 *     Acceltra | Stock      | 65  | 1
 *     Acceltra | Blueprint  | 15  | 1
 *
 * "Inventory" sheet (List A and List B, side by side):
 *   List A: columns A:C -> Set (Wanted) | Part | Owned
 *   List B: columns E:G -> Set (Wanted) | Part | Owned
 *   Row 1 = title/name, Row 2 = headers, data from row 3.
 *
 *   "Set (Wanted)" cells look like "Acceltra (1)" meaning the owner wants
 *   1 full Acceltra set. This is read once per set (on its first row) and
 *   applies to every part row underneath it until the next named Set cell,
 *   matching the visual "merged cell" layout of the sheet.
 *
 *   A part that the owner has zero of is simply omitted from their list
 *   (e.g. no "Stock" row for Acceltra) -- the script infers Owned = 0 for
 *   any part required by a wanted set that isn't explicitly listed.
 *
 * Run via the "Inventory Tools > Compare Lists" custom menu, or directly
 * from the Apps Script editor by running compareLists().
 */

var INVENTORY_SHEET = 'Inventory';
var TRADES_SHEET = 'Trades';
var PRIME_DATA_SHEET = 'Prime Data';

var LIST_A_RANGE = 'A3:C'; // Set (Wanted), Part, Owned
var LIST_B_RANGE = 'E3:G';

var PRIME_DATA_RANGE = 'A2:D'; // Set, Part, Ducats, Qty Required

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory Tools')
    .addItem('Compare Lists', 'compareLists')
    .addToUi();
}

/**
 * Reads the Prime Data lookup table into a structure keyed by Set name:
 *   { "Acceltra": [ {part, ducats, qtyRequired}, ... ], ... }
 */
function readPrimeData_(sheet) {
  var values = sheet.getRange(PRIME_DATA_RANGE).getValues();
  var bySet = {};

  values.forEach(function (row) {
    var set = String(row[0]).trim();
    var part = String(row[1]).trim();
    if (set === '' || part === '') return;

    var ducats = (row[2] === '' || row[2] === null) ? '' : row[2];
    var qtyRequired = (row[3] === '' || row[3] === null) ? 1 : Number(row[3]);

    if (!bySet[set]) bySet[set] = [];
    bySet[set].push({ part: part, ducats: ducats, qtyRequired: qtyRequired });
  });

  return bySet;
}

/**
 * Parses a "Set (Wanted)" cell like "Acceltra (1)" into { set, wantedSets }.
 * Returns null if the cell is blank.
 */
function parseSetWantedCell_(raw) {
  var text = String(raw).trim();
  if (text === '') return null;

  var match = text.match(/^(.*?)\s*\((\d+)\)\s*$/);
  if (match) {
    return { set: match[1].trim(), wantedSets: Number(match[2]) };
  }
  // Fallback: no "(n)" found -- treat as a set name with 0 wanted.
  return { set: text, wantedSets: 0 };
}

/**
 * Reads one list's data block (Set (Wanted) | Part | Owned) and expands it,
 * using the Prime Data lookup table, into a full per-part row list:
 *   { set, part, ducats, wanted, owned, extra, needed }
 *
 * Any part required by a wanted set but not explicitly listed in the
 * owner's rows is inferred with Owned = 0.
 */
function readList_(sheet, rangeA1, primeData) {
  var values = sheet.getRange(rangeA1).getValues();

  // First pass: figure out which Set each row belongs to (carrying the
  // "Set (Wanted)" cell down through blank-Set rows), and how many of
  // each set are wanted, and what's explicitly Owned per Set+Part.
  var currentSet = null;
  var wantedBySet = {};   // set -> number of sets wanted
  var ownedByKey = {};    // "set||part" -> owned count

  values.forEach(function (row) {
    var setCell = row[0];
    var part = String(row[1]).trim();
    var owned = row[2];

    var parsed = parseSetWantedCell_(setCell);
    if (parsed) {
      currentSet = parsed.set;
      wantedBySet[currentSet] = parsed.wantedSets;
    }

    if (part === '' || currentSet === null) return; // spacer row

    owned = (owned === '' || owned === null) ? 0 : Number(owned);
    ownedByKey[currentSet + '||' + part] = owned;
  });

  // Second pass: for every set that appears (wanted or just owned), expand
  // via the lookup table into one row per part.
  var rows = [];
  var allSets = Object.keys(wantedBySet);

  allSets.forEach(function (setName) {
    var wantedSets = wantedBySet[setName] || 0;
    var partsForSet = primeData[setName];

    if (!partsForSet) {
      throw new Error('Set "' + setName + '" not found in Prime Data lookup table. ' +
        'Check spelling on the Inventory sheet.');
    }

    partsForSet.forEach(function (p) {
      var key = setName + '||' + p.part;
      var owned = (key in ownedByKey) ? ownedByKey[key] : 0;
      var wantedQty = wantedSets * p.qtyRequired;

      rows.push({
        set: setName,
        part: p.part,
        ducats: p.ducats,
        wanted: wantedQty,
        owned: owned,
        extra: Math.max(owned - wantedQty, 0),
        needed: Math.max(wantedQty - owned, 0)
      });
    });
  });

  return rows;
}

function indexList_(rows) {
  var index = {};
  rows.forEach(function (r) {
    index[r.set + '||' + r.part] = r;
  });
  return index;
}

function compareLists() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName(INVENTORY_SHEET);
  var tradesSheet = ss.getSheetByName(TRADES_SHEET);
  var primeSheet = ss.getSheetByName(PRIME_DATA_SHEET);

  if (!invSheet) throw new Error('Sheet "' + INVENTORY_SHEET + '" not found.');
  if (!tradesSheet) throw new Error('Sheet "' + TRADES_SHEET + '" not found.');
  if (!primeSheet) throw new Error('Sheet "' + PRIME_DATA_SHEET + '" not found.');

  var primeData = readPrimeData_(primeSheet);

  var listA = readList_(invSheet, LIST_A_RANGE, primeData);
  var listB = readList_(invSheet, LIST_B_RANGE, primeData);
  var indexA = indexList_(listA);
  var indexB = indexList_(listB);

  var aWantedByB = [];
  var bWantedByA = [];
  var aToSell = [];
  var bToSell = [];

  listA.forEach(function (rowA) {
    if (rowA.extra <= 0) return;
    var match = indexB[rowA.set + '||' + rowA.part];
    var remaining = rowA.extra;

    if (match && match.needed > 0) {
      var transferable = Math.min(remaining, match.needed);
      aWantedByB.push({ set: rowA.set, part: rowA.part, ducats: rowA.ducats, qty: transferable });
      remaining -= transferable;
    }
    if (remaining > 0) {
      aToSell.push({ set: rowA.set, part: rowA.part, ducats: rowA.ducats, qty: remaining });
    }
  });

  listB.forEach(function (rowB) {
    if (rowB.extra <= 0) return;
    var match = indexA[rowB.set + '||' + rowB.part];
    var remaining = rowB.extra;

    if (match && match.needed > 0) {
      var transferable = Math.min(remaining, match.needed);
      bWantedByA.push({ set: rowB.set, part: rowB.part, ducats: rowB.ducats, qty: transferable });
      remaining -= transferable;
    }
    if (remaining > 0) {
      bToSell.push({ set: rowB.set, part: rowB.part, ducats: rowB.ducats, qty: remaining });
    }
  });

  function sortByDucats(arr) {
    arr.sort(function (a, b) {
      var da = (a.ducats === '' || a.ducats === null) ? -Infinity : Number(a.ducats);
      var db = (b.ducats === '' || b.ducats === null) ? -Infinity : Number(b.ducats);
      if (db !== da) return db - da;
      if (a.set !== b.set) return a.set < b.set ? -1 : 1;
      return a.part < b.part ? -1 : (a.part > b.part ? 1 : 0);
    });
    return arr;
  }
  function sortBySetPart(arr) {
    arr.sort(function (a, b) {
      if (a.set !== b.set) return a.set < b.set ? -1 : 1;
      return a.part < b.part ? -1 : (a.part > b.part ? 1 : 0);
    });
    return arr;
  }

  sortByDucats(aToSell);
  sortByDucats(bToSell);
  sortBySetPart(aWantedByB);
  sortBySetPart(bWantedByA);

  writeOutput_(tradesSheet, aWantedByB, bWantedByA, aToSell, bToSell);
}

function writeOutput_(sheet, aWantedByB, bWantedByA, aToSell, bToSell) {
  sheet.clear();

  var row = 1;
  row = writeTable_(sheet, row, 'Extras from List A wanted by List B',
    ['Set', 'Part', 'Ducats', 'Qty'], aWantedByB);
  row += 2;
  row = writeTable_(sheet, row, 'Extras from List B wanted by List A',
    ['Set', 'Part', 'Ducats', 'Qty'], bWantedByA);
  row += 2;
  row = writeTable_(sheet, row, 'Extras from List A to sell (by Ducats)',
    ['Set', 'Part', 'Ducats', 'Qty'], aToSell);
  row += 2;
  writeTable_(sheet, row, 'Extras from List B to sell (by Ducats)',
    ['Set', 'Part', 'Ducats', 'Qty'], bToSell);
}

function writeTable_(sheet, startRow, title, headers, rows) {
  sheet.getRange(startRow, 1).setValue(title).setFontWeight('bold');
  sheet.getRange(startRow + 1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  var dataRow = startRow + 2;

  if (rows.length === 0) {
    sheet.getRange(dataRow, 1).setValue('(none)').setFontStyle('italic');
    return dataRow;
  }

  var values = rows.map(function (r) {
    return [r.set, r.part, r.ducats, r.qty];
  });
  sheet.getRange(dataRow, 1, values.length, headers.length).setValues(values);

  return dataRow + values.length - 1;
}