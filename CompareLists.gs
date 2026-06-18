/**
 * Compares "List A" and "List B" inventory tables on the Inventory sheet,
 * using a normalized Prime-item lookup table on the "Prime Data" sheet,
 * and writes six result tables to the Trades sheet, each with a "Resolve"
 * checkbox column that applies the trade/sale directly back to Inventory:
 *   1. Extras from A wanted by B       (Resolve: A.Owned -1, B.Owned +1)
 *   2. Extras from B wanted by A       (Resolve: B.Owned -1, A.Owned +1)
 *   3. Extras from A to sell (Ducats)  (Resolve: A.Owned -1)
 *   4. Extras from B to sell (Ducats)  (Resolve: B.Owned -1)
 *   5. List A parts to sell on market  (Resolve: A.Owned -1)
 *   6. List B parts to sell on market  (Resolve: B.Owned -1)
 *
 * --- Data model ---
 *
 * "Prime Data" sheet (the immutable lookup table):
 *   Columns: Set | Part | Ducats | Qty Required | Sell on Market
 *   Row 1 = headers, data from row 2.
 *   "Sell on Market" is any non-empty flag (e.g. "x") marking a part that,
 *   when left over, should be routed to its own "sell on market" output
 *   instead of the regular ducat-sorted "to sell" table. Flagged parts are
 *   still offered to the other list first if that owner needs them --
 *   the flag only changes where *leftover* extras land.
 *
 * "Inventory" sheet (List A and List B, side by side):
 *   List A: columns A:C -> Set (Wanted) | Part | Owned
 *   List B: columns E:G -> Set (Wanted) | Part | Owned
 *   Row 1 = title/name, Row 2 = headers, data from row 3.
 *
 *   "Set (Wanted)" cells look like "Acceltra (1)" meaning the owner wants
 *   1 full Acceltra set. If no "(n)" is present, wanting 1 set is assumed.
 *   A part the owner has zero of can be omitted entirely from their list;
 *   Owned = 0 is inferred for any part required by a wanted set that isn't
 *   explicitly listed.
 *
 * --- Resolve feature ---
 *
 * Each output row gets a checkbox in column E ("Resolve"). Columns F:G
 * hold a hidden action payload (not for manual editing): F = action type
 * ("transfer" or "sell"), G = a pipe-delimited descriptor of which Set+Part
 * to update on which list (e.g. "sell|A|Acceltra|Stock"). Ticking the
 * checkbox triggers onEditResolve_ (installed as an installable onEdit
 * trigger -- see setupTrigger()), which:
 *   1. Looks up the live Inventory row for the Set+Part fresh, by scanning
 *      the sheet at the moment of the click (never relies on a cached row
 *      number), so it stays correct even if earlier Resolve clicks in the
 *      same session inserted or shifted rows.
 *   2. Verifies the decrement won't go negative (warns and aborts if so,
 *      leaving the checkbox unchecked, in case Inventory was hand-edited
 *      since the last Compare Lists run).
 *   3. Decrements the source Owned cell by 1 (floor 0, never deleted --
 *      the row stays, just set to 0) and, for transfers, increments the
 *      destination Owned cell by 1. If the destination part has no
 *      existing row in Inventory, a new row is inserted directly under
 *      that Set's last existing row.
 *   4. Marks the Trades row "Finished": checkbox replaced with locked
 *      "Finished" text, row grayed out.
 *
 * The Resolve column (and its hidden helper columns) are regenerated
 * fresh every time Compare Lists runs, so previously-finished rows reset.
 *
 * Run via the "Inventory Tools > Compare Lists" custom menu, or directly
 * from the Apps Script editor by running compareLists(). Run setupTrigger()
 * once (from the script editor, not the menu) to enable the Resolve
 * buttons -- this requires one-time authorization.
 */

var INVENTORY_SHEET = 'Inventory';
var TRADES_SHEET = 'Trades';
var PRIME_DATA_SHEET = 'Prime Data';

var LIST_A_RANGE = 'A3:C'; // Set (Wanted), Part, Owned
var LIST_B_RANGE = 'E3:G';

// Column letters for List A / List B on Inventory (needed for direct writes
// back to specific Owned cells, and for inserting new rows).
var LIST_A_SET_COL = 1;   // A
var LIST_A_PART_COL = 2;  // B
var LIST_A_OWNED_COL = 3; // C
var LIST_B_SET_COL = 5;   // E
var LIST_B_PART_COL = 6;  // F
var LIST_B_OWNED_COL = 7; // G
var INVENTORY_DATA_START_ROW = 3;

var PRIME_DATA_RANGE = 'A2:E'; // Set, Part, Ducats, Qty Required, Sell on Market

// Trades sheet output column layout per table:
// A=Set, B=Part, C=Ducats, D=Qty, E=Resolve (checkbox), F=action type (hidden),
// G=action payload (hidden)
var RESOLVE_COL = 5;
var ACTION_TYPE_COL = 6;
var ACTION_PAYLOAD_COL = 7;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory Tools')
    .addItem('Compare Lists', 'compareLists')
    .addItem('Enable Resolve buttons (run once)', 'setupTrigger')
    .addToUi();
}

/**
 * One-time setup: installs an installable onEdit trigger so that ticking a
 * Resolve checkbox actually fires onEditResolve_ with full edit permissions
 * (simple onEdit cannot write to other sheets without this). Run this once
 * from the menu; re-running is harmless (it won't create duplicate triggers).
 */
function setupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var exists = triggers.some(function (t) {
    return t.getHandlerFunction() === 'onEditResolve_';
  });
  if (exists) {
    SpreadsheetApp.getUi().alert('Resolve buttons are already enabled.');
    return;
  }
  ScriptApp.newTrigger('onEditResolve_')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert('Resolve buttons are now enabled.');
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
    var sellOnMarket = (row[4] !== '' && row[4] !== null && String(row[4]).trim() !== '');

    if (!bySet[set]) bySet[set] = [];
    bySet[set].push({ part: part, ducats: ducats, qtyRequired: qtyRequired, sellOnMarket: sellOnMarket });
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
  // Fallback: no "(n)" found -- default to wanting 1 full set.
  return { set: text, wantedSets: 1 };
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
        sellOnMarket: p.sellOnMarket,
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
  var aToSellOnMarket = [];
  var bToSellOnMarket = [];

  listA.forEach(function (rowA) {
    if (rowA.extra <= 0) return;
    var match = indexB[rowA.set + '||' + rowA.part];
    var remaining = rowA.extra;

    if (match && match.needed > 0) {
      var transferable = Math.min(remaining, match.needed);
      aWantedByB.push({
        set: rowA.set, part: rowA.part, ducats: rowA.ducats, qty: transferable,
        action: 'transfer',
        from: { list: 'A', set: rowA.set, part: rowA.part },
        to: { list: 'B', set: rowA.set, part: rowA.part }
      });
      remaining -= transferable;
    }
    if (remaining > 0) {
      var bucket = rowA.sellOnMarket ? aToSellOnMarket : aToSell;
      bucket.push({
        set: rowA.set, part: rowA.part, ducats: rowA.ducats, qty: remaining,
        action: 'sell',
        from: { list: 'A', set: rowA.set, part: rowA.part }
      });
    }
  });

  listB.forEach(function (rowB) {
    if (rowB.extra <= 0) return;
    var match = indexA[rowB.set + '||' + rowB.part];
    var remaining = rowB.extra;

    if (match && match.needed > 0) {
      var transferable = Math.min(remaining, match.needed);
      bWantedByA.push({
        set: rowB.set, part: rowB.part, ducats: rowB.ducats, qty: transferable,
        action: 'transfer',
        from: { list: 'B', set: rowB.set, part: rowB.part },
        to: { list: 'A', set: rowB.set, part: rowB.part }
      });
      remaining -= transferable;
    }
    if (remaining > 0) {
      var bucketB = rowB.sellOnMarket ? bToSellOnMarket : bToSell;
      bucketB.push({
        set: rowB.set, part: rowB.part, ducats: rowB.ducats, qty: remaining,
        action: 'sell',
        from: { list: 'B', set: rowB.set, part: rowB.part }
      });
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
  sortBySetPart(aToSellOnMarket);
  sortBySetPart(bToSellOnMarket);

  writeOutput_(tradesSheet, aWantedByB, bWantedByA, aToSell, bToSell, aToSellOnMarket, bToSellOnMarket);
}

function writeOutput_(sheet, aWantedByB, bWantedByA, aToSell, bToSell, aToSellOnMarket, bToSellOnMarket) {
  sheet.clear();
  sheet.showColumns(1, 7); // in case columns were hidden from a prior run

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
  row = writeTable_(sheet, row, 'Extras from List B to sell (by Ducats)',
    ['Set', 'Part', 'Ducats', 'Qty'], bToSell);
  row += 2;
  row = writeTable_(sheet, row, 'List A parts to sell on market',
    ['Set', 'Part', 'Ducats', 'Qty'], aToSellOnMarket);
  row += 2;
  writeTable_(sheet, row, 'List B parts to sell on market',
    ['Set', 'Part', 'Ducats', 'Qty'], bToSellOnMarket);

  // Hide the helper columns (action type + payload) -- not for manual editing.
  sheet.hideColumns(ACTION_TYPE_COL, 2);
}

/**
 * Installable onEdit handler (see setupTrigger). Fires on every edit to
 * the spreadsheet; only acts when the edit is a checkbox being checked
 * in the Resolve column (E) of the Trades sheet, on a row that hasn't
 * already been finished.
 */
function onEditResolve_(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    if (sheet.getName() !== TRADES_SHEET) return;
    if (range.getColumn() !== RESOLVE_COL || range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;
    if (e.value !== 'TRUE') return; // only act on check, not uncheck

    var row = range.getRow();
    var actionType = sheet.getRange(row, ACTION_TYPE_COL).getValue();
    var payload = sheet.getRange(row, ACTION_PAYLOAD_COL).getValue();
    if (!actionType || !payload) return; // header/title/blank row, not a data row

    var ok = applyResolveAction_(actionType, payload);

    if (ok) {
      markRowFinished_(sheet, row);
    } else {
      // Revert the checkbox so the row stays actionable.
      range.setValue(false);
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert('Resolve failed: ' + err.message);
    try { e.range.setValue(false); } catch (e2) { /* ignore */ }
  }
}

/**
 * Parses the pipe-delimited action payload and applies it to Inventory.
 * Returns true on success, false if validation failed (e.g. would go
 * negative because Inventory was hand-edited since the last compare).
 */
function applyResolveAction_(actionType, payload) {
  var invSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVENTORY_SHEET);
  if (!invSheet) throw new Error('Inventory sheet not found.');

  var parts = payload.split('|');

  function parseSide(arr) {
    return { list: arr[0], set: arr[1], part: arr[2] };
  }

  if (actionType === 'sell') {
    var from = parseSide(parts.slice(1, 4));
    return decrementOwned_(invSheet, from);
  }

  if (actionType === 'transfer') {
    // payload: transfer|<from 3 fields>|<to 3 fields>
    var fromParts = parts.slice(1, 4);
    var toParts = parts.slice(4, 7);
    var from = parseSide(fromParts);
    var to = parseSide(toParts);

    var fromRow = findInventoryRow_(invSheet, from);
    if (!fromRow || Number(invSheet.getRange(fromRow, listColumns_(from.list).owned).getValue() || 0) < 1) {
      SpreadsheetApp.getUi().alert(
        'Could not resolve: ' + from.set + ' ' + from.part + ' on List ' + from.list +
        ' is already at 0 Owned (Inventory may have changed since Compare Lists last ran).');
      return false;
    }

    decrementOwned_(invSheet, from);
    incrementOwned_(invSheet, to);
    return true;
  }

  throw new Error('Unknown action type: ' + actionType);
}

function listColumns_(listLetter) {
  if (listLetter === 'A') {
    return { set: LIST_A_SET_COL, part: LIST_A_PART_COL, owned: LIST_A_OWNED_COL };
  }
  return { set: LIST_B_SET_COL, part: LIST_B_PART_COL, owned: LIST_B_OWNED_COL };
}

/**
 * Live scan of the Inventory sheet's List A or List B block to find the
 * sheet row number currently holding this Set+Part, carrying the most
 * recent non-blank Set cell down through blank-Set continuation rows
 * (same convention as readList_). Returns null if no row exists yet for
 * this Set+Part. Always reads fresh from the sheet -- never cached --
 * so it stays correct even after previous Resolve actions inserted rows.
 */
function findInventoryRow_(invSheet, side) {
  var cols = listColumns_(side.list);
  var lastRow = invSheet.getLastRow();
  if (lastRow < INVENTORY_DATA_START_ROW) return null;

  var numRows = lastRow - INVENTORY_DATA_START_ROW + 1;
  var setVals = invSheet.getRange(INVENTORY_DATA_START_ROW, cols.set, numRows, 1).getValues();
  var partVals = invSheet.getRange(INVENTORY_DATA_START_ROW, cols.part, numRows, 1).getValues();

  var currentSet = null;
  for (var i = 0; i < numRows; i++) {
    var setCell = String(setVals[i][0]).trim();
    var part = String(partVals[i][0]).trim();

    if (setCell !== '') {
      var parsed = parseSetWantedCell_(setCell);
      currentSet = parsed ? parsed.set : null;
    }
    if (part === '' || currentSet === null) continue;

    if (currentSet === side.set && part === side.part) {
      return INVENTORY_DATA_START_ROW + i;
    }
  }
  return null;
}

/**
 * Finds the last sheet row belonging to the given Set (for inserting a
 * new part row directly underneath it), via the same live scan.
 */
function findLastRowForSet_(invSheet, listLetter, setName) {
  var cols = listColumns_(listLetter);
  var lastRow = invSheet.getLastRow();
  if (lastRow < INVENTORY_DATA_START_ROW) return null;

  var numRows = lastRow - INVENTORY_DATA_START_ROW + 1;
  var setVals = invSheet.getRange(INVENTORY_DATA_START_ROW, cols.set, numRows, 1).getValues();
  var partVals = invSheet.getRange(INVENTORY_DATA_START_ROW, cols.part, numRows, 1).getValues();

  var currentSet = null;
  var lastMatchingRow = null;
  for (var i = 0; i < numRows; i++) {
    var setCell = String(setVals[i][0]).trim();
    var part = String(partVals[i][0]).trim();

    if (setCell !== '') {
      var parsed = parseSetWantedCell_(setCell);
      currentSet = parsed ? parsed.set : null;
    }
    if (part === '' || currentSet === null) continue;

    if (currentSet === setName) {
      lastMatchingRow = INVENTORY_DATA_START_ROW + i;
    }
  }
  return lastMatchingRow;
}

/**
 * Decrements the Owned cell for this side by 1, floored at 0. The row is
 * never deleted -- if Owned would go below 0, this returns false instead.
 */
function decrementOwned_(invSheet, side) {
  var rowNum = findInventoryRow_(invSheet, side);
  if (!rowNum) {
    SpreadsheetApp.getUi().alert(
      'Could not resolve: ' + side.set + ' ' + side.part + ' on List ' + side.list +
      ' has no Owned value to decrement (Inventory may have changed since Compare Lists last ran).');
    return false;
  }
  var cols = listColumns_(side.list);
  var cell = invSheet.getRange(rowNum, cols.owned);
  var current = Number(cell.getValue() || 0);
  if (current < 1) {
    SpreadsheetApp.getUi().alert(
      'Could not resolve: ' + side.set + ' ' + side.part + ' on List ' + side.list +
      ' is already at 0 Owned (Inventory may have changed since Compare Lists last ran).');
    return false;
  }
  cell.setValue(current - 1);
  return true;
}

/**
 * Increments the Owned cell for this side by 1. If no row currently exists
 * for this Set+Part (the part was previously inferred at Owned=0 with no
 * row on the sheet), a new row is inserted directly under the last known
 * row of that Set, with the Set cell left blank (since it's a continuation
 * of the existing Set block) and Owned set to 1.
 *
 * IMPORTANT: this uses Range.insertCells(Dimension.ROWS) scoped to just
 * this list's 3 columns (Set/Part/Owned), NOT Sheet.insertRowAfter().
 * insertRowAfter would insert a full-width row across the whole sheet,
 * shifting the *other* list's columns (which live on the same physical
 * rows but are a separate, unrelated list) down by one and corrupting
 * their alignment. insertCells confines the shift to this list's own
 * column range only, leaving the other list's rows completely untouched.
 *
 * Row lookup is always live, so this is safe to call after other Resolve
 * actions have already shifted rows earlier in the same session.
 */
function incrementOwned_(invSheet, side) {
  var cols = listColumns_(side.list);
  var rowNum = findInventoryRow_(invSheet, side);

  if (rowNum) {
    var cell = invSheet.getRange(rowNum, cols.owned);
    var current = Number(cell.getValue() || 0);
    cell.setValue(current + 1);
    return;
  }

  // No existing row for this Set+Part -- insert one under the set's last row,
  // scoped to only this list's 3 columns (Set, Part, Owned).
  var lastRowForSet = findLastRowForSet_(invSheet, side.list, side.set);
  var insertAt = (lastRowForSet || INVENTORY_DATA_START_ROW - 1) + 1;
  var firstCol = Math.min(cols.set, cols.part, cols.owned);
  var lastCol = Math.max(cols.set, cols.part, cols.owned);
  var numCols = lastCol - firstCol + 1;

  invSheet.getRange(insertAt, firstCol, 1, numCols).insertCells(SpreadsheetApp.Dimension.ROWS);
  invSheet.getRange(insertAt, cols.part).setValue(side.part);
  invSheet.getRange(insertAt, cols.owned).setValue(1);
  // Set (Wanted) cell intentionally left blank -- this row is a continuation
  // of the existing Set block above it, matching the sheet's layout convention.
}

/**
 * Marks a Trades row as finished: replaces the checkbox with a locked,
 * grayed-out "Finished" label.
 */
function markRowFinished_(sheet, row) {
  var resolveCell = sheet.getRange(row, RESOLVE_COL);
  resolveCell.removeCheckboxes();
  resolveCell.setValue('Finished');

  var fullRowRange = sheet.getRange(row, 1, 1, RESOLVE_COL);
  fullRowRange.setFontColor('#999999').setFontStyle('italic');
  resolveCell.setFontWeight('bold');
}

function writeTable_(sheet, startRow, title, headers, rows) {
  var fullHeaders = headers.concat(['Resolve']);
  sheet.getRange(startRow, 1).setValue(title).setFontWeight('bold');
  sheet.getRange(startRow + 1, 1, 1, fullHeaders.length).setValues([fullHeaders]).setFontWeight('bold');

  var dataRow = startRow + 2;

  if (rows.length === 0) {
    sheet.getRange(dataRow, 1).setValue('(none)').setFontStyle('italic');
    return dataRow;
  }

  var values = rows.map(function (r) {
    return [r.set, r.part, r.ducats, r.qty, false]; // false = unchecked checkbox
  });
  sheet.getRange(dataRow, 1, values.length, headers.length + 1).setValues(values);

  // Insert checkboxes into the Resolve column.
  var resolveRange = sheet.getRange(dataRow, RESOLVE_COL, values.length, 1);
  resolveRange.insertCheckboxes();

  // Write hidden action-type + payload columns (F, G) used by onEditResolve_.
  var actionRows = rows.map(function (r) {
    return [r.action, encodeAction_(r)];
  });
  sheet.getRange(dataRow, ACTION_TYPE_COL, actionRows.length, 2).setValues(actionRows);

  return dataRow + values.length - 1;
}

/**
 * Encodes a Resolve action's payload into a single pipe-delimited string
 * for storage in the hidden Trades column G. Each "side" is exactly 3
 * fields (list|set|part); the Inventory row to act on is looked up live
 * at resolve-time rather than cached, so the action stays correct even
 * if rows were inserted, deleted, or reordered on Inventory in between:
 *   "sell|A|Acceltra|Stock"
 *   "transfer|A|Acceltra|Stock|B|Acceltra|Stock"
 */
function encodeAction_(r) {
  function side(s) {
    return [s.list, s.set, s.part].join('|');
  }
  if (r.action === 'transfer') {
    return 'transfer|' + side(r.from) + '|' + side(r.to);
  }
  return 'sell|' + side(r.from);
}