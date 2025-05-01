// Branch: feature/seeds-used-cost - Attribute only the cost of seeds actually planted (Patches) to the cycle
// Branch: feature/formatting - Add number formatting to the farming-cycles summary sheet
// Branch: feature/formatting-transactions - Add number formatting to the farming-transactions sheet
// Branch: feature/conditional-formatting - Add conditional formatting to farming-cycles and farming-transactions
// Branch: feature/custom-menu-and-status - Add a custom menu for update/summarize, and display last refresh and last upload times
// Branch: fix/header-detection - Dynamically detect the header row in summarizeCycles
// Branch: feature/info-next-to-headers - Move data refresh info to columns after the last header in farming-transactions
/**
 * Grand Exchange Farming Dashboard Script
 * 
 * Main entry point for updating the farming profit dashboard.
 * Imports utilities, mapping, and farming logic.
 * Handles data loading, grouping by run, and sheet updates.
 */

// === CONFIGURATION ===
const RUNELITE_JSON_FILENAME = 'grand-exchange.json';
const RUNELITE_JSON_FOLDER_ID = '1XQJ9bIWpawauW7H8n89GECWQlWfV51z4';

// Import utility, mapping, and farming functions
// (Google Apps Script loads all .gs files in the project, so just use functions directly)

function loadRuneLiteJSONFromFolder(filename, folderId) {
  Logger.log('Looking for file: ' + filename + ' in folder: ' + folderId);
  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFilesByName(filename);
    if (!files.hasNext()) {
      throw new Error('Could not find ' + filename + ' in folder ' + folderId);
    }
    var file = files.next();
    var content = file.getBlob().getDataAsString();
    var jsonData = JSON.parse(content);
    return jsonData;
  } catch (error) {
    Logger.log('Error loading RuneLite JSON from folder: ' + error.toString());
    return [];
  }
}

function getLastJsonUploadTime() {
  try {
    var folder = DriveApp.getFolderById(RUNELITE_JSON_FOLDER_ID);
    var files = folder.getFilesByName(RUNELITE_JSON_FILENAME);
    if (!files.hasNext()) return '';
    var file = files.next();
    return file.getLastUpdated();
  } catch (e) {
    return '';
  }
}

function updateFarmingDashboard() {
  // 1. Load JSON data from specific folder in Drive
  var transactions = loadRuneLiteJSONFromFolder(RUNELITE_JSON_FILENAME, RUNELITE_JSON_FOLDER_ID);
  // 2. Fetch item mappings
  var apiMappings = fetchItemMappings();
  // 3. Process transactions (normalize, add profit/loss)
  var processed = processTransactions(transactions, apiMappings);
  // 4. Sort by time (newest to oldest) and limit to last 50
  processed.sort(function(a, b) { return b.time - a.time; });
  var limited = processed.slice(0, 50);
  // 5. Output to sheet with Cycle ID, Patches, Dead Patches columns
  var sheetName = 'farming-transactions';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  // Set headers (add Cycle ID, Patches, Dead Patches)
  var headers = ['Cycle ID', 'Patches', 'Dead Patches', 'Timestamp', 'Item Name', 'Buy/Sell', 'Quantity', 'Price', 'Profit/Loss'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  // Prepare data rows (Cycle ID, Patches, Dead Patches left blank for user to fill in)
  var rows = limited.map(function(tx) {
    return [
      '', // Cycle ID (user to fill in)
      '', // Patches (user to fill in)
      '', // Dead Patches (user to fill in)
      tx.time instanceof Date ? Utilities.formatDate(tx.time, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss') : tx.time,
      tx.itemName,
      tx.buy ? 'Buy' : 'Sell',
      tx.quantity,
      tx.price,
      tx.profitLoss
    ];
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.autoResizeColumns(1, headers.length);
    // Apply formatting: Quantity, Patches, Dead Patches as integer; Price and Profit/Loss as GP
    var numRows = rows.length;
    // Patches, Dead Patches, Quantity: integer with comma (cols 2, 3, 7)
    [2, 3, 7].forEach(function(col) {
      sheet.getRange(2, col, numRows, 1).setNumberFormat('#,##0');
    });
    // Price, Profit/Loss: GP (cols 8, 9)
    [8, 9].forEach(function(col) {
      sheet.getRange(2, col, numRows, 1).setNumberFormat('"₲"#,##0');
    });
    // Conditional formatting: Profit/Loss (col 9)
    var profitLossRange = sheet.getRange(2, 9, numRows, 1);
    var rules = sheet.getConditionalFormatRules();
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#d9ead3') // light green
      .setRanges([profitLossRange])
      .build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#f4cccc') // light red
      .setRanges([profitLossRange])
      .build());
    sheet.setConditionalFormatRules(rules);
  }
  // Move data refresh info to columns after the last header
  var now = new Date();
  var lastUpload = getLastJsonUploadTime();
  // Find the last header column
  var infoCol = headers.length + 1;
  // Write info in row 1, columns after headers
  sheet.getRange(1, infoCol, 1, 2).setValues([["Last Data Refresh:", Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss')]]);
  sheet.getRange(2, infoCol, 1, 2).setValues([["Last grand-exchange.json Upload:", lastUpload ? Utilities.formatDate(lastUpload, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss') : 'Not found']]);
  sheet.getRange(1, infoCol, 2, 2).setFontWeight('bold');
  // Optionally, set a background color for clarity
  sheet.getRange(1, infoCol, 2, 2).setBackground('#f3f3f3');
  Logger.log('Wrote ' + rows.length + ' transactions to sheet: ' + sheetName);
}

/**
 * Summarizes cycles based on Cycle ID in the farming-transactions sheet.
 * Outputs summary to a new sheet 'farming-cycles'.
 * Uses user-entered Patches and Dead Patches columns.
 * Attributes only the cost of seeds actually planted (Patches) to the cycle.
 */
function summarizeCycles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var txSheet = ss.getSheetByName('farming-transactions');
  if (!txSheet) {
    Logger.log('No farming-transactions sheet found.');
    return;
  }
  var data = txSheet.getDataRange().getValues();
  // Dynamically find the header row (look for row starting with 'Cycle ID')
  var headerRowIdx = -1;
  for (var i = 0; i < Math.min(10, data.length); i++) {
    if (data[i][0] && data[i][0].toString().trim() === 'Cycle ID') {
      headerRowIdx = i;
      break;
    }
  }
  if (headerRowIdx === -1) {
    Logger.log('Could not find header row in farming-transactions.');
    return;
  }
  var headers = data[headerRowIdx];
  var rows = data.slice(headerRowIdx + 1);
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });
  var cycles = {};
  rows.forEach(function(row) {
    var cycleId = row[idx['Cycle ID']];
    if (!cycleId) return; // skip if no Cycle ID
    if (!cycles[cycleId]) cycles[cycleId] = [];
    cycles[cycleId].push(row);
  });
  // Prepare summary rows
  var summary = [['Cycle ID', 'Crop', 'Seeds Bought', 'Seeds Used', 'Harvest Sold', 'Patches', 'Dead Patches', 'Effective Patches', 'Total Cost (Used Seeds)', 'Total Revenue', 'Net Profit', 'Avg Yield/Effective Patch']];
  Object.keys(cycles).forEach(function(cycleId) {
    var txs = cycles[cycleId];
    var crop = '';
    var seeds = 0, harvest = 0, cost = 0, revenue = 0;
    var patches = null, deadPatches = null;
    var seedBuyCost = 0;
    txs.forEach(function(row) {
      var type = row[idx['Buy/Sell']];
      var qty = Number(row[idx['Quantity']]);
      var price = Number(row[idx['Price']]);
      var profit = Number(row[idx['Profit/Loss']]);
      var item = row[idx['Item Name']];
      // Use first non-blank Patches/Dead Patches value in the cycle
      if (patches === null && row[idx['Patches']] !== '') patches = Number(row[idx['Patches']]);
      if (deadPatches === null && row[idx['Dead Patches']] !== '') deadPatches = Number(row[idx['Dead Patches']]);
      if (type === 'Buy') {
        seeds += qty;
        seedBuyCost += -profit; // profit is negative for buys
        crop = item; // assume crop name from seed
      } else {
        harvest += qty;
        revenue += profit;
        crop = item; // assume crop name from harvest
      }
    });
    // Default patches to seeds if not set
    if (patches === null) patches = seeds;
    if (deadPatches === null) deadPatches = 0;
    var effectivePatches = Math.max(0, patches - deadPatches);
    // Attribute only the cost of seeds actually planted (patches)
    var seedsUsed = Math.min(seeds, patches);
    var costPerSeed = seeds > 0 ? (seedBuyCost / seeds) : 0;
    var usedSeedCost = costPerSeed * seedsUsed;
    var net = revenue - usedSeedCost;
    var avgYield = effectivePatches > 0 ? (harvest / effectivePatches) : 0;
    summary.push([cycleId, crop, seeds, seedsUsed, harvest, patches, deadPatches, effectivePatches, usedSeedCost, revenue, net, avgYield]);
  });
  // Output to farming-cycles sheet
  var cycleSheet = ss.getSheetByName('farming-cycles');
  if (!cycleSheet) {
    cycleSheet = ss.insertSheet('farming-cycles');
  } else {
    cycleSheet.clear();
  }
  cycleSheet.getRange(1, 1, summary.length, summary[0].length).setValues(summary);
  cycleSheet.autoResizeColumns(1, summary[0].length);

  // Apply formatting: GP columns as currency, numbers with commas, decimals for yield
  if (summary.length > 1) {
    var numRows = summary.length - 1;
    // Seeds Bought, Seeds Used, Harvest Sold, Patches, Dead Patches, Effective Patches: integer with comma
    var intCols = [3, 4, 5, 6, 7, 8];
    intCols.forEach(function(col) {
      cycleSheet.getRange(2, col, numRows, 1).setNumberFormat('#,##0');
    });
    // GP columns: Total Cost, Total Revenue, Net Profit (cols 9, 10, 11)
    [9, 10, 11].forEach(function(col) {
      cycleSheet.getRange(2, col, numRows, 1).setNumberFormat('"₲"#,##0');
    });
    // Avg Yield/Effective Patch: 2 decimals
    cycleSheet.getRange(2, 12, numRows, 1).setNumberFormat('#,##0.00');
    // Conditional formatting: Net Profit (col 11)
    var netProfitRange = cycleSheet.getRange(2, 11, numRows, 1);
    var rules = cycleSheet.getConditionalFormatRules();
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#d9ead3') // light green
      .setRanges([netProfitRange])
      .build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#f4cccc') // light red
      .setRanges([netProfitRange])
      .build());
    cycleSheet.setConditionalFormatRules(rules);
  }
  Logger.log('Wrote ' + (summary.length - 1) + ' cycle summaries to sheet: farming-cycles');
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Farming Dashboard')
    .addItem('Update Transactions', 'updateFarmingDashboard')
    .addItem('Summarize Cycles', 'summarizeCycles')
    .addToUi();
} 