/**
 * Utilities for Grand Exchange Farming Dashboard
 *
 * Contains helper functions for profit calculation, normalization, grouping, and linking transactions to runs.
 */

/**
 * Normalizes a single transaction object.
 * Converts types, standardizes keys, and adds readable item name.
 */
function normalizeTransaction(tx, apiMappings) {
  return {
    quantity: Number(tx.quantity),
    price: Number(tx.price),
    itemId: Number(tx.itemId),
    buy: Boolean(tx.buy),
    time: new Date(Number(tx.time)),
    itemName: getDisplayName(Number(tx.itemId), apiMappings)
  };
}

/**
 * Calculates profit or loss for a transaction (positive for sales, negative for buys).
 */
function calculateProfitLoss(tx) {
  var profit = tx.quantity * tx.price;
  return tx.buy ? -profit : profit;
}

/**
 * Processes an array of raw transactions: normalizes and adds profit/loss.
 */
function processTransactions(rawTxs, apiMappings) {
  return rawTxs.map(function(tx) {
    var norm = normalizeTransaction(tx, apiMappings);
    norm.profitLoss = calculateProfitLoss(norm);
    return norm;
  });
}

// Placeholder: Group transactions by run
function groupTransactionsByRun(transactions) {
  // ...
}

// Placeholder: Link transactions to farming runs
function linkTransactionsToRuns(transactions, runs) {
  // ...
} 