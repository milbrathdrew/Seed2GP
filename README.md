# SproutLedger

**SproutLedger** is a Google Sheets-based Grand Exchange farming profit tracker for Old School RuneScape. It is designed for players who want to track their farming cycles, seed costs, harvest sales, and overall profitability—especially when other trackers (like ge-tracker) don't fit their needs.

## Features (MVP)
- **Automated import** of RuneLite `grand-exchange.json` from Google Drive
- **Transaction normalization**: item names, buy/sell, quantity, price, profit/loss
- **Manual cycle matching**: assign a Cycle ID to link seed buys and harvest sales
- **Custom patch and dead patch tracking**: enter patches planted and dead patches for each cycle
- **Cycle summary sheet**: see total cost, revenue, net profit, and average yield per effective patch
- **Beautiful formatting**: GP columns, yields, and conditional coloring for profit/loss
- **Custom menu**: update transactions and summarize cycles with a click
- **Sheet info**: see last data refresh and last JSON upload time at a glance

## How It Works
1. **Export your `grand-exchange.json`** from RuneLite and upload it to the configured Google Drive folder.
2. **Open your bound Google Sheet** and use the "Farming Dashboard" menu to update transactions.
3. **Assign Cycle IDs** and fill in Patches/Dead Patches in the `farming-transactions` sheet.
4. **Click "Summarize Cycles"** to generate a summary in the `farming-cycles` sheet.

## Sheet Structure
- **farming-transactions**: All recent transactions, ready for manual cycle matching and patch entry.
- **farming-cycles**: Summary of each cycle, with profit, yield, and patch info.

## Restoring This MVP
- All core logic is in `code.gs`, with supporting files for utilities and mapping.
- If future changes break functionality, restore these files from this version.
- This README and the script files can be stored in a backup folder or version control for reference.

## Requirements
- Google Sheets with bound Apps Script project
- Access to your RuneLite `grand-exchange.json` in Google Drive

## Why SproutLedger?
SproutLedger is a playful, powerful, and customizable alternative to commercial GE trackers—built for real farming cycles and real RuneScape players.

---

*This is the MVP reference version. For new features or major changes, refer back to this README and script set to restore working functionality if needed.* 