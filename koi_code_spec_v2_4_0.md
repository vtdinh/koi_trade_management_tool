# KOI Trading Portfolio Workbook Spec (v2.4.0)
Generated: 2025-09-05

## Overview
Excel/VBA workbook to aggregate crypto orders into positions, P&L, dashboard totals, and portfolio charts.

Core macros
- Update_All_Position: rebuilds Position table up to the cutoff, computes Cash/Coin/NAV, Deposit/Withdraw/Total PnL, formats the sheet, and updates charts (Cash vs Coin, Portfolio1, Portfolio2). Shows a single final message on completion.
- Update_MarketPrice_ByCutoff_OpenOnly_Simple: updates Market Price for Open rows only using cutoff rules (see Time & Cutoff). Stablecoins are priced at 1.
- Take_Daily_Snapshot: upserts a row per date into Daily_Snapshot (see layout below).

## Sheets & Key Cells
### Position (SHEET_PORTFOLIO)
- B3 (CELL_CUTOFF): Cutoff datetime in UTC+7. If only a date is provided, treat as end-of-day 23:59:59 (UTC+7).
- Dashboard cells (configurable):
  - CELL_CASH = B5 (Cash)
  - CELL_COIN = B6 (Coin market value of open holdings)
  - CELL_NAV  = B7 (NAV = Cash + Coin)
  - CELL_SUM_DEPOSIT  = B8 (Total deposit to cutoff)
  - CELL_SUM_WITHDRAW = B9 (Total withdraw to cutoff)
  - CELL_TOTAL_PNL    = B10 (Total profit to cutoff)

### Order_History (SHEET_ORDERS)
- Default header row = 2 (auto-detection supported).
- Timestamps are UTC-4; workbook logic uses UTC+7 (+11h).
- Supported columns (case/spacing tolerant):
  - Required: Date | Type | Coin | Qty
  - Optional: Price | Fee | Exchange | Total
- If Total is present it is used for cash legs. Else:
  - Buy cash  = Qty*Price + Fee
  - Sell cash = Qty*Price - Fee

### Category (SHEET_CATEGORY)
- Sheet name default: "Catagory" (also accepts "Category"/"Categories").
- Two supported layouts:
  1) Mapping layout: row 1 headers "Coin | Group", rows below map each coin to a group name.
  2) Column layout: row 1 contains group names (e.g., BTC, Alt.TOP, Alt.MID, Alt.LOW); coins are listed beneath each column.

### Daily_Snapshot (SHEET_SNAPSHOT)
- Structure (A:L): Date | Cash | Coin | NAV | Total deposit | Total withdraw | Total profit | BTC | Alt.TOP | Alt.MID | Alt.LOW | Holdings
- UPSERT by Date; sorted ascending; formats: yyyy-mm-dd and #,##0; Holdings is plain text.

## Time & Cutoff Rules
- Order_History timestamps = UTC-4; converted to UTC+7 via +11h.
- Cutoff read from Position!B3 (UTC+7). If date-only, treat as end-of-day 23:59:59.
- Pricing:
  - If cutoff < today: fetch Binance D1 close.
  - If cutoff = today: fetch realtime ticker.
  - Fallback: SYMBOLUSDT -> SYMBOLUSDC -> SYMBOLBUSD.
  - Stablecoins (USDT/USDC/BUSD/FDUSD/TUSD) = 1.

## Position Building (Update_All_Position)
1) Map headers and clear old output.
2) Iterate orders <= cutoff (UTC+7), maintain per-coin session state:
   - BUY: extend session; Cost += Qty*Price + Fee; BuyQty += Qty.
   - SELL: extend session; SellProceeds += Qty*Price - Fee; SellQty += Qty; close when AvailableQty≈0.
   - DEPOSIT/WITHDRAW: affect only cash aggregates.
3) Flush open sessions; compute AvailableQty.
4) Pre-fetch market prices for Open coins; stablecoins = 1.
5) Write sessions to Position table: Open/Closed, Qty, Cost, Proceeds, Avg, Profit, %PnL, Storage; color PnL (green/red).
6) Formats: dates yyyy-mm-dd; %PnL "0.00%"; money #,##0; price #,##0.00; AutoFit; clear trailing rows.

## Dashboard Metrics
- Cash = (ΣDeposit + ΣSell) − (ΣBuy + ΣWithdraw)
- Coin = Σ AvailableQty_open × MarketPrice (per open coin)
- NAV  = Cash + Coin
- Total deposit = ΣDeposit
- Total withdraw = ΣWithdraw
- Total profit   = NAV − (Total deposit − Total withdraw)

## Holdings Value
- Built from Position table (rows Open).
- Value = Available Balance, or Available Qty × Market Price when balance cell is absent.
- Aggregated for display and charting.

## Charts (updated automatically by Update_All_Position)
- Cash vs Coin
  - Location: Position sheet, ChartObject named "Cash vs Coin" (found by name or title).
  - Behavior: chart type preserved (pie recommended); labels show Category + Percentage; values hidden. Negative values are skipped.
  - If no holdings (Coin=0), chart resets to a single "No holdings" slice.

- Portfolio1 (Group breakdown)
  - Location: Position sheet (ChartObject) or chart sheet named/titled "Portfolio1". If missing, created on Position.
  - Groups: BTC, Alt.TOP, Alt.MID, Alt.LOW.
  - Mapping source: Category sheet (either mapping or column layout). Unmapped non-BTC coins default to Alt.LOW.
  - Labels: Category + Percentage; values hidden; "0%" format.
  - If all group values are zero, chart resets to a single "No holdings" slice.

- Portfolio2 (Per-coin breakdown)
  - Location: Position sheet (ChartObject) or chart sheet named/titled "Portfolio2". If missing, created on Position.
  - Data: each open coin's Available Balance share.
  - Labels: Category + Percentage; values hidden; "0%" format.
  - If total is zero, chart resets to a single "No holdings" slice.

## Symbols & Mapping
- Symbol mapping: COIN -> COINUSDT (unless already ending with USDT/USDC/BUSD). Fallback to USDC/BUSD when USDT unavailable.
- Stablecoins recognized: USDT, USDC, BUSD, FDUSD, TUSD.

## Error Handling & Safety
- Graceful messages for missing sheets/headers/data.
- CLEAR_MARKET_PRICE=True: clears Market Price cells unless overwritten during rebuild.
- Robust last-row detection and header auto-detection for both Position and Order_History.
- Update_All_Position shows only one final message on completion.

## Version History
- v2.4.0: Added automatic chart updates (Cash vs Coin, Portfolio1 groups, Portfolio2 per-coin), Category sheet support for both layouts, single final message from Update_All_Position. Daily_Snapshot expanded to include group totals and Holdings string.
- v2.3.0: Added Deposit/Withdraw/Total P&L aggregates, expanded Daily_Snapshot, clarified pricing/timezone.
- v2.2.x: Position builder, cutoff normalization, Binance pricing, %PnL coloring, header auto-detection, stablecoin handling.

