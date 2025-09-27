# KOI Trading Portfolio Workbook Spec (v2.7.0)
Generated: 2025-09-10

## Overview
Excel/VBA workbook to aggregate crypto orders into positions, P&L, dashboard totals, and portfolio charts.

Core macros
- Update_All_Position: rebuilds Position table up to the cutoff, computes Cash/Coin/NAV, Deposit/Withdraw/Total PnL, formats the sheet, and updates portfolio charts (Cash vs Coin, Portfolio_Category_Daily, Portfolio_Alt.TOP_Daily, Portfolio_Alt.MID_Daily, Portfolio_Alt.LOW_Daily). Shows a single final message on completion.
- Update_MarketPrice_ByCutoff_OpenOnly_Simple: updates Market Price for Open rows only using cutoff rules (see Time & Cutoff). Stablecoins are priced at 1.
- Take_Daily_Snapshot: upserts a row per date into Daily_Snapshot (see layout below).
- Update_All_Snapshot: fills all missing daily snapshot rows from Daily_Snapshot!A2 (start date) to Position!B3 (cutoff). For each missing date it sets the cutoff, rebuilds Position (silent), writes the snapshot, then restores the original cutoff.
- Update_Dashboard: builds Dashboard charts (NAV with drawdown annotation, PnL, Deposit & Withdraw combined) for the date range on Dashboard!B2:B3.

## Sheets & Key Cells
### Position (SHEET_PORTFOLIO)
- B3 (CELL_CUTOFF): Cutoff datetime in UTC+0. If only a date is provided, treat as end-of-day 23:59:59 (UTC+0).
- Totals (cells configurable; current layout):
  - CELL_CASH = B7 (Cash)
  - CELL_COIN = B8 (Coin market value of open holdings)
  - CELL_NAV  = B9 (NAV = Cash + Coin)
  - CELL_NAV_ATH = B10 (3‑month NAV high)
  - CELL_NAV_ATL = B11 (3‑month NAV low)
  - CELL_NAV_DD  = B12 (drawdown vs ATH at cutoff)
  - CELL_SUM_DEPOSIT  = B13 (Total deposit to cutoff)
  - CELL_SUM_WITHDRAW = B14 (Total withdraw to cutoff)
  - CELL_TOTAL_PNL    = B15 (Total profit to cutoff)
- Allocation metrics (percent cells use PCT_FMT):
  - CELL_PCT_COIN = B18 (%Coin = Coin/NAV)
  - CELL_PCT_BTC  = B19 (%BTC of Coin)
  - CELL_PCT_ALT_TOP = B20, CELL_NUM_ALT_TOP = B21
  - CELL_PCT_ALT_MID = B22, CELL_NUM_ALT_MID = B23
  - CELL_PCT_ALT_LOW = B24, CELL_NUM_ALT_LOW = B25

### Order_History (SHEET_ORDERS)
- Default header row = 2 (auto-detection supported).
- Timestamps are already UTC+0; workbook logic uses them directly.
- Supported columns (case/spacing tolerant):
  - Required: Date | Type | Coin | Qty
  - Optional: Price | Fee | Exchange | Total
- If Total is present it is used for cash legs. Else:
  - Buy cash  = Qty*Price + Fee
  - Sell cash = Qty*Price - Fee

### Categoty / Catagory / Category (SHEET_CATEGORY)
- Sheet name default: "Categoty" (current, intentional), also accepts "Catagory" and "Category".
- Supported layouts:
  1) Mapping layout: row 1 headers "Coin | Group"; rows below map each coin to a group name.
  2) Column layout: row 1 contains group names (e.g., BTC, Alt.TOP, Alt.MID, Alt.LOW); coins are listed beneath each column under the appropriate group.

### Daily_Snapshot (SHEET_SNAPSHOT)
- Structure (A:L): Date | Cash | Coin | NAV | Total deposit | Total withdraw | Total profit | BTC | Alt.TOP | Alt.MID | Alt.LOW | Holdings
- UPSERT by Date; sorted ascending; formats: yyyy-mm-dd and #,##0; Holdings is plain text.

#### Bulk Snapshot Update
- Start date: place the first date to backfill in `Daily_Snapshot!A2`.
- Cutoff date: set target date in `Position!B3`.
- Run `Update_All_Snapshot`: creates rows for any missing dates between A2 and the cutoff. Existing dates are left unchanged.

### Dashboard (SHEET_DASHBOARD)
- Inputs:
  - B2: Start date
  - B3: End date
- Charts built by Update_Dashboard:
  - NAV (ChartObject name: "NAV")
    - Series: Daily_Snapshot Date (A) vs NAV (D)
    - Drawdown textbox: "Drawdown: Max. xx%, Current: xx%"
      - Max drawdown = largest peak-to-trough decline over the range
      - Current drawdown = drop from all-time peak to the latest value (0 if latest equals all-time high)
    - X axis: time scale; date format = SNAPSHOT_DATE_FMT; Y format = MONEY_FMT
  - PnL (ChartObject name: "PnL")
    - Series: Date (A) vs Total profit (G)
    - X axis: time scale; Y axis always placed at bottom
    - No drawdown annotation on PnL (only NAV has drawdown text)
  - Deposit & Withdraw (ChartObject name: "Deposit")
    - Series 1: Date (A) vs Total deposit (E)
    - Series 2: Date (A) vs Total withdraw (F)
    - Legend enabled; no drawdown annotation
  - Cash vs NAV (ChartObject name: "Cash vs NAV")
    - Series: Date (A) vs Cash/NAV ratio (percent)
    - Value axis as percent (e.g., 0–100%)
  - Portfolio_Category (ChartObject name: "Portfolio_Category")
    - Preferred computation: parse `Holdings` and map coins to groups using Catagory sheet; stack amounts per group (BTC, Alt.TOP, Alt.MID, Alt.LOW, then Others)
    - Fallback: if no `Holdings` column, use snapshot group columns (H..)
  - Portfolio_Alt.TOP / Portfolio_Alt.MID / Portfolio_Alt.LOW
    - Per‑coin stacked amounts within each Alt group
    - Group matching is punctuation/spacing tolerant (e.g., "Alt MID" == "Alt.MID")

## Time & Cutoff Rules
- Order_History timestamps = UTC+0; no additional conversion applied.
- Cutoff read from Position!B3 (UTC+0). If date-only, treat as end-of-day 23:59:59.
- Pricing source & priority:
  - First try Binance:
    - If cutoff < today: Binance D1 close (UTC-aligned candle close)
    - If cutoff = today: Binance realtime ticker
    - Fallback quote: SYMBOLUSDT -> SYMBOLUSDC -> SYMBOLBUSD
  - If Binance has no price/symbol: use Exchange from Order_History (storage) for realtime price
    - Supported: OKX, Bybit (spot ticker realtime)
    - For historical dates on non-Binance, realtime is used as a safe fallback
  - Stablecoins (USDT/USDC/BUSD/FDUSD/TUSD) = 1.

## Position Building (Update_All_Position)
1) Map headers and clear old output.
2) Iterate orders <= cutoff (UTC+0), maintain per-coin session state:
   - BUY: extend session; Cost += Qty*Price + Fee; BuyQty += Qty.
   - SELL: extend session; SellProceeds += Qty*Price - Fee; SellQty += Qty; close when AvailableQty ~ 0.
   - DEPOSIT/WITHDRAW: affect only cash aggregates.
3) Flush open sessions; compute AvailableQty.
4) Pre-fetch market prices for Open coins; stablecoins = 1.
5) Write sessions to Position table: Open/Closed, Qty, Cost, Proceeds, Avg, Profit, %PnL, %NAV (open rows), Storage; color PnL (green/red).
   - %NAV = Available Balance (open row) / total NAV (closed rows 0%).
6) Formats: dates yyyy-mm-dd; %PnL "0.00%"; money #,##0; price #,##0.00; AutoFit; clear trailing rows.

## Dashboard Metrics
- Cash = (sum Deposit + sum Sell) - (sum Buy + sum Withdraw)
- Coin = sum AvailableQty_open * MarketPrice (per open coin)
- NAV  = Cash + Coin
- Total deposit = sum Deposit
- Total withdraw = sum Withdraw
- Total profit   = NAV - (Total deposit - Total withdraw)
 
## NAV Sanity Check
- The workbook optionally compares calculated NAV (Position `CELL_NAV=B9`) with a manual “Real NAV” (Position `CELL_NAV_REAL=C9`).
- If the relative difference ≥ `CAPITAL_RULE_DIFF_THRESHOLD_PCT` (default 0.5%), it writes "Check NAV !" to `CELL_NAV_ACTION=D9` in red/bold.
- To disable, clear `CELL_NAV_REAL`, increase the threshold, or remove the call to `checkCapitalRuleViolation` in `Update_All_Position`.

## Holdings Value
- Built from Position table (rows Open).
- Value = Available Balance, or Available Qty * Market Price when balance cell is absent.
- Aggregated for display and charting.

## Charts (updated automatically)
- From Update_All_Position (on Position sheet):
  - Cash vs Coin (pie)
  - Coin Category: pie by group (BTC, Alt.TOP, Alt.MID, Alt.LOW)
  - Alt.TOP: pie by coin within Alt.TOP
  - Alt.MID: pie by coin within Alt.MID
  - Alt.LOW: pie by coin within Alt.LOW
  - NAV 3M (line): last 3 months of NAV; legend hidden. X axis is Date (BaseUnit Days) using `dd/mm/yy` per `mod_config.POS_DATE_AXIS_FMT`. XValues bind to a worksheet helper date range, and the chart has `PlotVisibleOnly=False` so hidden helper columns still plot. Y-axis is scaled each run with ~10% margins around 3‑month min/max, rounded to thousands.
- From Update_Dashboard (on Dashboard sheet):
  - NAV with drawdown annotation; PnL; Deposit & Withdraw combined.

## Configuration (mod_config)
- Sheet names:
  - SHEET_PORTFOLIO = "Position"
  - SHEET_ORDERS    = "Order_History"
  - SHEET_SNAPSHOT  = "Daily_Snapshot"
  - SHEET_CATEGORY  = "Categoty"  (fallback to "Catagory" and "Category" accepted)
  - CHART_PORTFOLIO1 = "Coin Category"
  - AUTOFIT_POSITION_COLUMNS = False (keep user column widths)
  - PCT_FMT = "0.0%" (one decimal for percents)
  - POS_DATE_AXIS_FMT = "dd/mm/yy" (Position NAV 3M axis tick labels)
  - CAPITAL_RULE_DIFF_THRESHOLD_PCT = 0.005 (0.5% NAV check threshold)
  - CELL_NAV_REAL = C9; CELL_NAV_ACTION = D9 (Position sheet addresses for NAV sanity check)

### Formatting Alignment
- Position sheet number formats:
  - The columns “Buy Qty”, “Sell Qty”, and “Available Qty” mirror the NumberFormat of the “Qty” column in `Order_History` (first numeric cell below header). If unavailable, fall back to the column format or a default based on `ROUND_QTY_DECIMALS`.
  - SHEET_DASHBOARD = "Dashboard"
- Numbers & formats: DATE_FMT, MONEY_FMT, PRICE_FMT, PCT_FMT, SNAPSHOT_DATE_FMT, SNAPSHOT_NUMBER_FMT
- Tolerances: EPS_ZERO, EPS_CLOSE
- NAV drawdown textbox (Dashboard/NAV):
  - NAV_MDD_ANCHOR: "UnderTitle" | "PlotTopRight" | "PlotTopLeft"
  - NAV_MDD_OFFSET_X / NAV_MDD_OFFSET_Y
  - NAV_MDD_WIDTH / NAV_MDD_HEIGHT
  - NAV_MDD_ALIGN: "Center" | "Left" | "Right"

## Version History
- v2.7.0:
  - Pricing priority: Binance first (D1 close/realtime), then Exchange-specific pricing (OKX/Bybit) using\r\n    daily close for past cutoffs and realtime for today.
  - Position charts: added three daily pies — `Portfolio_Alt.TOP_Daily`, `Portfolio_Alt.MID_Daily`, `Portfolio_Alt.LOW_Daily` — showing per-coin breakdowns within Alt groups.
  - Removed per-coin pie `Portfolio_Coin` from Position.
  - Avg. cost and avg sell price now rounded using `ROUND_PRICE_DECIMALS` instead of 0 decimals.
  - Added `NAV 3M` chart (line) with date axis; legend hidden; auto‑scaled Y axis with 10% margin around 3M min/max, rounded to thousands.
  - Added NAV metrics cells: NAV ATH/ATL/Drawdown (3M window) and allocation metrics (%Coin, %BTC, %Alt.*, counts).
  - Added `%NAV` column calculation for open rows (Available Balance / total NAV).
- v2.6.2:
  - Position: renamed charts — `Portfolio1` → `Portfolio_Category_Daily`; `Portfolio2` → `Portfolio_Coin`.
  - Position: quantity display format for “Buy Qty”, “Sell Qty”, and “Available Qty” now syncs with the `Order_History!Qty` column format.
- v2.6.1:
  - Default Category sheet renamed to "Catagory" (intentional spelling); code accepts both "Catagory" and "Category".
  - Dashboard: added charts Cash vs NAV and Portfolio_Category (replaces legacy Portfolio_Group).
  - Portfolio_Category prefers computing from Holdings + Catagory mapping; falls back to snapshot columns.
  - Alt.* charts: tolerant group name comparison; avoid a lone "Other" by keeping all coins when all slices are tiny.
  - GetOrCreateChart finds and renames legacy/typo chart names to the canonical ones.
  - Kept drawdown annotation only on NAV; PnL drawdown removed.
- v2.6.0:
  - New Update_Dashboard with three charts: NAV (with drawdown annotation), PnL, and combined Deposit & Withdraw.
  - Current drawdown = 0 when latest NAV equals all‑time high (tolerant by EPS_CLOSE).
  - PnL X axis always at bottom.
  - Added SHEET_DASHBOARD and NAV_MDD_* config; default SHEET_CATEGORY older spec used "Category".
- v2.5.0: Added Update_All_Snapshot (bulk daily backfill), standardized Daily_Snapshot A:L layout with group totals + Holdings string, chart reset to "No holdings" when empty, single-message run mode.
- v2.4.x: Automatic chart updates (Cash vs Coin, Portfolio1 groups, Portfolio2 per-coin), Category sheet dual-layout support.
- v2.3.0: Deposit/Withdraw/Total P&L aggregates, expanded Daily_Snapshot, clarified pricing/timezone.
- v2.2.x: Position builder, cutoff normalization, Binance pricing, %PnL coloring, header auto-detection, stablecoin handling.
