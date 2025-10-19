# Koi+Stock Investment Workbook Spec 

## Overview
Excel/VBA workbook to aggregate crypto orders into positions, P&L, dashboard totals, and portfolio charts.

Core macros
- Refresh_Daily_Data: rebuilds Position table up to the cutoff, computes Cash/Coin/NAV, Deposit/Withdraw/Total PnL, formats the sheet, and updates portfolio charts (Cash vs Coin, Portfolio_Category_Daily, Alt.TOP, Alt.MID, Alt.LOW). Shows a single final message on completion.
- Update_MarketPrice_ByCutoff_OpenOnly_Simple: updates Market Price for Open rows only using cutoff rules (see Time & Cutoff). Stablecoins are priced at 1.
- Take_Daily_Snapshot: upserts a row per date into Daily_Snapshot (see layout below).
- Update_All_Snapshot: fills all missing daily snapshot rows from Daily_Snapshot!A2 (start date) to Position!B3 (cutoff). For each missing date it sets the cutoff, rebuilds Position (silent), writes the snapshot, then restores the original cutoff.
- Update_Dashboard: builds Dashboard charts (NAV with drawdown annotation, PnL, Deposit & Withdraw combined) for the date range on Dashboard!B2:B3.

Progress forms (modeless, auto-closed before final message)
- Update_Capital_and_Position: shows a progress form titled "Updating Capital and Position..." during the rebuild.
- Take_Daily_Snapshot: shows a progress form titled "Updating snapshot..." while taking the snapshot.
- Backfill (3M helper within Refresh_Daily_Data): shows a progress form titled "Updating missing snapshots..." while filling gaps.

## Sheets & Key Cells
### Position (SHEET_PORTFOLIO)
- B3 (CELL_CUTOFF): Cutoff datetime in UTC+0. If only a date is provided, treat as end-of-day 23:59:59 (UTC+0).
- Totals (cells configurable; current layout):
  - CELL_CASH = B7 (Cash)
  - CELL_COIN = B8 (Coin market value of open holdings)
  - CELL_NAV  = B9 (NAV = Cash + Coin)
  - CELL_NAV_ATH = B10 (3-month NAV high)
  - CELL_NAV_ATL = B11 (3-month NAV low)
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
    - Per-coin stacked amounts within each Alt group
    - Group matching is punctuation/spacing tolerant (e.g., "Alt MID" == "Alt.MID")

## Time & Cutoff Rules
- Order_History timestamps = UTC+0; no additional conversion applied.
- Cutoff read from Position!B3 (UTC+0). If date-only, treat as end-of-day 23:59:59.
- Pricing source & priority:
  - First try Binance:
    - If cutoff < today: Binance D1 close (UTC-aligned candle close)
    - If cutoff = today: Binance realtime ticker
  - If Binance lacks the symbol or returns no price: use the Exchange from Order_History (storage).
    - For historical cutoffs the exchange fetch uses the daily close; for the current day it uses realtime.
    - If neither source returns a price, the macro stops with "Can not fetch the ""coin"" price."
  - Stablecoins (USDT/USDC/BUSD/FDUSD/TUSD) = 1.

## Position Building (Refresh_Daily_Data)
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
- If the relative difference = `CAPITAL_RULE_DIFF_THRESHOLD_PCT` (default 0.5%), it writes "Check NAV !" to `CELL_NAV_ACTION=D9` in red/bold.
- To disable, clear `CELL_NAV_REAL`, increase the threshold, or remove the call to `checkCapitalRuleViolation` in `Refresh_Daily_Data`.

## Holdings Value
- Built from Position table (rows Open).
- Value = Available Balance, or Available Qty * Market Price when balance cell is absent.
- Aggregated for display and charting.

## Charts (updated automatically)
- From Refresh_Daily_Data (on Position sheet):
  - Cash vs Coin (pie)
  - Coin Category: pie by group (BTC, Alt.TOP, Alt.MID, Alt.LOW)
  - Alt.TOP (ChartObject name: "Alt.TOP"): pie by coin within Alt.TOP
  - Alt.MID (ChartObject name: "Alt.MID"): pie by coin within Alt.MID
  - Alt.LOW (ChartObject name: "Alt.LOW"): pie by coin within Alt.LOW
  - NAV 3M (line): last 3 months of NAV; legend hidden. X axis is Date (BaseUnit Days) using `dd/mm/yy` per `mod_config.POS_DATE_AXIS_FMT`. XValues bind to a worksheet helper date range, and the chart has `PlotVisibleOnly=False` so hidden helper columns still plot. Y-axis is scaled each run with ~5% margins around 3-month min/max, rounded to 500s.
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

## Metrics & Formulas
- Time & Cutoff
  - Cutoff cell `Position!B3` is UTC+0. A date-only value is treated as end-of-day 23:59:59 UTC+0.
  - Orders are already UTC+0; filtering uses the cutoff datetime (<= cutoff).

- Orders → Cash legs (prefer "Total" when present)
  - If `Total` column is present and numeric on the row, use it for cash math; else:
    - Buy cash  = Qty*Price + Fee
    - Sell cash = Qty*Price - Fee
  - Aggregates to the cutoff:
    - Total deposit  = sum of DEPOSIT amounts
    - Total withdraw = sum of WITHDRAW amounts
    - Total buy      = sum of BUY cash legs
    - Total sell     = sum of SELL cash legs

- Position Totals (Position sheet)
  - Holdings value (Coin) = sum of Available Balance for Open rows, where per row:
    - Available Balance =
      - If the `Available balance` column exists and is numeric: that value
      - Else: `Available qty` * `Market price`
  - Cash = (Total deposit + Total sell) - (Total buy + Total withdraw)
  - NAV  = Cash + Coin
  - Total profit = NAV - (Total deposit - Total withdraw)
  - Rounding/formatting: totals use integer money formats; NAV is truncated/rounded to 0 decimals per code and formatted with `MONEY_FMT`.

- Per-row PnL (Position table rows)
  - Closed rows: Profit = Sell proceeds - Cost
  - Open rows:   Profit = Sell proceeds + Available Balance - Cost
  - %PnL = Profit / Cost when Cost > 0, else blank; formatted with `PCT_FMT`.

- %NAV column (Position table rows)
  - Only meaningful for Open rows with AvailableQty > 0 and numeric Available Balance.
  - %NAV = Available Balance / total NAV (0 when NAV ≈ 0); formatted with `PCT_FMT`.
  - After computing, rows are sorted by %NAV descending in the output block.

- Allocation Metrics (Position sheet cells)
  - Build holdings by coin from Open rows’ Available Balance (see above).
  - Map coins to groups using the Category sheet (accepts both mapping and multi-column layouts; tolerant group name comparison).
  - %Coin = Coin / NAV (0 when NAV ≈ 0)
  - %BTC  = BTC group value / Coin (0 when Coin ≈ 0)
  - %Alt.TOP = Alt.TOP value / Coin; Count Alt.TOP = number of Alt.TOP coins with value > `EPS_CLOSE`
  - %Alt.MID = Alt.MID value / Coin; Count Alt.MID = number of Alt.MID coins with value > `EPS_CLOSE`
  - %Alt.LOW = Alt.LOW value / Coin; Count Alt.LOW = number of Alt.LOW coins with value > `EPS_CLOSE`
  - Percent cells are formatted using `PCT_FMT`.
  - Missing Category entries: if a coin is not found in the Category mapping (and is not BTC), the run appends a warning to the final message: `Warning: "<COIN>" is not in the Category` (one line per coin).

- NAV Metrics (3M window on Position)
  - Window: from 3 months prior to the cutoff day through the cutoff day.
  - Data source: `Daily_Snapshot` rows within the window; if the cutoff day’s snapshot row is missing, the live NAV from the current rebuild is included for ATH/ATL consideration.
  - NAV ATH/ATL cells hold the window’s max/min NAV (integer money convention).
  - Drawdown cell = max(0, (ATH - NAV_at_cutoff) / ATH). If `NAV_DD_USE_TRUNCATED` is True, ATH and NAV_at_cutoff are truncated to 0 decimals before computing. If the result ≤ `NAV_DD_TOLERANCE_PCT`, show 0%. Formatted with `PCT_FMT`.

- Daily_Snapshot row (A:L)
  - Date = cutoff date (UTC+0, date only)
  - Cash, Coin, NAV, Total deposit, Total withdraw, Total profit = values from Position totals (integer money)
  - BTC, Alt.TOP, Alt.MID, Alt.LOW = group totals (money) computed from holdings and Category mapping
  - Holdings = text string of per-coin holdings and values (used to recompute category splits when preferred)
  - UPSERT by Date; sheet sorted ascending by Date; formats per config: `SNAPSHOT_DATE_FMT` and `SNAPSHOT_NUMBER_FMT`.

- Capital rule check (Position sheet)
  - Compare `CELL_NAV` (calculated) vs `CELL_NAV_REAL` (user input). If `abs(calc - real) / abs(real) >= CAPITAL_RULE_DIFF_THRESHOLD_PCT`, write a warning message to `CELL_NAV_ACTION`.

## Version History
- v1.0
  - Initial consolidated specification for the Koi+Stock Investment Workbook (Position-focused).
  - Defines sheets, key cells, macros, chart contracts, configuration, and detailed metrics/formulas.
  - Standardizes Category mapping behavior, pricing priority (Binance → exchange fallback), snapshot layout (A:L), and Position charting (NAV 3M, group and per‑coin pies).
