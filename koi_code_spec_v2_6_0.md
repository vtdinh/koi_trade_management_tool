# KOI Trading Portfolio Workbook Spec (v2.6.0)
Generated: 2025-09-06

## Overview
Excel/VBA workbook to aggregate crypto orders into positions, P&L, dashboard totals, and portfolio charts.

Core macros
- Update_All_Position: rebuilds Position table up to the cutoff, computes Cash/Coin/NAV, Deposit/Withdraw/Total PnL, formats the sheet, and updates portfolio charts (Cash vs Coin, Portfolio1, Portfolio2). Shows a single final message on completion.
- Update_MarketPrice_ByCutoff_OpenOnly_Simple: updates Market Price for Open rows only using cutoff rules (see Time & Cutoff). Stablecoins are priced at 1.
- Take_Daily_Snapshot: upserts a row per date into Daily_Snapshot (see layout below).
- Update_All_Snapshot: fills all missing daily snapshot rows from Daily_Snapshot!A2 (start date) to Position!B3 (cutoff). For each missing date it sets the cutoff, rebuilds Position (silent), writes the snapshot, then restores the original cutoff.
- Update_Dashboard: builds Dashboard charts (NAV with drawdown annotation, PnL, Deposit & Withdraw combined) for the date range on Dashboard!B2:B3.

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
  - Portfolio_Catagory (ChartObject name: "Portfolio_Catagory")
    - Preferred computation: parse `Holdings` and map coins to groups using Catagory sheet; stack amounts per group (BTC, Alt.TOP, Alt.MID, Alt.LOW, then Others)
    - Fallback: if no `Holdings` column, use snapshot group columns (H..)
  - Portfolio_Alt.TOP / Portfolio_Alt.MID / Portfolio_Alt.LOW
    - Per‑coin stacked amounts within each Alt group
    - Group matching is punctuation/spacing tolerant (e.g., "Alt MID" == "Alt.MID")

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
   - SELL: extend session; SellProceeds += Qty*Price - Fee; SellQty += Qty; close when AvailableQty ~ 0.
   - DEPOSIT/WITHDRAW: affect only cash aggregates.
3) Flush open sessions; compute AvailableQty.
4) Pre-fetch market prices for Open coins; stablecoins = 1.
5) Write sessions to Position table: Open/Closed, Qty, Cost, Proceeds, Avg, Profit, %PnL, Storage; color PnL (green/red).
6) Formats: dates yyyy-mm-dd; %PnL "0.00%"; money #,##0; price #,##0.00; AutoFit; clear trailing rows.

## Dashboard Metrics
- Cash = (sum Deposit + sum Sell) - (sum Buy + sum Withdraw)
- Coin = sum AvailableQty_open * MarketPrice (per open coin)
- NAV  = Cash + Coin
- Total deposit = sum Deposit
- Total withdraw = sum Withdraw
- Total profit   = NAV - (Total deposit - Total withdraw)

## Holdings Value
- Built from Position table (rows Open).
- Value = Available Balance, or Available Qty * Market Price when balance cell is absent.
- Aggregated for display and charting.

## Charts (updated automatically)
- From Update_All_Position (on Position sheet):
  - Cash vs Coin, Portfolio1, Portfolio2 — behavior unchanged from v2.5.0.
- From Update_Dashboard (on Dashboard sheet):
  - NAV with drawdown annotation; PnL; Deposit & Withdraw combined.

## Configuration (mod_config)
- Sheet names:
  - SHEET_PORTFOLIO = "Position"
  - SHEET_ORDERS    = "Order_History"
  - SHEET_SNAPSHOT  = "Daily_Snapshot"
  - SHEET_CATEGORY  = "Catagory"  (fallback to "Category" accepted)
  - SHEET_DASHBOARD = "Dashboard"
- Numbers & formats: DATE_FMT, MONEY_FMT, PRICE_FMT, PCT_FMT, SNAPSHOT_DATE_FMT, SNAPSHOT_NUMBER_FMT
- Tolerances: EPS_ZERO, EPS_CLOSE
- NAV drawdown textbox (Dashboard/NAV):
  - NAV_MDD_ANCHOR: "UnderTitle" | "PlotTopRight" | "PlotTopLeft"
  - NAV_MDD_OFFSET_X / NAV_MDD_OFFSET_Y
  - NAV_MDD_WIDTH / NAV_MDD_HEIGHT
  - NAV_MDD_ALIGN: "Center" | "Left" | "Right"

## Version History
- v2.6.1:
  - Default Category sheet renamed to "Catagory" (intentional spelling); code accepts both "Catagory" and "Category".
  - Dashboard: added charts Cash vs NAV and Portfolio_Catagory (replaces legacy Portfolio_Group).
  - Portfolio_Catagory prefers computing from Holdings + Catagory mapping; falls back to snapshot columns.
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
