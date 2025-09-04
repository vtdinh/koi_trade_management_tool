# koi_code_5 – Trading Portfolio Workbook Spec (v2.3.0)
**Generated:** 2025-09-01

## 📌 Overview
Excel/VBA workbook to aggregate crypto orders into positions, P&L, and dashboards.  
Core macros:
- **Update_All_PositionColumns** – rebuilds Position sheet, computes cash/coin/NAV and aggregates.  
- **Update_MarketPrice_ByCutoff_OpenOnly_Simple** – updates Market Price for Open rows only (optional helper).  
- **Take_Daily_Snapshot** – logs one row per day into *Daily_Snapshot* (Date, Cash, Coin, NAV, Deposit, Withdraw, Profit, Holdings).

## 📑 Sheets & Key Cells
### Position (`SHEET_PORTFOLIO`)
- **B3**: Cutoff datetime in UTC+7. If only date provided → treated as end-of-day 23:59:59 (UTC+7).  
- **Dashboard cells (configurable via constants):**
  - `CELL_CASH` = B5 (Cash)  
  - `CELL_COIN` = B6 (Coin market value of open holdings)  
  - `CELL_NAV`  = B7 (NAV = Cash + Coin)  
  - `CELL_SUM_DEPOSIT`  = D5 (Total deposit to cutoff)  
  - `CELL_SUM_WITHDRAW` = D6 (Total withdraw to cutoff)  
  - `CELL_TOTAL_PNL`    = D7 (Total profit to cutoff)

### Order_History (`SHEET_ORDERS`)
- Default header row = 2 (auto-detection supported).  
- Timestamps are **UTC-4** → all workbook logic uses **UTC+7** (+11h).  
- Supported columns (case/spacing tolerant):  
  - **Required:** `Date | Type | Coin | Qty`  
  - **Optional:** `Price | Fee | Exchange | Total`  
- If *Total* present → used for cash legs. Else:  
  ```vba
  Buy cash  = Qty*Price + Fee
  Sell cash = Qty*Price – Fee
  ```

### Daily_Snapshot
- Structure (A:H): `Date | Cash | Coin | NAV | Total deposit | Total withdraw | Total profit | Holdings`.  
- One row per date; UPSERT by Date; sorted ascending.

## ⏱ Time & Cutoff Rules
- Order_History timestamps = UTC-4; converted to UTC+7 via +11h.  
- Cutoff read from `Position!B3` (UTC+7). If Date-only → end-of-day 23:59:59.  
- **Pricing:**  
  - If cutoff < today: fetch Binance D1 close.  
  - If cutoff = today: fetch realtime ticker.  
  - Fallback: SYMBOLUSDT → SYMBOLUSDC → SYMBOLBUSD.  
  - Stablecoins (USDT/USDC/BUSD/FDUSD/TUSD) = 1.

## ⚙️ Position Building (`Update_All_PositionColumns`)
1. Map headers, clear old output.  
2. Iterate orders ≤ cutoff (UTC+7), maintain per-coin states:  
   - **BUY**: extend session, `Cost += Qty*Price + Fee`.  
   - **SELL**: extend session, `SellProceeds += Qty*Price – Fee`. If qty=0 → session Closed.  
   - **DEPOSIT/WITHDRAW**: affect only cash aggregates.  
3. Flush open sessions, compute AvailableQty.  
4. Pre-fetch market prices for Open coins.  
5. Write sessions to Position table: Open/Closed, Qty, Cost, Proceeds, Avg, Profit, %PnL, Exchange, Color-coded (green/red).  
6. Formats:  
   - Dates `yyyy-mm-dd`; %PnL `"0.00%"`; money `#,##0`; price `#,##0.00`.  
   - AutoFit columns, clear trailing rows.  

## 📊 Dashboard Metrics
- **Cash** = (ΣDeposit + ΣSell) – (ΣBuy + ΣWithdraw)  
- **Coin** = Σ AvailableQty_open × MarketPrice  
- **NAV**  = Cash + Coin  
- **Total deposit** = ΣDeposit  
- **Total withdraw** = ΣWithdraw  
- **Total profit**   = NAV – (Total deposit – Total withdraw)

## 📦 Holdings Value
- Built from Position table (rows `Open`).  
- `Value = Available Balance` or `Available Qty × Market Price`.  
- Aggregated per coin → `"BTC: 120,000; ETH: 3,612; …"`.

## 📅 Daily Snapshot (`Take_Daily_Snapshot`)
- Reads dashboard + holdings list.  
- Upsert by Date, never delete rows.  
- Sort ascending by Date.  
- Formats: `yyyy-mm-dd`, numeric `#,##0`, Holdings plain text.

## 🔤 Symbols & Mapping
- Mapping: `COIN → COINUSDT` (unless ends with USDT/USDC/BUSD).  
- Stablecoins recognized: USDT, USDC, BUSD, FDUSD, TUSD.

## 🛡 Error Handling & Safety
- Graceful messages for missing sheets/headers/data.  
- `CLEAR_MARKET_PRICE=True`: clear Market Price unless overwritten.  
- Robust LastRow detection.

## 🕒 Version History
- **v2.3.0:** Added Deposit/Withdraw/Total P&L aggregates, expanded Daily_Snapshot, clarified pricing/timezone.  
- **v2.2.x:** Position builder, cutoff normalization, Binance pricing, %PnL coloring, header auto-detection, stablecoin handling.
