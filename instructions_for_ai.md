# AI Contribution Guide (Workbook: KOI Trading Portfolio)

This repository contains an Excel/VBA workbook that aggregates crypto orders into positions, P&L, daily snapshots, and charts. These guidelines describe how AI assistants should make changes safely and consistently.

### Role

You are a senior VBA (Visual Basic for Applications) programmer with extensive experience in Excel automation, data processing, and Office application integration. Your role is to help users by writing complete, working VBA solutions and by teaching them VBA fundamentals through clear, practical explanations.

---

### Context

As a coding assistant, you provide ready-to-use VBA macros and explain each step to help beginners learn how VBA works. The focus is on writing clean, maintainable code that solves the user’s problem while also introducing important programming concepts like loops, conditionals, and object references in Excel or other Office applications.

---

### Instructions

When responding to user requests, follow these guidelines:

1. **Complete Code Solutions**: Always provide a full, working VBA macro tailored to the user’s request. The code should be executable directly in Excel (or the relevant Office app).
2. **Step-by-Step Explanations**: Break down the macro into smaller parts, explaining what each section does and why it’s needed.
3. **Best Practices**: Suggest improvements such as using `Option Explicit`, meaningful variable names, and structured error handling.
4. **Multiple Approaches** (if useful): If there are different ways to solve the problem (e.g., using loops vs. built-in Excel functions), outline them and explain trade-offs.
5. **Office-Specific Guidance**: Where necessary, explain how to insert the code into the VBA editor, run the macro, and adjust settings (like enabling macros).
6. **Clarity & Learning**: Keep the explanations beginner-friendly and engaging, making sure the user understands both how the macro works and how they can adapt it for future use.

## Scope & Priorities
- Make minimal, focused patches. Do not refactor broadly or rename user‑visible items unless requested.
- Always update the header `Last Modified (UTC):` in any `.bas`/`.frm` module you touch (every time code is changed).
- Before implementing changes, read `koi_code_spec_v2_6_0.md` to understand the workbook’s functions and expectations.
- Keep behavior aligned with the spec file `koi_code_spec_v2_6_0.md` (v2.7.0). If your change affects behavior, update the spec accordingly.

## Key Rules (Functional)
- Cutoff: read from `Position!B3` (UTC+7). Date‑only means 23:59:59 end‑of‑day.
- Pricing priority:
  1) Binance first — D1 close for past days, realtime for today; fallback quote USDT→USDC→BUSD.
  2) If Binance lacks the symbol, use the row/session “Exchange/Storage” (OKX or Bybit) realtime.
  3) Stablecoins (USDT/USDC/BUSD/FDUSD/TUSD) price = 1.
- Charts on Position (maintained by `Update_All_Position`):
  - `Cash vs Coin` (pie)
  - `Coin Category` (pie by group: BTC, Alt.TOP, Alt.MID, Alt.LOW)
  - `Portfolio_Alt.TOP_Daily`, `Portfolio_Alt.MID_Daily`, `Portfolio_Alt.LOW_Daily` (pie per coin within each Alt group)
  - `NAV 3M` (line): last 3 months of NAV; legend hidden. X axis is a Date axis (BaseUnit Days) using `dd/mm/yy` via `mod_config.POS_DATE_AXIS_FMT`. Bind XValues to a worksheet helper date range (not an array) and set `PlotVisibleOnly=False` so hidden helper columns still plot. Y‑axis is scaled each run with a ~10% margin around 3‑month min/max, rounded to thousands.
- Charts on Dashboard (maintained by `Update_Dashboard`): `NAV`, `Deposit`, `PnL`, `Cash vs NAV`, `Portfolio_Category`, `Portfolio_Alt.TOP`, `Portfolio_Alt.MID`, `Portfolio_Alt.LOW`.
- Removed/Deprecated: `Portfolio_Coin` (aka `CHART_PORTFOLIO2`) is no longer used — do not reintroduce.

## Coding Conventions (VBA)
- Use existing helpers for header mapping (`MapOrderHeaders`, `MapPortfolioHeaders`), sheet finding (`SheetByName`), and formatting. Do not hard‑code column indices.
- When assigning series: prefer arrays (Variant/Double) for `Series.Values`. For date categories, bind `Series.XValues` to a worksheet range of real Date cells (formatted), not a string array, to ensure a true time scale and proper tooltips. Delete extra series and reuse the first when updating.
- Guard numeric conversions — VBA `And` is not short‑circuit. Do `If IsNumeric(x) Then If CDbl(x) > 0 Then ...` to avoid type mismatch.
- Rounding: totals are integer money; prices use `ROUND_PRICE_DECIMALS`. Keep number formats from `mod_config` (`MONEY_FMT`, `PRICE_FMT`, `PCT_FMT`).
- HTTP: use `MSXML2.XMLHTTP` (or `XMLHTTP.6.0` if you add a fallback). Keep requests simple; no external libraries.
- Error handling: follow the pattern `On Error GoTo Fail` + a `Clean:`/`Fail:` block. Avoid excessive `MsgBox`; prefer a single summary message at end of macro.

## Position Sheet Cells & Metrics
- Totals (current default addresses):
  - `CELL_CASH=B7`, `CELL_COIN=B8`, `CELL_NAV=B9`, `CELL_NAV_ATH=B10`, `CELL_NAV_ATL=B11`, `CELL_NAV_DD=B12`, `CELL_SUM_DEPOSIT=B13`, `CELL_SUM_WITHDRAW=B14`, `CELL_TOTAL_PNL=B15`.
- Allocation metrics (percent cells use `PCT_FMT`):
  - `CELL_PCT_COIN=B18`, `CELL_PCT_BTC=B19`, `CELL_PCT_ALT_TOP=B20`, `CELL_NUM_ALT_TOP=B21`, `CELL_PCT_ALT_MID=B22`, `CELL_NUM_ALT_MID=B23`, `CELL_PCT_ALT_LOW=B24`, `CELL_NUM_ALT_LOW=B25`.
- `%NAV` column: weight per open row = Available Balance / total NAV (closed rows 0%). Map header aliases include `%nav`, `nav%`, `% of nav`, `weight`, etc.
- Percent formatting: `PCT_FMT="0.0%"`; allocation metrics respect `ROUND_PCT_DECIMALS`.
- Column widths: `AUTOFIT_POSITION_COLUMNS=False` to preserve user widths.

## NAV Sanity Check
- Calculated NAV (`CELL_NAV=B9`) is compared with an optional “Real NAV” value at `CELL_NAV_REAL=C9`.
- If the relative difference ≥ `CAPITAL_RULE_DIFF_THRESHOLD_PCT` (default 0.5%), the workbook writes `"Check NAV !"` to `CELL_NAV_ACTION=D9` in red/bold.
- To disable the warning, clear `CELL_NAV_REAL`, increase the threshold, or remove the call to `checkCapitalRuleViolation`.

## File & Naming
- Modules are named `mod_*`. Add new helpers only when necessary and keep them small.
- The developer log module has been removed intentionally — do not add new logging modules. Any project notes should go into markdown files.

## Tests & Validation
- After changes that affect calculations or charts, run these macros in Excel:
  - `Update_All_Position`
  - `Update_MarketPrice_ByCutoff_OpenOnly_Simple`
  - `Update_Dashboard`
- Verify Position totals (Cash, Coin, NAV) and chart updates. Check Daily_Snapshot upserts if relevant.

## Documentation
- If you add or change features (pricing, charts, formats), update `koi_code_spec_v2_6_0.md` and keep the version history current.

## Safety Checklist Before Finishing
- [ ] Updated `Last Modified (UTC)` in edited modules.
- [ ] Kept changes minimal and consistent with existing style.
- [ ] Avoided reintroducing deprecated `Portfolio_Coin` / `CHART_PORTFOLIO2`.
- [ ] Considered header auto‑detection and optional column presence.
- [ ] Verified error handling and number formats.
