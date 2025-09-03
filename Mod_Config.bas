Attribute VB_Name = "Mod_Config"
' ===========================================================
' Mod_Config.bas - v1.0.3
' - Centralized user configuration for the workbook
' - Keep constants first, then helper functions
' ===========================================================
Option Explicit

' ===== Sheet names =====
Public Const SHEET_PORTFOLIO As String = "Position"
Public Const SHEET_ORDERS    As String = "Order_History"
Public Const SHEET_SNAPSHOT  As String = "Daily_Snapshot"

' ===== Position sheet cells =====
Public Const CELL_CUTOFF        As String = "B3"  ' cutoff datetime (UTC+7)
Public Const CELL_CASH          As String = "B5"
Public Const CELL_COIN          As String = "B6"
Public Const CELL_NAV           As String = "B7"
Public Const CELL_SUM_DEPOSIT   As String = "D5"
Public Const CELL_SUM_WITHDRAW  As String = "D6"
Public Const CELL_TOTAL_PNL     As String = "D7"

' ===== Order_History defaults =====
Public Const ORDERS_HEADER_ROW_DEFAULT As Long = 2
' Source logs are UTC-4; workbook uses UTC+7 => +11h
Public Const ORDERS_TZ_OFFSET_HOURS As Long = 11

' ===== Cutoff rules =====
' If cutoff is a date only (no time), treat as end-of-day 23:59:59 (UTC+7)
Public Const TREAT_DATE_ONLY_AS_END_OF_DAY As Boolean = True

' ===== Pricing rules & mapping =====
' If cutoff is today (UTC+7) use realtime; otherwise D1 close
Public Const PRICE_USE_REALTIME_IF_TODAY As Boolean = True
Public Const STABLECOIN_UNIT_PRICE As Double = 1#

' Optional: provider info (reference only)
Public Const PRICE_PROVIDER_NAME As String = "Binance"
Public Const BINANCE_API_BASE    As String = "https://api.binance.com"

' ===== Formats =====
Public Const DATE_FMT  As String = "yyyy-mm-dd"
Public Const MONEY_FMT As String = "#,##0"
Public Const PRICE_FMT As String = "#,##0.00"
Public Const PCT_FMT   As String = "0.00%"

' Rounding digits (use WorksheetFunction.Round in core code if needed)
Public Const ROUND_QTY_DECIMALS   As Long = 3
Public Const ROUND_MONEY_DECIMALS As Long = 0
Public Const ROUND_PRICE_DECIMALS As Long = 2

' ===== Position rendering =====
Public Const AUTOFIT_WRITTEN_COLUMNS  As Boolean = True
Public Const CLEAR_MARKET_PRICE       As Boolean = True
Public Const TRAILING_CLEAR_UNTIL_ROW As Long = 50

' ===== Daily Snapshot =====
' Columns: Date | Cash | Coin | NAV | Total deposit | Total withdraw | Total profit | Holdings
Public Const SNAPSHOT_DATE_FMT   As String = "yyyy-mm-dd"
Public Const SNAPSHOT_NUMBER_FMT As String = "#,##0"

' ===== Behavior flags =====
Public Const ALLOW_FETCH_D1_CLOSE_FOR_PAST  As Boolean = True
Public Const ALLOW_FETCH_REALTIME_FOR_TODAY As Boolean = True

' ===== Color helpers (as functions returning Long) =====
Public Function COLOR_PNL_POSITIVE() As Long
    COLOR_PNL_POSITIVE = RGB(0, 176, 80)
End Function

Public Function COLOR_PNL_NEGATIVE() As Long
    COLOR_PNL_NEGATIVE = RGB(192, 0, 0)
End Function

' ===== Mapping & lists (Functions AFTER all Const) =====
Public Function SymbolSuffixFallbacks() As Variant
    ' Map "COIN" -> "COIN{QUOTE}" try in order:
    SymbolSuffixFallbacks = Array("USDT", "USDC", "BUSD")
End Function

Public Function Stablecoins() As Variant
    ' Stablecoins valued at 1
    Stablecoins = Array("USDT", "USDC", "BUSD", "FDUSD", "TUSD")
End Function

' ===== Flexible header aliases for Order_History =====
Public Function HDR_DATE_ALIASES() As Variant
    HDR_DATE_ALIASES = Array("date", "datetime", "time", "timestamp")
End Function

Public Function HDR_TYPE_ALIASES() As Variant
    HDR_TYPE_ALIASES = Array("type", "side", "action")
End Function

Public Function HDR_COIN_ALIASES() As Variant
    HDR_COIN_ALIASES = Array("coin", "symbol", "asset")
End Function

Public Function HDR_QTY_ALIASES() As Variant
    HDR_QTY_ALIASES = Array("qty", "quantity", "amount")
End Function

Public Function HDR_PRICE_ALIASES() As Variant
    HDR_PRICE_ALIASES = Array("price", "unit price", "avg price")
End Function

Public Function HDR_FEE_ALIASES() As Variant
    HDR_FEE_ALIASES = Array("fee", "fees", "commission")
End Function

Public Function HDR_EXCHANGE_ALIASES() As Variant
    HDR_EXCHANGE_ALIASES = Array("exchange", "venue", "broker", "storage")
End Function

Public Function HDR_TOTAL_ALIASES() As Variant
    HDR_TOTAL_ALIASES = Array("total", "cash", "gross", "notional")
End Function

' ===== Helper: normalize header =====
Public Function NormalizeHeader(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace$(t, " ", "")
    t = Replace$(t, "_", "")
    t = Replace$(t, "-", "")
    NormalizeHeader = t
End Function

' ===== Reference (for humans)
' Cash = (SumDeposit + SumSell) - (SumBuy + SumWithdraw)
' Coin = Sum of (AvailableQty_open * MarketPrice)
' NAV  = Cash + Coin
' Total profit = NAV - (Total deposit - Total withdraw)

