Attribute VB_Name = "mod_config"
Option Explicit

' ===================== GLOBAL CONFIG =====================
' Centralized configuration and helper lists for all modules
'
' Sheets
Public Const SHEET_PORTFOLIO As String = "Position"
Public Const SHEET_ORDERS    As String = "Order_History"
Public Const SHEET_SNAPSHOT  As String = "Daily_Snapshot"
Public Const SHEET_CATEGORY  As String = "Catagory"   ' mapping Coin -> Group

' Charts
Public Const CHART_PORTFOLIO1 As String = "Portfolio1"
Public Const CHART_PORTFOLIO2 As String = "Portfolio2"

' Key cells on Position sheet
Public Const CELL_CUTOFF       As String = "B3"  ' cutoff datetime (UTC+7)
Public Const CELL_CASH         As String = "B5"
Public Const CELL_COIN         As String = "B6"
Public Const CELL_NAV          As String = "B7"
Public Const CELL_SUM_DEPOSIT  As String = "B8"
Public Const CELL_SUM_WITHDRAW As String = "B9"
Public Const CELL_TOTAL_PNL    As String = "B10"

' Orders table defaults
Public Const ORDERS_HEADER_ROW_DEFAULT As Long = 2
Public Const ORDERS_TZ_OFFSET_HOURS    As Long = 11   ' UTC-4 -> UTC+7 = +11h

' Behavior flags
Public Const CLEAR_MARKET_PRICE As Boolean = True

' Number formats
Public Const DATE_FMT  As String = "yyyy-mm-dd"
Public Const MONEY_FMT As String = "#,##0"
Public Const PRICE_FMT As String = "#,##0.00"
Public Const PCT_FMT   As String = "0.00%"

' Rounding defaults
Public Const ROUND_QTY_DECIMALS   As Long = 3
Public Const ROUND_MONEY_DECIMALS As Long = 0
Public Const ROUND_PRICE_DECIMALS As Long = 2

' Tolerances
Public Const EPS_ZERO As Double = 0.0000000001
Public Const EPS_CLOSE As Double = 0.0001

' Snapshot formats (if needed by snapshot module)
Public Const SNAPSHOT_DATE_FMT   As String = "yyyy-mm-dd"
Public Const SNAPSHOT_NUMBER_FMT As String = "#,##0"

' Colors (as helpers returning Long)
Public Function COLOR_PNL_POSITIVE() As Long
    ' #48A428
    COLOR_PNL_POSITIVE = RGB(72, 164, 40)
End Function

Public Function COLOR_PNL_NEGATIVE() As Long
    ' Red
    COLOR_PNL_NEGATIVE = RGB(192, 0, 0)
End Function

' Stablecoins treated as unit price = 1
Public Function Stablecoins() As Variant
    Stablecoins = Array("USDT", "USDC", "BUSD", "FDUSD", "TUSD")
End Function

' Symbol suffix fallbacks for pricing (try in order)
Public Function SymbolSuffixFallbacks() As Variant
    SymbolSuffixFallbacks = Array("USDT", "USDC", "BUSD")
End Function

' Flexible header aliases for Order_History
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

' Normalize header helper
Public Function NormalizeHeader(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace$(t, " ", "")
    t = Replace$(t, "_", "")
    t = Replace$(t, "-", "")
    NormalizeHeader = t
End Function
