Attribute VB_Name = "mod_config"
Option Explicit

' ===================== DASHBOARD: NAV Drawdown Text =====================
' Positioning and style config for the Max Drawdown textbox on NAV chart
' NAV_MDD_ANCHOR options:
'   - "UnderTitle"   : centered just below the chart title (default)
'   - "PlotTopRight" : inside plot area, top-right
'   - "PlotTopLeft"  : inside plot area, top-left
Public Const NAV_MDD_ANCHOR   As String = "UnderTitle"
Public Const NAV_MDD_OFFSET_X As Single = 0       ' pixels relative to anchor
Public Const NAV_MDD_OFFSET_Y As Single = 0       ' pixels relative to anchor
Public Const NAV_MDD_WIDTH    As Single = 160
Public Const NAV_MDD_HEIGHT   As Single = 18
' NAV_MDD_ALIGN: "Center" | "Left" | "Right"
Public Const NAV_MDD_ALIGN    As String = "Center"

' ===================== GLOBAL CONFIG =====================
' Centralized configuration and helper lists for all modules
'
' Sheets
Public Const SHEET_PORTFOLIO As String = "Position"
Public Const SHEET_ORDERS    As String = "Order_History"
Public Const SHEET_SNAPSHOT  As String = "Daily_Snapshot"
Public Const SHEET_CATEGORY  As String = "Category"   ' mapping Coin -> Group
Public Const SHEET_DASHBOARD As String = "Dashboard"

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
 ' (Removed unused helper lists and NormalizeHeader to keep module minimal)

