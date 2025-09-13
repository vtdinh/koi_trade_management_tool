Attribute VB_Name = "mod_config"
Option Explicit
' Last Modified (UTC): 2025-09-12T03:38:30Z

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
Public Const SHEET_CATEGORY  As String = "Categoty"   ' mapping Coin -> Group (sheet header row=category, below=list of coins)
Public Const SHEET_DASHBOARD As String = "Dashboard"

' Charts
' Position sheet group pie (legacy names auto-renamed in code)
Public Const CHART_PORTFOLIO1 As String = "Coin Category"

' Key cells on Position sheet
Public Const CELL_CUTOFF       As String = "B3"   ' cutoff datetime (UTC+7)
Public Const CELL_CASH         As String = "B7"
Public Const CELL_COIN         As String = "B8"
Public Const CELL_NAV          As String = "B9"
Public Const CELL_NAV_ATH      As String = "B10"  ' All‑time high of NAV
Public Const CELL_NAV_ATL      As String = "B11"  ' All‑time low of NAV
Public Const CELL_NAV_DD       As String = "B12"  ' NAV drawdown text/value
Public Const CELL_SUM_DEPOSIT  As String = "B13"
Public Const CELL_SUM_WITHDRAW As String = "B14"
Public Const CELL_TOTAL_PNL    As String = "B15"

' Capital rule check (Position sheet)
Public Const CELL_NAV_REAL     As String = "C9"   ' Real NAV (user input)
Public Const CELL_NAV_ACTION   As String = "D9"   ' Warning/Action cell
Public Const CAPITAL_RULE_DIFF_THRESHOLD_PCT As Double = 0.005  ' 0.5%

' NAV drawdown behavior
Public Const NAV_DD_USE_TRUNCATED As Boolean = True            ' compute drawdown using truncated (0‑decimal) NAVs
Public Const NAV_DD_TOLERANCE_PCT As Double = 0.001            ' <=0.1% → treat as 0%

' NAV drawdown threshold & action (Position sheet)
Public Const CELL_NAV_DD_LIMIT   As String = "C12"  ' threshold value (e.g., 30% as 0.30)
Public Const CELL_NAV_DD_ACTION  As String = "D12"  ' action message cell

' Allocation and counts (today/cutoff)
Public Const CELL_PCT_COIN     As String = "B18"  ' %Coin = Coin/NAV
Public Const CELL_PCT_BTC      As String = "B19"  ' %BTC of total Coin
Public Const CELL_PCT_ALT_TOP  As String = "B20"  ' %Alt.TOP of total Coin
Public Const CELL_NUM_ALT_TOP  As String = "B21"  ' Number of Alt.TOP coins
Public Const CELL_PCT_ALT_MID  As String = "B22"  ' %Alt.MID of total Coin
Public Const CELL_NUM_ALT_MID  As String = "B23"  ' Number of Alt.MID coins
Public Const CELL_PCT_ALT_LOW  As String = "B24"  ' %Alt.LOW of total Coin
Public Const CELL_NUM_ALT_LOW  As String = "B25"  ' Number of Alt.LOW coins

' Orders table defaults
Public Const ORDERS_HEADER_ROW_DEFAULT As Long = 2
Public Const ORDERS_TZ_OFFSET_HOURS    As Long = 11   ' UTC-4 -> UTC+7 = +11h

' Behavior flags
Public Const CLEAR_MARKET_PRICE As Boolean = True
Public Const AUTOFIT_POSITION_COLUMNS As Boolean = False  ' keep existing widths on Position

' Number formats
Public Const DATE_FMT  As String = "dd-mmm-yy"   ' e.g., 12-Sep-25
Public Const MONEY_FMT As String = "#,##0"
Public Const PRICE_FMT As String = "#,##0.00"
Public Const PCT_FMT   As String = "0.0%"   ' e.g., 12.3%

' Rounding defaults
Public Const ROUND_QTY_DECIMALS   As Long = 3
Public Const ROUND_MONEY_DECIMALS As Long = 0
Public Const ROUND_PRICE_DECIMALS As Long = 2
Public Const ROUND_PCT_DECIMALS   As Long = 0   ' decimals to show for percentages (e.g., 1 -> 0.0%)

' Tolerances
Public Const EPS_ZERO As Double = 0.0000000001
Public Const EPS_CLOSE As Double = 0.0001

' Snapshot formats (if needed by snapshot module)
Public Const SNAPSHOT_DATE_FMT   As String = "yyyy-mm-dd"
Public Const POS_DATE_AXIS_FMT  As String = "d-m-yy"            ' Position chart axis (e.g., 1-9-25)
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

