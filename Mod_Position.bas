Attribute VB_Name = "Mod_Position"
' ===========================================================
' Mod_Position.bas - v1.0.3
' - Build Position from Order_History up to cutoff (UTC+7)
' - Write Dashboard (Cash/Coin/NAV/Deposit/Withdraw/Total PnL)
' - Update Market Price for Open rows only
' - Compatible with Mod_Config.bas (v1.0.x)
' ===========================================================
Option Explicit

' ==== TYPES ====
Public Type OrderHeaderMap
    rowHeader As Long
    cDate As Long
    cType As Long
    cCoin As Long
    cQty As Long
    cPrice As Long
    cFee As Long
    cXchg As Long
    cTotal As Long
End Type

Public Type SessionRow
    coin As String
    Position As String
    BuyQty As Double
    SellQty As Double
    AvailQty As Double
    Cost As Double
    AvgCost As Double
    SellProceeds As Double
    AvgSell As Double
    MktPrice As Double
    AvailBalance As Double
    Profit As Double
    PctPnL As Double
    Storage As String
End Type

Public Type PortfolioAggregate
    Sessions() As SessionRow
    SumDeposit As Double
    SumWithdraw As Double
    SumSell As Double
    SumBuy As Double
    cash As Double
    coin As Double
    NAV As Double
    TotalPnL As Double
End Type

' ========= Public API =========
Public Sub Update_All_PositionColumns()
    Dim wsP As Worksheet, wsO As Worksheet
    Set wsP = SheetByName(Mod_Config.SHEET_PORTFOLIO)
    Set wsO = SheetByName(Mod_Config.SHEET_ORDERS)
    If wsP Is Nothing Or wsO Is Nothing Then
        MsgBox "Missing sheet '" & Mod_Config.SHEET_PORTFOLIO & "' or '" & Mod_Config.SHEET_ORDERS & "'.", vbExclamation
        Exit Sub
    End If

    Dim cutoff As Date
    If Not TryReadCutoff(wsP, cutoff) Then
        MsgBox "Cannot read cutoff at " & Mod_Config.SHEET_PORTFOLIO & " !" & Mod_Config.CELL_CUTOFF, vbExclamation
        Exit Sub
    End If

    Dim hdr As OrderHeaderMap
    If Not MapOrderHeaders(wsO, hdr) Then
        MsgBox "Required columns not found on '" & Mod_Config.SHEET_ORDERS & "'.", vbExclamation
        Exit Sub
    End If

    Dim agg As PortfolioAggregate
    BuildPortfolio wsO, hdr, cutoff, agg

    RenderPosition wsP, agg
    WriteDashboard wsP, agg

    If Mod_Config.AUTOFIT_WRITTEN_COLUMNS Then
        wsP.UsedRange.Columns.AutoFit
    End If
End Sub

' Optional helper: update Market Price for open rows only
Public Sub Update_MarketPrice_ByCutoff_OpenOnly_Simple()
    Dim wsP As Worksheet
    Set wsP = SheetByName(Mod_Config.SHEET_PORTFOLIO)
    If wsP Is Nothing Then Exit Sub

    Dim cutoff As Date
    If Not TryReadCutoff(wsP, cutoff) Then Exit Sub

    Dim lastRow As Long
    lastRow = LastRowFast(wsP, 1)
    If lastRow < 5 Then Exit Sub

    Dim colCoin As Long, colPos As Long, colMkt As Long, colAvail As Long
    FindPositionTableColumns wsP, colCoin, colPos, colMkt, colAvail
    If colCoin = 0 Or colPos = 0 Or colMkt = 0 Then Exit Sub

    Dim r As Long
    For r = 5 To lastRow
        If UCase$(Nz(wsP.Cells(r, colPos).Value)) = "OPEN" Then
            Dim coin As String: coin = CStr(Nz(wsP.Cells(r, colCoin).Value))
            If Len(coin) > 0 And Not IsStable(coin) Then
                Dim px As Double: px = GetMarketPrice_ByRule(coin, cutoff)
                If px > 0 Then wsP.Cells(r, colMkt).Value = Round(px, Mod_Config.ROUND_PRICE_DECIMALS)
            ElseIf IsStable(coin) Then
                wsP.Cells(r, colMkt).Value = Mod_Config.STABLECOIN_UNIT_PRICE
            End If
        End If
    Next
End Sub

' ========= Core Build (no Variant<->UDT coercion) =========
Private Sub BuildPortfolio(wsO As Worksheet, hdr As OrderHeaderMap, cutoff As Date, agg As PortfolioAggregate)
    Dim lastRow As Long
    lastRow = LastRowFast(wsO, IIf(hdr.cDate > 0, hdr.cDate, 1))

    ' Per-coin states
    Dim runQty As Object:       Set runQty = CreateObject("Scripting.Dictionary")
    Dim latestXchg As Object:   Set latestXchg = CreateObject("Scripting.Dictionary")
    Dim dBuyQty As Object:      Set dBuyQty = CreateObject("Scripting.Dictionary")
    Dim dSellQty As Object:     Set dSellQty = CreateObject("Scripting.Dictionary")
    Dim dCost As Object:        Set dCost = CreateObject("Scripting.Dictionary")
    Dim dSellCash As Object:    Set dSellCash = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = hdr.rowHeader + 1 To lastRow
        Dim vDate As Variant: vDate = wsO.Cells(i, hdr.cDate).Value
        If IsEmpty(vDate) Then GoTo ContinueLoop

        ' UTC-4 -> UTC+7 (+11h)
        Dim tZ7 As Date
        On Error Resume Next
        tZ7 = CDate(vDate) + (Mod_Config.ORDERS_TZ_OFFSET_HOURS / 24#)
        On Error GoTo 0
        If tZ7 = 0 Or tZ7 > cutoff Then GoTo ContinueLoop

        Dim t As String:   t = UCase$(Trim$(CStr(Nz(wsO.Cells(i, hdr.cType).Value))))
        Dim coin As String: coin = UCase$(Trim$(CStr(Nz(wsO.Cells(i, hdr.cCoin).Value))))
        Dim qty As Double:  qty = CDbl(val(Nz(wsO.Cells(i, hdr.cQty).Value)))
        Dim price As Double: price = CDbl(val(Nz(IfNZ(hdr.cPrice, wsO, i))))
        Dim fee As Double:   fee = CDbl(val(Nz(IfNZ(hdr.cFee, wsO, i))))
        Dim exch As String:  exch = CStr(Nz(IfNZText(hdr.cXchg, wsO, i)))
        Dim total As Variant: total = IfNZ(hdr.cTotal, wsO, i)

        If Len(exch) > 0 Then latestXchg(coin) = exch

        Select Case t
            Case "BUY"
                Dim buyCash As Double
                If Not IsMissingVal(total) Then buyCash = CDbl(total) Else buyCash = qty * price + fee
                agg.SumBuy = agg.SumBuy + buyCash
                runQty(coin) = Nz(runQty(coin)) + qty
                dBuyQty(coin) = Nz(dBuyQty(coin)) + qty
                dCost(coin) = Nz(dCost(coin)) + buyCash

            Case "SELL"
                Dim sellCash As Double
                If Not IsMissingVal(total) Then sellCash = CDbl(total) Else sellCash = qty * price - fee
                agg.SumSell = agg.SumSell + sellCash
                runQty(coin) = Nz(runQty(coin)) - qty
                dSellQty(coin) = Nz(dSellQty(coin)) + qty
                dSellCash(coin) = Nz(dSellCash(coin)) + sellCash

            Case "DEPOSIT"
                agg.SumDeposit = agg.SumDeposit + CDbl(val(Nz(IIf(IsMissingVal(total), qty * price, total))))

            Case "WITHDRAW"
                agg.SumWithdraw = agg.SumWithdraw + CDbl(val(Nz(IIf(IsMissingVal(total), qty * price, total))))
        End Select
ContinueLoop:
    Next i

    ' Build SessionRow array from dictionaries
    Dim keys As Object: Set keys = CreateObject("Scripting.Dictionary")
    CopyKeys runQty, keys: CopyKeys dBuyQty, keys: CopyKeys dSellQty, keys: CopyKeys dCost, keys

    Dim arr() As SessionRow
    Dim idx As Long: idx = 0
    Dim k As Variant
    For Each k In keys.keys
        Dim s As SessionRow
        s.coin = CStr(k)
        s.BuyQty = CDbl(Nz(dBuyQty(k)))
        s.SellQty = CDbl(Nz(dSellQty(k)))
        s.Cost = CDbl(Nz(dCost(k)))
        s.SellProceeds = CDbl(Nz(dSellCash(k)))
        If s.BuyQty > 0 Then s.AvgCost = SafeDiv(s.Cost, s.BuyQty)
        If s.SellQty > 0 Then s.AvgSell = SafeDiv(s.SellProceeds, s.SellQty)

        s.AvailQty = Round(Nz(runQty(s.coin)), Mod_Config.ROUND_QTY_DECIMALS)
        s.Position = IIf(s.AvailQty > 0, "Open", "Closed")
        If latestXchg.Exists(s.coin) Then s.Storage = CStr(latestXchg(s.coin))

        PushSession arr, idx, s
    Next k

    ' Prefetch price and compute balances/PNL
    Dim j As Long
    For j = LBound(arr) To IIf((Not (Not arr)), UBound(arr), -1)
        If j = -1 Then Exit For
        If arr(j).Position = "Open" Then
            Dim px As Double
            If IsStable(arr(j).coin) Then
                px = Mod_Config.STABLECOIN_UNIT_PRICE
            Else
                px = GetMarketPrice_ByRule(arr(j).coin, cutoff)
            End If
            arr(j).MktPrice = Round(px, Mod_Config.ROUND_PRICE_DECIMALS)
            If arr(j).AvailQty > 0 And px > 0 Then _
                arr(j).AvailBalance = Round(arr(j).AvailQty * px, Mod_Config.ROUND_MONEY_DECIMALS)
            arr(j).Profit = Round((arr(j).SellProceeds + (arr(j).AvailBalance)) - arr(j).Cost, Mod_Config.ROUND_MONEY_DECIMALS)
            If arr(j).Cost <> 0 Then arr(j).PctPnL = SafeDiv(arr(j).Profit, arr(j).Cost)
        Else
            arr(j).Profit = Round(arr(j).SellProceeds - arr(j).Cost, Mod_Config.ROUND_MONEY_DECIMALS)
            If arr(j).Cost <> 0 Then arr(j).PctPnL = SafeDiv(arr(j).Profit, arr(j).Cost)
        End If
    Next j

    ' Dashboard
    Dim cash As Double
    cash = (agg.SumDeposit + agg.SumSell) - (agg.SumBuy + agg.SumWithdraw)
    Dim coinVal As Double
    If (Not (Not arr)) Then
        For j = LBound(arr) To UBound(arr)
            If arr(j).Position = "Open" Then coinVal = coinVal + arr(j).AvailBalance
        Next j
    End If

    agg.cash = Round(cash, Mod_Config.ROUND_MONEY_DECIMALS)
    agg.coin = Round(coinVal, Mod_Config.ROUND_MONEY_DECIMALS)
    agg.NAV = agg.cash + agg.coin
    agg.TotalPnL = agg.NAV - (agg.SumDeposit - agg.SumWithdraw)

    AggAssignSessions agg, arr
End Sub

' ========= Render Position =========
Private Sub RenderPosition(wsP As Worksheet, agg As PortfolioAggregate)
    Dim headers As Variant
    headers = Array( _
        "Coin", "Position", "Buy Qty", "Sell Qty", "Available Qty", _
        "Cost", "Avg cost", "Sell proceeds", "Avg sell price", _
        "Market Price", "Available Balance", "Profit", "%PnL", "Storage" _
    )

    Dim r0 As Long: r0 = 4
    Dim c0 As Long: c0 = 1
    Dim i As Long
    For i = 0 To UBound(headers)
        wsP.Cells(r0, c0 + i).Value = headers(i)
    Next i

    Dim last As Long
    last = LastRowFast(wsP, 1)
    If last > r0 Then
        wsP.Rows(r0 + 1 & ":" & Application.Max(last, Mod_Config.TRAILING_CLEAR_UNTIL_ROW)).ClearContents
    End If

    Dim r As Long: r = r0 + 1
    Dim s As SessionRow
    Dim idx As Long
    If (Not (Not agg.Sessions)) Then
        For idx = LBound(agg.Sessions) To UBound(agg.Sessions)
            s = agg.Sessions(idx)
            wsP.Cells(r, 1).Value = s.coin
            wsP.Cells(r, 2).Value = s.Position
            wsP.Cells(r, 3).Value = Round(s.BuyQty, Mod_Config.ROUND_QTY_DECIMALS)
            wsP.Cells(r, 4).Value = Round(s.SellQty, Mod_Config.ROUND_QTY_DECIMALS)
            wsP.Cells(r, 5).Value = Round(s.AvailQty, Mod_Config.ROUND_QTY_DECIMALS)
            wsP.Cells(r, 6).Value = Round(s.Cost, Mod_Config.ROUND_MONEY_DECIMALS)
            wsP.Cells(r, 7).Value = Round(s.AvgCost, Mod_Config.ROUND_PRICE_DECIMALS)
            wsP.Cells(r, 8).Value = Round(s.SellProceeds, Mod_Config.ROUND_MONEY_DECIMALS)
            wsP.Cells(r, 9).Value = Round(s.AvgSell, Mod_Config.ROUND_PRICE_DECIMALS)
            wsP.Cells(r, 10).Value = Round(s.MktPrice, Mod_Config.ROUND_PRICE_DECIMALS)
            wsP.Cells(r, 11).Value = Round(s.AvailBalance, Mod_Config.ROUND_MONEY_DECIMALS)
            wsP.Cells(r, 12).Value = Round(s.Profit, Mod_Config.ROUND_MONEY_DECIMALS)
            wsP.Cells(r, 13).Value = s.PctPnL
            wsP.Cells(r, 14).Value = s.Storage

            wsP.Cells(r, 13).NumberFormat = Mod_Config.PCT_FMT
            If s.Profit > 0 Then
                wsP.Cells(r, 12).Font.Color = Mod_Config.COLOR_PNL_POSITIVE
                wsP.Cells(r, 13).Font.Color = Mod_Config.COLOR_PNL_POSITIVE
            ElseIf s.Profit < 0 Then
                wsP.Cells(r, 12).Font.Color = Mod_Config.COLOR_PNL_NEGATIVE
                wsP.Cells(r, 13).Font.Color = Mod_Config.COLOR_PNL_NEGATIVE
            End If
            r = r + 1
        Next idx
    End If

    wsP.Range(wsP.Cells(r0 + 1, 3), wsP.Cells(r - 1, 5)).NumberFormat = "0." & String(Mod_Config.ROUND_QTY_DECIMALS, "0")
    wsP.Range(wsP.Cells(r0 + 1, 6), wsP.Cells(r - 1, 6)).NumberFormat = Mod_Config.MONEY_FMT
    wsP.Range(wsP.Cells(r0 + 1, 7), wsP.Cells(r - 1, 7)).NumberFormat = Mod_Config.PRICE_FMT
    wsP.Range(wsP.Cells(r0 + 1, 8), wsP.Cells(r - 1, 8)).NumberFormat = Mod_Config.MONEY_FMT
    wsP.Range(wsP.Cells(r0 + 1, 9), wsP.Cells(r - 1, 9)).NumberFormat = Mod_Config.PRICE_FMT
    wsP.Range(wsP.Cells(r0 + 1, 10), wsP.Cells(r - 1, 10)).NumberFormat = Mod_Config.PRICE_FMT
    wsP.Range(wsP.Cells(r0 + 1, 11), wsP.Cells(r - 1, 11)).NumberFormat = Mod_Config.MONEY_FMT
    wsP.Range(wsP.Cells(r0 + 1, 12), wsP.Cells(r - 1, 12)).NumberFormat = Mod_Config.MONEY_FMT
End Sub

Private Sub WriteDashboard(wsP As Worksheet, agg As PortfolioAggregate)
    wsP.Range(Mod_Config.CELL_CASH).Value = agg.cash
    wsP.Range(Mod_Config.CELL_COIN).Value = agg.coin
    wsP.Range(Mod_Config.CELL_NAV).Value = agg.NAV
    wsP.Range(Mod_Config.CELL_SUM_DEPOSIT).Value = agg.SumDeposit
    wsP.Range(Mod_Config.CELL_SUM_WITHDRAW).Value = agg.SumWithdraw
    wsP.Range(Mod_Config.CELL_TOTAL_PNL).Value = agg.TotalPnL

    wsP.Range(Mod_Config.CELL_CASH).NumberFormat = Mod_Config.MONEY_FMT
    wsP.Range(Mod_Config.CELL_COIN).NumberFormat = Mod_Config.MONEY_FMT
    wsP.Range(Mod_Config.CELL_NAV).NumberFormat = Mod_Config.MONEY_FMT
    wsP.Range(Mod_Config.CELL_SUM_DEPOSIT).NumberFormat = Mod_Config.MONEY_FMT
    wsP.Range(Mod_Config.CELL_SUM_WITHDRAW).NumberFormat = Mod_Config.MONEY_FMT
    wsP.Range(Mod_Config.CELL_TOTAL_PNL).NumberFormat = Mod_Config.MONEY_FMT
End Sub

' ========= Price Rules =========
Private Function GetMarketPrice_ByRule(ByVal coin As String, ByVal cutoff As Date) As Double
    If IsTodayUTC7(cutoff) And Mod_Config.ALLOW_FETCH_REALTIME_FOR_TODAY Then
        ' TODO: plug in realtime provider, e.g. PriceRealtime(MapSymbol(coin))
        GetMarketPrice_ByRule = 0#
    ElseIf Mod_Config.ALLOW_FETCH_D1_CLOSE_FOR_PAST Then
        ' TODO: plug in D1 close provider, e.g. PriceD1Close(MapSymbol(coin), DateValue(cutoff))
        GetMarketPrice_ByRule = 0#
    Else
        GetMarketPrice_ByRule = 0#
    End If
End Function

Private Function MapSymbol(ByVal coin As String) As String
    Dim up As String: up = UCase$(Trim$(coin))
    If EndsWithAny(up, Mod_Config.SymbolSuffixFallbacks) Then
        MapSymbol = up
    Else
        MapSymbol = up & "USDT"
    End If
End Function

Private Function IsStable(ByVal coin As String) As Boolean
    Dim s As Variant
    For Each s In Mod_Config.Stablecoins
        If UCase$(coin) = UCase$(CStr(s)) Then
            IsStable = True
            Exit Function
        End If
    Next s
End Function

' ========= Header Mapping / Utilities =========
Private Function MapOrderHeaders(ws As Worksheet, ByRef hdr As OrderHeaderMap) As Boolean
    Dim rH As Long: rH = Mod_Config.ORDERS_HEADER_ROW_DEFAULT
    hdr.rowHeader = rH

    Dim lastCol As Long
    lastCol = ws.Cells(rH, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        Dim head As String: head = CStr(Nz(ws.Cells(rH, c).Value))
        Dim n As String: n = NormalizeHeader(head)
        If InAliases(n, Mod_Config.HDR_DATE_ALIASES) Then hdr.cDate = c
        If InAliases(n, Mod_Config.HDR_TYPE_ALIASES) Then hdr.cType = c
        If InAliases(n, Mod_Config.HDR_COIN_ALIASES) Then hdr.cCoin = c
        If InAliases(n, Mod_Config.HDR_QTY_ALIASES) Then hdr.cQty = c
        If InAliases(n, Mod_Config.HDR_PRICE_ALIASES) Then hdr.cPrice = c
        If InAliases(n, Mod_Config.HDR_FEE_ALIASES) Then hdr.cFee = c
        If InAliases(n, Mod_Config.HDR_EXCHANGE_ALIASES) Then hdr.cXchg = c
        If InAliases(n, Mod_Config.HDR_TOTAL_ALIASES) Then hdr.cTotal = c
    Next c

    MapOrderHeaders = (hdr.cDate > 0 And hdr.cType > 0 And hdr.cCoin > 0 And hdr.cQty > 0)
End Function

Private Function TryReadCutoff(ws As Worksheet, ByRef cutoff As Date) As Boolean
    Dim v As Variant: v = ws.Range(Mod_Config.CELL_CUTOFF).Value
    If IsDate(v) Then
        cutoff = CDate(v)
        If Mod_Config.TREAT_DATE_ONLY_AS_END_OF_DAY Then
            If Int(cutoff) = cutoff Then cutoff = cutoff + TimeSerial(23, 59, 59)
        End If
        TryReadCutoff = True
    End If
End Function

Private Sub FindPositionTableColumns(ws As Worksheet, ByRef colCoin As Long, ByRef colPos As Long, ByRef colMkt As Long, ByRef colAvail As Long)
    Dim r As Long: r = 4
    Dim lastC As Long: lastC = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastC
        Dim h As String: h = CStr(Nz(ws.Cells(r, c).Value))
        Select Case NormalizeHeader(h)
            Case "coin": colCoin = c
            Case "position": colPos = c
            Case "marketprice": colMkt = c
            Case "availableqty": colAvail = c
        End Select
    Next c
End Sub

' ========= Low-level helpers =========
Private Function SheetByName(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

Private Function LastRowFast(ws As Worksheet, Optional ByCol As Long = 1) As Long
    LastRowFast = ws.Cells(ws.Rows.Count, ByCol).End(xlUp).Row
End Function

Private Function Nz(ByVal v As Variant, Optional ByVal ifNull As Variant = 0) As Variant
    If IsError(v) Then
        Nz = ifNull
    ElseIf IsEmpty(v) Then
        Nz = ifNull
    ElseIf VarType(v) = vbString And Len(v) = 0 Then
        Nz = ifNull
    Else
        Nz = v
    End If
End Function

Private Function IfNZ(colIdx As Long, ws As Worksheet, rowIdx As Long) As Variant
    If colIdx > 0 Then
        IfNZ = ws.Cells(rowIdx, colIdx).Value
    Else
        IfNZ = Empty
    End If
End Function

Private Function IfNZText(colIdx As Long, ws As Worksheet, rowIdx As Long) As String
    If colIdx > 0 Then
        IfNZText = CStr(Nz(ws.Cells(rowIdx, colIdx).Value))
    Else
        IfNZText = vbNullString
    End If
End Function

Private Function IsMissingVal(ByVal v As Variant) As Boolean
    IsMissingVal = (IsEmpty(v) Or (VarType(v) = vbString And Len(v) = 0))
End Function

Private Function SafeDiv(ByVal a As Double, ByVal b As Double) As Double
    If b = 0 Then SafeDiv = 0 Else SafeDiv = a / b
End Function

Private Function NormalizeHeader(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace$(t, " ", "")
    t = Replace$(t, "_", "")
    t = Replace$(t, "-", "")
    NormalizeHeader = t
End Function

Private Function InAliases(ByVal n As String, aliasesArr As Variant) As Boolean
    Dim x As Variant
    For Each x In aliasesArr
        If n = NormalizeHeader(CStr(x)) Then
            InAliases = True
            Exit Function
        End If
    Next x
End Function

Private Function EndsWithAny(ByVal s As String, arr As Variant) As Boolean
    Dim x As Variant
    For Each x In arr
        If Right$(s, Len(CStr(x))) = UCase$(CStr(x)) Then
            EndsWithAny = True
            Exit Function
        End If
    Next x
End Function

Private Sub CopyKeys(src As Object, dst As Object)
    Dim k As Variant
    If (Not src Is Nothing) Then
        For Each k In src.keys
            If Not dst.Exists(k) Then dst.Add k, True
        Next k
    End If
End Sub

Private Sub PushSession(ByRef arr() As SessionRow, ByRef idx As Long, ByRef s As SessionRow)
    If idx = 0 Then
        ReDim arr(0 To 0)
    Else
        ReDim Preserve arr(0 To idx)
    End If
    arr(idx) = s
    idx = idx + 1
End Sub

Private Sub AggAssignSessions(ByRef agg As PortfolioAggregate, ByRef arr() As SessionRow)
    If (Not (Not arr)) Then
        agg.Sessions = arr
    Else
        Erase agg.Sessions
    End If
End Sub

Private Function IsTodayUTC7(ByVal dt As Date) As Boolean
    IsTodayUTC7 = (DateSerial(Year(dt), Month(dt), Day(dt)) = Date)
End Function


