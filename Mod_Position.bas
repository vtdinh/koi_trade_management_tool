Attribute VB_Name = "mod_position"
Option Explicit
' Last Modified (UTC): 2025-09-13T08:15:00Z

' Batch control: suppress message boxes from Update_All_Position when running multi-day updates
Private gSuppressPositionMsg As Boolean

' Uses global config from mod_config.bas
Private Const OUTPUT_OFFSET_ROWS As Long = 1

Public Sub Update_All_Position()
    On Error GoTo Fail

    Dim wsP As Worksheet, wsO As Worksheet
    Set wsP = SheetByName(mod_config.SHEET_PORTFOLIO)
    Set wsO = SheetByName(mod_config.SHEET_ORDERS)
    If wsP Is Nothing Then Err.Raise 1004, , "Sheet '" & mod_config.SHEET_PORTFOLIO & "' not found."
    If wsO Is Nothing Then Err.Raise 1004, , "Sheet '" & mod_config.SHEET_ORDERS & "' not found."

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim statusMsg As String: statusMsg = vbNullString

    ' --- Detect headers
    Dim hdrP As Long: hdrP = DetectPortfolioHeaderRow(wsP)
    If hdrP = 0 Then Err.Raise 1004, , "Cannot find header row on '" & mod_config.SHEET_PORTFOLIO & "'."
    Dim OUT_START As Long: OUT_START = hdrP + OUTPUT_OFFSET_ROWS

    Dim hdrO As Long: hdrO = DetectOrderHeaderRow(wsO, mod_config.ORDERS_HEADER_ROW_DEFAULT)

    ' --- Map headers
    Dim portCols As Object: Set portCols = MapPortfolioHeaders(wsP, hdrP)
    Dim ordCols As Object: Set ordCols = MapOrderHeaders(wsO, hdrO)

    ' Ensure required columns on Position sheet
    EnsureMapped portCols, "Position"
    EnsureMapped portCols, "Coin"
    EnsureMapped portCols, "Open date"
    EnsureMapped portCols, "Close date"
    EnsureMapped portCols, "Buy Qty"
    EnsureMapped portCols, "Cost"
    EnsureMapped portCols, "Avg. cost"
    EnsureMapped portCols, "sell qty"
    EnsureMapped portCols, "sell proceeds"
    EnsureMapped portCols, "avg sell price"
    EnsureMapped portCols, "available qty"

    ' --- Robust last order row
    Dim lastO As Long
    lastO = Application.WorksheetFunction.Max( _
                LastRowIn(wsO, ordCols("Date"), hdrO), _
                LastRowIn(wsO, ordCols("Type"), hdrO), _
                LastRowIn(wsO, ordCols("Coin"), hdrO), _
                LastRowIn(wsO, ordCols("Qty"), hdrO))
    If lastO < hdrO + 1 Then
        ClearPortfolioOutput wsP, portCols, hdrP, OUT_START
        wsP.Range(mod_config.CELL_CASH).Value = vbNullString
        wsP.Range(mod_config.CELL_COIN).Value = vbNullString
        wsP.Range(mod_config.CELL_NAV).Value = vbNullString
        wsP.Range(mod_config.CELL_SUM_DEPOSIT).Value = vbNullString
        wsP.Range(mod_config.CELL_SUM_WITHDRAW).Value = vbNullString
        wsP.Range(mod_config.CELL_TOTAL_PNL).Value = vbNullString
        statusMsg = "No data in Order_History."
        GoTo Clean
    End If

    ' --- Cutoff from Position!B3 (UTC+7). Date-only -> 23:59:59 of that day.
    Dim cutoffEnabled As Boolean: cutoffEnabled = False
    Dim cutoffDtUTC7 As Date
    If GetCutoffFromPositionB3(cutoffDtUTC7) Then cutoffEnabled = True

    Dim cutoffHi As Date
    If cutoffEnabled Then
        If cutoffDtUTC7 = Int(cutoffDtUTC7) Then
            cutoffHi = DateAdd("s", -1, DateAdd("d", 1, cutoffDtUTC7)) ' 23:59:59 end-of-day
        Else
            cutoffHi = cutoffDtUTC7
        End If
    End If

    Dim dayCutoffUTC7 As Date: dayCutoffUTC7 = IIf(cutoffEnabled, DateValue(cutoffDtUTC7), Date)
    Dim todayUTC7 As Date: todayUTC7 = Date

    ' --- Build sessions & aggregate cash
    Dim sessions As Collection: Set sessions = New Collection
    Dim stateRun As Object: Set stateRun = CreateObject("Scripting.Dictionary"): stateRun.CompareMode = vbTextCompare
    Dim stateSess As Object: Set stateSess = CreateObject("Scripting.Dictionary"): stateSess.CompareMode = vbTextCompare

    Dim totalDeposit As Double, totalWithdraw As Double, totalBuy As Double, totalSell As Double
    totalDeposit = 0#: totalWithdraw = 0#: totalBuy = 0#: totalSell = 0#

    ' Optional "Total" column support
    Dim colTotal As Long: colTotal = 0
    On Error Resume Next
    If ordCols.Exists("Total") Then colTotal = CLng(ordCols("Total"))
    On Error GoTo 0

    Dim r As Long, vDate As Variant, vDateUTC7 As Date
    Dim vSide As String, vCoin As String
    Dim vQty As Double, vPrice As Double, vFee As Double, vEx As String
    Dim vTotal As Double, hasTotal As Boolean
    Dim runQty As Double
    Dim sess As Object

    For r = hdrO + 1 To lastO
        vDate = wsO.Cells(r, ordCols("Date")).Value
        If IsDate(vDate) Then
            vDateUTC7 = OrderToUTC7(CDate(vDate))
            If (Not cutoffEnabled) Or (vDateUTC7 <= cutoffHi) Then
                vSide = UCase$(Trim$(CStr(wsO.Cells(r, ordCols("Type")).Value)))
                vCoin = Trim$(CStr(wsO.Cells(r, ordCols("Coin")).Value))
                vQty = NzD(wsO.Cells(r, ordCols("Qty")).Value)
                If ordCols("Price") > 0 Then vPrice = NzD(wsO.Cells(r, ordCols("Price")).Value) Else vPrice = 0#
                If ordCols("Fee") > 0 Then vFee = NzD(wsO.Cells(r, ordCols("Fee")).Value) Else vFee = 0#
                If ordCols("Exchange") > 0 Then vEx = Trim$(CStr(wsO.Cells(r, ordCols("Exchange")).Value)) Else vEx = ""

                hasTotal = (colTotal > 0 And IsNumeric(wsO.Cells(r, colTotal).Value))
                If hasTotal Then vTotal = NzD(wsO.Cells(r, colTotal).Value) Else vTotal = 0#

                ' ---- Cash legs (prefer Total when available)
                Select Case vSide
                    Case "DEPOSIT"
                        If hasTotal Then totalDeposit = totalDeposit + vTotal Else totalDeposit = totalDeposit + vQty

                    Case "WITHDRAW"
                        If hasTotal Then totalWithdraw = totalWithdraw + vTotal Else totalWithdraw = totalWithdraw + vQty

                    Case "BUY"
                        If hasTotal And vTotal <> 0 Then _
                            totalBuy = totalBuy + vTotal _
                        Else _
                            totalBuy = totalBuy + (vQty * vPrice + vFee)

                    Case "SELL"
                        If hasTotal And vTotal <> 0 Then _
                            totalSell = totalSell + vTotal _
                        Else _
                            totalSell = totalSell + (vQty * vPrice - vFee)
                End Select

                ' ---- Sessions (build P/L rows)
                If vCoin <> "" And vQty <> 0 Then
                    If stateRun.Exists(vCoin) Then runQty = stateRun(vCoin) Else runQty = 0#
                    Select Case vSide
                        Case "BUY"
                            If Abs(runQty) <= mod_config.EPS_CLOSE Then
                                Set sess = NewSession(vCoin, vDateUTC7)
                                Set stateSess(vCoin) = sess
                            Else
                                Set sess = stateSess(vCoin)
                            End If
                            sess("BuyQty") = sess("BuyQty") + vQty
                            sess("Cost") = sess("Cost") + (vQty * vPrice) + vFee
                            If vEx <> "" Then UpdateLatestExchangeInSession sess, vDateUTC7, vEx
                            runQty = runQty + vQty
                            stateRun(vCoin) = runQty
                            Set stateSess(vCoin) = sess

                        Case "SELL"
                            If Not stateSess.Exists(vCoin) Then
                                Set sess = NewSession(vCoin, vDateUTC7)
                            Else
                                Set sess = stateSess(vCoin)
                            End If
                            sess("SellQty") = sess("SellQty") + vQty
                            sess("SellProceeds") = sess("SellProceeds") + (vQty * vPrice) - vFee
                            If vEx <> "" Then UpdateLatestExchangeInSession sess, vDateUTC7, vEx
                            runQty = runQty - vQty
                            stateRun(vCoin) = runQty
                            Set stateSess(vCoin) = sess

                            If runQty <= mod_config.EPS_CLOSE Then
                                sess("CloseDate") = vDateUTC7
                                sess("AvailableQty") = 0#
                                sessions.Add sess
                                stateSess.Remove vCoin
                                stateRun(vCoin) = 0#
                            End If
                    End Select
                End If
            End If
        End If
    Next r

    ' --- Flush remaining open sessions
    Dim k As Variant
    For Each k In stateSess.Keys
        Dim sOpen As Object: Set sOpen = stateSess(k)
        sOpen("AvailableQty") = NzD(stateRun(k))
        sessions.Add sOpen
    Next k

    ' --- Prefetch prices for Open positions only
    Dim priceMap As Object: Set priceMap = CreateObject("Scripting.Dictionary"): priceMap.CompareMode = vbTextCompare
    Dim openExByCoin As Object: Set openExByCoin = CreateObject("Scripting.Dictionary"): openExByCoin.CompareMode = vbTextCompare

    Dim i As Long, ss As Object
    For i = 1 To sessions.Count
        Set ss = sessions(i)
        If ss("AvailableQty") > mod_config.EPS_CLOSE Then
            openExByCoin(ss("Coin")) = CStr(NzStr(ss("Storage")))
        End If
    Next i

    Dim coin As Variant, exName As String, px As Variant
    If openExByCoin.Count > 0 Then
        For Each coin In openExByCoin.Keys
            exName = CStr(openExByCoin(coin))
            Dim sym As String: sym = MapCoinToBinanceSymbol(CStr(coin))
            ' 1) Try Binance first (D1 close for history, realtime for today)
            If dayCutoffUTC7 < todayUTC7 Then
                px = GetBinanceDailyCloseUTC(sym, dayCutoffUTC7)
            Else
                px = GetBinanceRealtimePrice(sym)
            End If
            ' 1b) Binance quote fallback to USDC/BUSD if USDT primary failed
            If (Not IsNumeric(px) Or px <= 0) And Right$(sym, 4) = "USDT" Then
                px = GetFallbackRealtimeOrCloseUTC(sym, dayCutoffUTC7, todayUTC7)
            End If
            ' 2) If Binance unavailable, try the exchange from Order_History (storage)
            If (Not IsNumeric(px) Or px <= 0) And Len(exName) > 0 And LCase$(exName) <> "binance" Then
                px = GetRealtimePriceByExchange(exName, CStr(coin))
            End If
            ' 3) Stablecoin fixed price
            If (Not IsNumeric(px) Or px <= 0) And IsStableCoin(CStr(coin)) Then px = 1#
            If IsNumeric(px) And px > 0 Then priceMap(coin) = CDbl(px)
        Next coin
    End If

    ' --- Clear output columns
    ClearPortfolioOutput wsP, portCols, hdrP, OUT_START

    ' --- Write sessions
    Dim rowOut As Long: rowOut = OUT_START
    Dim mktPrice As Variant, availBal As Variant, profitVal As Variant, pnlPctVal As Variant
    Dim avgBuy As Variant, avgSell As Variant, posText As String
    Dim totalHoldValue As Double: totalHoldValue = 0#
    ' Per-coin holdings for Alt.*_Daily pies
    Dim coinVals As Object: Set coinVals = CreateObject("Scripting.Dictionary"): coinVals.CompareMode = vbTextCompare
    ' Per-row available balance for Open positions to compute %NAV after NAV is known
    Dim openRowVals As Object: Set openRowVals = CreateObject("Scripting.Dictionary"): openRowVals.CompareMode = vbTextCompare

    For i = 1 To sessions.Count
        Set ss = sessions(i)

        If ss("BuyQty") > mod_config.EPS_ZERO Then avgBuy = ss("Cost") / ss("BuyQty") Else avgBuy = vbNullString
        If ss("SellQty") > mod_config.EPS_ZERO Then avgSell = ss("SellProceeds") / ss("SellQty") Else avgSell = vbNullString
        posText = IIf(ss("AvailableQty") > mod_config.EPS_CLOSE, "Open", "Closed")

        ' Market price for Open only
        mktPrice = vbNullString
        If posText = "Open" And portCols("market price") > 0 Then
            If priceMap.Exists(ss("Coin")) Then mktPrice = priceMap(ss("Coin"))
        End If

        ' Available balance (unrealized)
        availBal = vbNullString
        If portCols("available balance") > 0 Then
            If ss("AvailableQty") > mod_config.EPS_CLOSE And IsNumeric(mktPrice) Then
                availBal = ss("AvailableQty") * CDbl(mktPrice)
            ElseIf ss("AvailableQty") <= mod_config.EPS_CLOSE Then
                availBal = 0#
            End If
        End If
        If IsNumeric(availBal) Then totalHoldValue = totalHoldValue + CDbl(availBal)

        ' PnL & %PnL
        profitVal = vbNullString: pnlPctVal = vbNullString
        If ss("AvailableQty") <= mod_config.EPS_CLOSE Then
            profitVal = ss("SellProceeds") - ss("Cost")
            If ss("Cost") > 0 And IsNumeric(profitVal) Then pnlPctVal = CDbl(profitVal) / ss("Cost")
        Else
            If IsNumeric(availBal) Then
                profitVal = ss("SellProceeds") + CDbl(availBal) - ss("Cost")
                If ss("Cost") > 0 Then pnlPctVal = CDbl(profitVal) / ss("Cost")
            End If
        End If

        ' Round & write
        WriteCellSafe wsP, rowOut, portCols("Buy Qty"), RoundN(ss("BuyQty"), 3)
        WriteCellSafe wsP, rowOut, portCols("sell qty"), RoundN(ss("SellQty"), 3)
        WriteCellSafe wsP, rowOut, portCols("available qty"), RoundN(ss("AvailableQty"), 3)

        If IsNumeric(ss("Cost")) Then ss("Cost") = RoundN(ss("Cost"), 0)
        ' Keep average prices with price precision (do not round to 0)
        If IsNumeric(avgBuy) Then avgBuy = RoundN(avgBuy, mod_config.ROUND_PRICE_DECIMALS)
        If IsNumeric(avgSell) Then avgSell = RoundN(avgSell, mod_config.ROUND_PRICE_DECIMALS)
        If IsNumeric(ss("SellProceeds")) Then ss("SellProceeds") = RoundN(ss("SellProceeds"), 0)
        If IsNumeric(availBal) Then availBal = RoundN(availBal, 0)
        If IsNumeric(profitVal) Then profitVal = RoundN(profitVal, 0)

        WriteCellSafe wsP, rowOut, portCols("Position"), posText
        WriteCellSafe wsP, rowOut, portCols("Coin"), ss("Coin")
        WriteCellSafe wsP, rowOut, portCols("Open date"), ss("OpenDate")
        If Not IsEmpty(ss("CloseDate")) Then WriteCellSafe wsP, rowOut, portCols("Close date"), ss("CloseDate")
        WriteCellSafe wsP, rowOut, portCols("Cost"), ss("Cost")
        WriteCellSafe wsP, rowOut, portCols("Avg. cost"), avgBuy
        WriteCellSafe wsP, rowOut, portCols("sell proceeds"), ss("SellProceeds")
        WriteCellSafe wsP, rowOut, portCols("avg sell price"), avgSell

        If portCols("market price") > 0 Then
            If IsNumeric(mktPrice) Then
                WriteCellSafe wsP, rowOut, portCols("market price"), mktPrice
            ElseIf mod_config.CLEAR_MARKET_PRICE Then
                WriteCellSafe wsP, rowOut, portCols("market price"), vbNullString
            End If
        End If
        If portCols("available balance") > 0 Then _
            WriteCellSafe wsP, rowOut, portCols("available balance"), availBal
        ' Collect for %NAV write later when totalNAV is available
        If StrComp(posText, "Open", vbTextCompare) = 0 Then
            If IsNumeric(availBal) Then openRowVals(CStr(rowOut)) = CDbl(availBal)
        End If

        ' Accumulate per-coin value for open positions
        If posText = "Open" Then
            ' In VBA, logical And is not short-circuit; evaluate IsNumeric first to avoid
            ' type mismatch when availBal is vbNullString/Empty.
            If IsNumeric(availBal) Then
                If CDbl(availBal) > 0 Then
                    Dim ckey As String: ckey = CStr(ss("Coin"))
                    If coinVals.Exists(ckey) Then
                        coinVals(ckey) = CDbl(coinVals(ckey)) + CDbl(availBal)
                    Else
                        coinVals(ckey) = CDbl(availBal)
                    End If
                End If
            End If
        End If

        ' Profit color
        If portCols("profit") > 0 Then
            WriteCellSafe wsP, rowOut, portCols("profit"), profitVal
            If IsNumeric(profitVal) Then
                If profitVal > 0 Then
                    wsP.Cells(rowOut, portCols("profit")).Font.Color = mod_config.COLOR_PNL_POSITIVE
                ElseIf profitVal < 0 Then
                    wsP.Cells(rowOut, portCols("profit")).Font.Color = mod_config.COLOR_PNL_NEGATIVE
                Else
                    wsP.Cells(rowOut, portCols("profit")).Font.Color = vbBlack
                End If
            Else
                wsP.Cells(rowOut, portCols("profit")).Font.Color = vbBlack
            End If
        End If

        ' %PnL color
        If portCols("%PnL") > 0 Then
            WriteCellSafe wsP, rowOut, portCols("%PnL"), pnlPctVal
            If IsNumeric(pnlPctVal) Then
                If pnlPctVal > 0 Then
                    wsP.Cells(rowOut, portCols("%PnL")).Font.Color = mod_config.COLOR_PNL_POSITIVE
                ElseIf pnlPctVal < 0 Then
                    wsP.Cells(rowOut, portCols("%PnL")).Font.Color = mod_config.COLOR_PNL_NEGATIVE
                Else
                    wsP.Cells(rowOut, portCols("%PnL")).Font.Color = vbBlack
                End If
            Else
                wsP.Cells(rowOut, portCols("%PnL")).Font.Color = vbBlack
            End If
        End If

        If portCols("storage") > 0 Then WriteCellSafe wsP, rowOut, portCols("storage"), ss("Storage")

        rowOut = rowOut + 1
    Next i

    ' --- Dashboard totals (Cash, Coin, NAV) + Deposit/Withdraw + Total P/L
    Dim totalCash As Double, totalNAV As Double
    Dim sumDeposit As Double, sumWithdraw As Double, totalPnL As Double

    totalCash = (totalDeposit + totalSell) - (totalBuy + totalWithdraw)
    totalNAV = totalCash + totalHoldValue
    sumDeposit = totalDeposit
    sumWithdraw = totalWithdraw
    totalPnL = totalNAV - (sumDeposit - sumWithdraw)

    With wsP
        .Range(mod_config.CELL_CASH).Value = Round(totalCash, 0)
        .Range(mod_config.CELL_COIN).Value = Round(totalHoldValue, 0)
        .Range(mod_config.CELL_NAV).Value = Application.WorksheetFunction.RoundDown(totalNAV, 0)

        .Range(mod_config.CELL_SUM_DEPOSIT).Value = Round(sumDeposit, 0)
        .Range(mod_config.CELL_SUM_WITHDRAW).Value = Round(sumWithdraw, 0)
        .Range(mod_config.CELL_TOTAL_PNL).Value = Round(totalPnL, 0)

        .Range(mod_config.CELL_CASH & ":" & mod_config.CELL_CASH).NumberFormat = mod_config.MONEY_FMT
        .Range(mod_config.CELL_COIN & ":" & mod_config.CELL_COIN).NumberFormat = mod_config.MONEY_FMT
        .Range(mod_config.CELL_NAV & ":" & mod_config.CELL_NAV).NumberFormat = mod_config.MONEY_FMT
        .Range(mod_config.CELL_SUM_DEPOSIT & ":" & mod_config.CELL_SUM_DEPOSIT).NumberFormat = mod_config.MONEY_FMT
        .Range(mod_config.CELL_SUM_WITHDRAW & ":" & mod_config.CELL_SUM_WITHDRAW).NumberFormat = mod_config.MONEY_FMT
        .Range(mod_config.CELL_TOTAL_PNL & ":" & mod_config.CELL_TOTAL_PNL).NumberFormat = mod_config.MONEY_FMT
    End With

    ' --- Now write %NAV weights
    If portCols("%NAV") > 0 Then
        If totalNAV > mod_config.EPS_CLOSE Then
            For Each k In openRowVals.Keys
                wsP.Cells(CLng(k), portCols("%NAV")).Value = CDbl(openRowVals(k)) / totalNAV
            Next k
        Else
            For Each k In openRowVals.Keys
                wsP.Cells(CLng(k), portCols("%NAV")).Value = 0#
            Next k
        End If
    End If

    ' --- NAV metrics over last 3 months (from Daily_Snapshot)
    UpdateNav3M_MetricsAndChart wsP, dayCutoffUTC7, totalNAV
    ' --- Allocation metrics (%Coin, %BTC, %Alt.*, Num Alt.* coins)
    UpdateAllocationMetrics wsP, coinVals, totalHoldValue, totalNAV

    ' --- Now write %NAV weights (after NAV known) and sort by %NAV desc
    If portCols("%NAV") > 0 Then
        If totalNAV > mod_config.EPS_CLOSE Then
            For Each k In openRowVals.Keys
                wsP.Cells(CLng(k), portCols("%NAV")).Value = CDbl(openRowVals(k)) / totalNAV
            Next k
        Else
            For Each k In openRowVals.Keys
                wsP.Cells(CLng(k), portCols("%NAV")).Value = 0#
            Next k
        End If
        ' Fill blanks to zero for consistent sorting
        Dim rFill As Long
        For rFill = OUT_START To rowOut - 1
            If Not IsNumeric(wsP.Cells(rFill, portCols("%NAV")).Value) Then
                wsP.Cells(rFill, portCols("%NAV")).Value = 0#
            End If
        Next rFill
        ' Sort entire block by %NAV desc
        Dim sortRange As Range, keyRange As Range
        Set sortRange = wsP.Range(wsP.Cells(OUT_START, 1), wsP.Cells(rowOut - 1, MaxCol(portCols)))
        Set keyRange = wsP.Range(wsP.Cells(OUT_START, portCols("%NAV")), wsP.Cells(rowOut - 1, portCols("%NAV")))
        On Error Resume Next
        sortRange.Sort Key1:=keyRange, Order1:=xlDescending, Header:=xlNo
        On Error GoTo 0
    End If

    ' --- Portfolio rule checks (limits in column C → actions in column D)
    On Error Resume Next
    CheckPortfolioRuleViolations wsP
    On Error GoTo 0

    ' --- Update charts
    ' Cash vs Coin (keep chart type)
    UpdateCashCoinChart wsP
    ' Portfolio breakdown (BTC / Alt.TOP / Alt.MID / Alt.LOW)
    Update_Portfolio1_FromCategory
    ' Portfolio2 per-coin pie removed per request
    ' Alt daily pies: per-coin within Alt.TOP/MID/LOW (on Position sheet)
    UpdatePortfolioAltDailyPies wsP, coinVals

    ' --- Clear tail & format
    Dim lastRowPos As Long: lastRowPos = rowOut - 1
    If lastRowPos < 50 Then wsP.Range(wsP.Rows(lastRowPos + 1), wsP.Rows(50)).ClearContents

    SafeFormat wsP, portCols, rowOut - 1, hdrP, OUT_START
    If Len(statusMsg) = 0 Then statusMsg = "Positions, dashboard, and charts updated."

    ' Capital rule check – write warning if NAV diff within threshold per config
    On Error Resume Next
    checkCapitalRuleViolation
    On Error GoTo 0

Clean:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If (Not gSuppressPositionMsg) And Len(statusMsg) > 0 Then MsgBox statusMsg, vbInformation
    Exit Sub

Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub

' =============================================================================
' ======= INDEPENDENT MACRO: UPDATE MARKET PRICE (OPEN ONLY) ==================
' =============================================================================
' Cutoff is from Position!B3 (UTC+7). We decide:
' - If cutoff DAY (UTC+7) < today (UTC+7): fetch Binance D1 close by interpreting the cutoff DATE as a UTC calendar day.
' - Else: fetch realtime.
Public Sub Update_MarketPrice_ByCutoff_OpenOnly_Simple()
    On Error GoTo Fail

    Dim wsP As Worksheet
    Set wsP = SheetByName(mod_config.SHEET_PORTFOLIO)
    If wsP Is Nothing Then Err.Raise 1004, , "Sheet '" & mod_config.SHEET_PORTFOLIO & "' not found."

    Dim cutoffUTC7 As Date
    If Not GetCutoffFromPositionB3(cutoffUTC7) Then
        MsgBox "Cutoff is missing/invalid. Please fill Position!B3 (UTC+7), e.g., 2025-08-31 or 2025-08-31 15:30.", vbExclamation
        GoTo Clean
    End If

    Dim dayCutoffUTC7 As Date: dayCutoffUTC7 = DateValue(cutoffUTC7)
    Dim todayUTC7 As Date: todayUTC7 = Date

    Dim hdrRow As Long: hdrRow = DetectPortfolioHeaderRow(wsP)
    If hdrRow = 0 Then Err.Raise 1004, , "Header not found on Position."
    Dim OUT_START As Long: OUT_START = hdrRow + 1

    Dim portCols As Object: Set portCols = MapPortfolioHeaders(wsP, hdrRow)
    EnsureMapped portCols, "Coin"
    EnsureMapped portCols, "Position"
    EnsureMapped portCols, "market price"

    Dim lastR As Long: lastR = LastRowIn(wsP, portCols("Coin"), hdrRow)
    If lastR < OUT_START Then
        MsgBox "No rows to update.", vbInformation
        GoTo Clean
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim r As Long, coin As String, posVal As String, sym As String
    Dim px As Variant

    For r = OUT_START To lastR
        coin = Trim$(CStr(wsP.Cells(r, portCols("Coin")).Value))
        posVal = Trim$(CStr(wsP.Cells(r, portCols("Position")).Value))

        If Len(coin) > 0 And StrComp(posVal, "Open", vbTextCompare) = 0 Then
            ' Prefer Binance first (realtime or D1 close), then fallback to the row's exchange (OKX/Bybit)
            Dim exName As String: exName = ""
            If portCols("storage") > 0 Then exName = Trim$(CStr(wsP.Cells(r, portCols("storage")).Value))
            sym = MapCoinToBinanceSymbol(coin)
            If dayCutoffUTC7 < todayUTC7 Then
                Dim dayUTC As Date: dayUTC = DateValue(cutoffUTC7)
                px = GetBinanceDailyCloseUTC(sym, dayUTC)
            Else
                px = GetBinanceRealtimePrice(sym)
            End If
            If (Not IsNumeric(px) Or px <= 0) And Right$(sym, 4) = "USDT" Then
                px = GetFallbackRealtimeOrCloseUTC(sym, dayCutoffUTC7, todayUTC7)
            End If
            If (Not IsNumeric(px) Or px <= 0) And Len(exName) > 0 And LCase$(exName) <> "binance" Then
                px = GetRealtimePriceByExchange(exName, coin)
            End If
            ' Stablecoin -> 1
            If (Not IsNumeric(px) Or px <= 0) And IsStableCoin(coin) Then px = 1#

            ' Write or clear
            If IsNumeric(px) And px > 0 Then
                wsP.Cells(r, portCols("market price")).Value = CDbl(px)
            Else
                wsP.Cells(r, portCols("market price")).ClearContents
            End If
        End If
    Next r

    On Error Resume Next
    wsP.Range(wsP.Cells(OUT_START, portCols("market price")), wsP.Cells(lastR, portCols("market price"))).NumberFormat = mod_config.PRICE_FMT
    On Error GoTo 0

    MsgBox "Market price updated for Open rows (UTC D1 close / realtime).", vbInformation
Clean:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub

' =============================================================================
' ======================== CHART UPDATE HELPERS ===============================
' =============================================================================
Private Sub UpdateAllocationMetrics(wsP As Worksheet, ByVal coinVals As Object, ByVal totalCoinVal As Double, ByVal totalNAV As Double)
    On Error GoTo Done
    ' Build percent format based on config
    Dim fmtPct As String
    If mod_config.ROUND_PCT_DECIMALS <= 0 Then
        fmtPct = "0%"
    Else
        fmtPct = "0." & String(mod_config.ROUND_PCT_DECIMALS, "0") & "%"
    End If

    ' %Coin = Coin / NAV
    If Len(mod_config.CELL_PCT_COIN) > 0 Then
        Dim pctCoin As Double
        If totalNAV > mod_config.EPS_CLOSE Then pctCoin = totalCoinVal / totalNAV Else pctCoin = 0#
        pctCoin = RoundN(pctCoin, mod_config.ROUND_PCT_DECIMALS + 2)
        wsP.Range(mod_config.CELL_PCT_COIN).Value = pctCoin
        On Error Resume Next: wsP.Range(mod_config.CELL_PCT_COIN).NumberFormat = fmtPct: On Error GoTo 0
    End If

    If coinVals Is Nothing Then GoTo Done

    ' Build Category mapping
    Dim wsC As Worksheet
    Set wsC = SheetByName(mod_config.SHEET_CATEGORY)
    If wsC Is Nothing Then Set wsC = SheetByName("Categoty")
    If wsC Is Nothing Then Set wsC = SheetByName("Catagory")
    If wsC Is Nothing Then Set wsC = SheetByName("Category")
    Dim mapCG As Object: Set mapCG = BuildCoinToGroupFromCategorySheet(wsC)

    Dim vBTC As Double: vBTC = 0#
    Dim vTop As Double, vMid As Double, vLow As Double
    Dim cTop As Long, cMid As Long, cLow As Long
    vTop = 0#: vMid = 0#: vLow = 0#
    cTop = 0: cMid = 0: cLow = 0

    Dim k As Variant, name As String, grp As String, v As Double
    For Each k In coinVals.Keys
        name = CStr(k)
        v = CDbl(coinVals(k))
        If v > mod_config.EPS_CLOSE Then
            If UCase$(name) = "BTC" Then
                vBTC = vBTC + v
            Else
                grp = vbNullString
                If Not (mapCG Is Nothing) And mapCG.Exists(name) Then grp = CStr(mapCG(name))
                Select Case NormalizeHeader(grp)
                    Case NormalizeHeader("Alt.TOP"): vTop = vTop + v: cTop = cTop + 1
                    Case NormalizeHeader("Alt.MID"): vMid = vMid + v: cMid = cMid + 1
                    Case NormalizeHeader("Alt.LOW"): vLow = vLow + v: cLow = cLow + 1
                End Select
            End If
        End If
    Next k

    ' %BTC
    If Len(mod_config.CELL_PCT_BTC) > 0 Then
        Dim pctBtc As Double: If totalCoinVal > mod_config.EPS_CLOSE Then pctBtc = vBTC / totalCoinVal Else pctBtc = 0#
        pctBtc = RoundN(pctBtc, mod_config.ROUND_PCT_DECIMALS + 2)
        wsP.Range(mod_config.CELL_PCT_BTC).Value = pctBtc
        On Error Resume Next: wsP.Range(mod_config.CELL_PCT_BTC).NumberFormat = fmtPct: On Error GoTo 0
    End If
    ' %Alt.TOP and count
    If Len(mod_config.CELL_PCT_ALT_TOP) > 0 Then
        Dim pctTop As Double: If totalCoinVal > mod_config.EPS_CLOSE Then pctTop = vTop / totalCoinVal Else pctTop = 0#
        pctTop = RoundN(pctTop, mod_config.ROUND_PCT_DECIMALS + 2)
        wsP.Range(mod_config.CELL_PCT_ALT_TOP).Value = pctTop
        On Error Resume Next: wsP.Range(mod_config.CELL_PCT_ALT_TOP).NumberFormat = fmtPct: On Error GoTo 0
    End If
    If Len(mod_config.CELL_NUM_ALT_TOP) > 0 Then
        wsP.Range(mod_config.CELL_NUM_ALT_TOP).Value = cTop
        On Error Resume Next: wsP.Range(mod_config.CELL_NUM_ALT_TOP).NumberFormat = mod_config.MONEY_FMT: On Error GoTo 0
    End If
    ' %Alt.MID and count
    If Len(mod_config.CELL_PCT_ALT_MID) > 0 Then
        Dim pctMid As Double: If totalCoinVal > mod_config.EPS_CLOSE Then pctMid = vMid / totalCoinVal Else pctMid = 0#
        pctMid = RoundN(pctMid, mod_config.ROUND_PCT_DECIMALS + 2)
        wsP.Range(mod_config.CELL_PCT_ALT_MID).Value = pctMid
        On Error Resume Next: wsP.Range(mod_config.CELL_PCT_ALT_MID).NumberFormat = fmtPct: On Error GoTo 0
    End If
    If Len(mod_config.CELL_NUM_ALT_MID) > 0 Then
        wsP.Range(mod_config.CELL_NUM_ALT_MID).Value = cMid
        On Error Resume Next: wsP.Range(mod_config.CELL_NUM_ALT_MID).NumberFormat = mod_config.MONEY_FMT: On Error GoTo 0
    End If
    ' %Alt.LOW and count
    If Len(mod_config.CELL_PCT_ALT_LOW) > 0 Then
        Dim pctLow As Double: If totalCoinVal > mod_config.EPS_CLOSE Then pctLow = vLow / totalCoinVal Else pctLow = 0#
        pctLow = RoundN(pctLow, mod_config.ROUND_PCT_DECIMALS + 2)
        wsP.Range(mod_config.CELL_PCT_ALT_LOW).Value = pctLow
        On Error Resume Next: wsP.Range(mod_config.CELL_PCT_ALT_LOW).NumberFormat = fmtPct: On Error GoTo 0
    End If
    If Len(mod_config.CELL_NUM_ALT_LOW) > 0 Then
        wsP.Range(mod_config.CELL_NUM_ALT_LOW).Value = cLow
        On Error Resume Next: wsP.Range(mod_config.CELL_NUM_ALT_LOW).NumberFormat = mod_config.MONEY_FMT: On Error GoTo 0
    End If
    Exit Sub
Done:
End Sub

' ============================= CAPITAL RULE CHECK ============================
Public Sub checkCapitalRuleViolation()
    On Error GoTo Done
    Dim ws As Worksheet
    Set ws = SheetByName(mod_config.SHEET_PORTFOLIO)
    If ws Is Nothing Then Exit Sub

    Dim addrCalc As String, addrReal As String, addrAct As String
    addrCalc = mod_config.CELL_NAV
    addrReal = mod_config.CELL_NAV_REAL
    addrAct = mod_config.CELL_NAV_ACTION
    If Len(addrCalc) = 0 Or Len(addrReal) = 0 Or Len(addrAct) = 0 Then Exit Sub

    Dim vCalc As Variant, vReal As Variant
    vCalc = ws.Range(addrCalc).Value
    vReal = ws.Range(addrReal).Value
    Dim thr As Double: thr = mod_config.CAPITAL_RULE_DIFF_THRESHOLD_PCT

    Dim warn As Boolean: warn = False
    If IsNumeric(vCalc) And IsNumeric(vReal) Then
        Dim calcD As Double, realD As Double
        calcD = CDbl(vCalc): realD = CDbl(vReal)
        If Abs(realD) > mod_config.EPS_CLOSE Then
            Dim diffPct As Double
            diffPct = Abs(calcD - realD) / Abs(realD)
            ' Warn when difference is large (>= threshold)
            If diffPct >= thr Then warn = True
        End If
    End If

    WriteActionText ws, addrAct, IIf(warn, "Check NAV !", vbNullString), IIf(warn, vbRed, vbBlack), IIf(warn, True, False)
Done:
End Sub

' Write text safely to a (possibly merged) cell address and apply basic formatting
Private Sub WriteActionText(ByVal ws As Worksheet, ByVal addr As String, ByVal text As String, ByVal color As Long, ByVal isBold As Boolean)
    On Error Resume Next
    Dim rg As Range
    Set rg = ws.Range(addr)
    If rg.MergeCells Then Set rg = rg.MergeArea.Cells(1, 1)
    rg.Value = text
    rg.Font.Color = color
    rg.Font.Bold = isBold
    On Error GoTo 0
End Sub

' ============================= PORTFOLIO RULES ===============================
Private Sub CheckPortfolioRuleViolations(wsP As Worksheet)
    On Error GoTo Done
    Dim addrs(1 To 8) As String
    addrs(1) = mod_config.CELL_PCT_COIN
    addrs(2) = mod_config.CELL_PCT_BTC
    addrs(3) = mod_config.CELL_PCT_ALT_TOP
    addrs(4) = mod_config.CELL_NUM_ALT_TOP
    addrs(5) = mod_config.CELL_PCT_ALT_MID
    addrs(6) = mod_config.CELL_NUM_ALT_MID
    addrs(7) = mod_config.CELL_PCT_ALT_LOW
    addrs(8) = mod_config.CELL_NUM_ALT_LOW

    Dim i As Long
    For i = 1 To 8
        If Len(addrs(i)) > 0 Then
            Dim r As Long
            r = wsP.Range(addrs(i)).Row
            Dim valV As Variant, limV As Variant, label As String
            valV = wsP.Range(addrs(i)).Value
            limV = wsP.Cells(r, 3).Value   ' column C = limit
            label = CStr(wsP.Cells(r, 1).Value) ' column A = label
            If IsNumeric(valV) And IsNumeric(limV) Then
                Dim v As Double, lim As Double
                v = CDbl(valV): lim = CDbl(limV)
                If lim > 0 And v >= lim Then
                    WriteActionText wsP, wsP.Cells(r, 4).Address(False, False), "Reduce " & label, vbRed, True
                Else
                    WriteActionText wsP, wsP.Cells(r, 4).Address(False, False), vbNullString, vbBlack, False
                End If
            Else
                WriteActionText wsP, wsP.Cells(r, 4).Address(False, False), vbNullString, vbBlack, False
            End If
        End If
    Next i
Done:
End Sub
Private Sub UpdateNav3M_MetricsAndChart(wsP As Worksheet, ByVal cutoffDate As Date, ByVal fallbackNav As Double)
    On Error GoTo Done
    Dim wsS As Worksheet: Set wsS = SheetByName(mod_config.SHEET_SNAPSHOT)
    If wsS Is Nothing Then GoTo Done

    Dim fromDt As Date: fromDt = DateAdd("m", -3, cutoffDate)
    Dim lastR As Long: lastR = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row

    Dim r As Long, dt As Date, nav As Double
    Dim cnt As Long: cnt = 0
    Dim navATH As Double: navATH = 0#: Dim navATL As Double: navATL = 0#
    Dim navAtCutoff As Variant: navAtCutoff = Empty

    ' Count rows in window
    For r = 2 To lastR
        If IsDate(wsS.Cells(r, 1).Value) Then
            dt = DateValue(wsS.Cells(r, 1).Value)
            If dt >= fromDt And dt <= cutoffDate Then
                cnt = cnt + 1
            End If
        End If
    Next r

    Dim xs() As Variant, ys() As Double
    If cnt > 0 Then
        ReDim xs(1 To cnt)
        ReDim ys(1 To cnt)
    End If

    Dim i As Long: i = 0
    For r = 2 To lastR
        If IsDate(wsS.Cells(r, 1).Value) Then
            dt = DateValue(wsS.Cells(r, 1).Value)
            If dt >= fromDt And dt <= cutoffDate Then
                i = i + 1
                xs(i) = dt
                If IsNumeric(wsS.Cells(r, 4).Value) Then nav = CDbl(wsS.Cells(r, 4).Value) Else nav = 0#
                ys(i) = nav
                If nav > 0 Then
                    If navATH = 0# Or nav > navATH Then navATH = nav
                    If navATL = 0# Or nav < navATL Then navATL = nav
                End If
                If dt = cutoffDate Then navAtCutoff = nav
            End If
        End If
    Next r

    If IsEmpty(navAtCutoff) Then navAtCutoff = fallbackNav
    ' Ensure ATH/ATL consider today's live NAV (fallbackNav) even if snapshot row is missing
    If IsNumeric(fallbackNav) Then
        If CDbl(fallbackNav) > navATH Then navATH = CDbl(fallbackNav)
        If navATL = 0# Or CDbl(fallbackNav) < navATL Then navATL = CDbl(fallbackNav)
    End If

    ' Write metrics if configured
    If Len(mod_config.CELL_NAV_ATH) > 0 Then wsP.Range(mod_config.CELL_NAV_ATH).Value = Application.WorksheetFunction.RoundDown(navATH, 0)
    If Len(mod_config.CELL_NAV_ATL) > 0 Then wsP.Range(mod_config.CELL_NAV_ATL).Value = Application.WorksheetFunction.RoundDown(navATL, 0)
    If Len(mod_config.CELL_NAV_DD) > 0 Then
        Dim dd As Double, baseAth As Double, baseCut As Double
        baseAth = navATH
        baseCut = CDbl(navAtCutoff)
        If mod_config.NAV_DD_USE_TRUNCATED Then
            baseAth = Application.WorksheetFunction.RoundDown(baseAth, 0)
            baseCut = Application.WorksheetFunction.RoundDown(baseCut, 0)
        End If
        If baseAth > 0 Then
            dd = (baseAth - baseCut) / baseAth
            If dd < 0 Then dd = 0# ' never negative
            If dd <= mod_config.NAV_DD_TOLERANCE_PCT Then dd = 0#
        Else
            dd = 0#
        End If
        wsP.Range(mod_config.CELL_NAV_DD).Value = dd
        On Error Resume Next
        wsP.Range(mod_config.CELL_NAV_DD).NumberFormat = mod_config.PCT_FMT
        On Error GoTo 0

        ' Compare with threshold in CELL_NAV_DD_LIMIT and write action in CELL_NAV_DD_ACTION
        Dim lim As Variant
        lim = wsP.Range(mod_config.CELL_NAV_DD_LIMIT).Value
        If IsNumeric(lim) Then
            Dim limD As Double: limD = CDbl(lim)
            If dd >= limD And limD > 0 Then
                WriteActionText wsP, mod_config.CELL_NAV_DD_ACTION, " Cut loss NOW!!!", vbRed, True
            Else
                WriteActionText wsP, mod_config.CELL_NAV_DD_ACTION, vbNullString, vbBlack, False
            End If
        End If
    End If
    On Error Resume Next
    If Len(mod_config.CELL_NAV_ATH) > 0 Then wsP.Range(mod_config.CELL_NAV_ATH).NumberFormat = mod_config.MONEY_FMT
    If Len(mod_config.CELL_NAV_ATL) > 0 Then wsP.Range(mod_config.CELL_NAV_ATL).NumberFormat = mod_config.MONEY_FMT
    On Error GoTo 0

    ' Update NAV 3M chart
    Dim co As ChartObject, ch As Chart
    Dim created As Boolean: created = False
    Set co = Nothing
    On Error Resume Next
    Set co = wsP.ChartObjects("NAV 3M")
    On Error GoTo 0
    If co Is Nothing Then
        Set co = wsP.ChartObjects.Add(Left:=20, Top:=500, Width:=520, Height:=260)
        co.Name = "NAV 3M"
        co.Chart.HasTitle = True
        co.Chart.ChartTitle.Text = "NAV 3M"
        created = True
    End If
    Set ch = co.Chart
    ' Ensure chart can plot data from hidden helper column
    On Error Resume Next
    ch.PlotVisibleOnly = False
    On Error GoTo 0
    ' Hide legend for a cleaner single-series chart
    On Error Resume Next
    ch.HasLegend = False
    On Error GoTo 0

    Dim s As Series
    If ch.SeriesCollection.Count = 0 Then Set s = ch.SeriesCollection.NewSeries Else Set s = ch.SeriesCollection(1)
    Dim k As Long
    For k = ch.SeriesCollection.Count To 2 Step -1: ch.SeriesCollection(k).Delete: Next k

    If cnt > 0 Then
        ' Bind X to a worksheet date range so tooltips show dates, not serials
        Dim helperCol As Long: helperCol = wsP.Columns.Count - 1 ' second-to-last column (hidden)
        Dim dest As Range: Set dest = wsP.Range(wsP.Cells(2, helperCol), wsP.Cells(1 + cnt, helperCol))
        Dim v2() As Variant: ReDim v2(1 To cnt, 1 To 1)
        For i = 1 To cnt
            v2(i, 1) = CDate(xs(i))
        Next i
        dest.Value = v2
        On Error Resume Next
        dest.NumberFormat = mod_config.POS_DATE_AXIS_FMT
        dest.NumberFormatLocal = mod_config.POS_DATE_AXIS_FMT
        On Error GoTo 0
        s.XValues = dest
        s.Values = ys
        ' Hide helper column only after binding XValues
        On Error Resume Next
        wsP.Columns(helperCol).Hidden = True
        On Error GoTo 0
    Else
        On Error Resume Next
        ch.SeriesCollection(1).Delete
        On Error GoTo 0
    End If
    ch.ChartType = xlLine
    ' Apply axis options when chart is created fresh (as per request)
    If created Then
        On Error Resume Next
        With ch.Axes(xlCategory)
            .CategoryType = xlTimeScale           ' Date axis
            .BaseUnit = xlDays
            .TickLabels.NumberFormatLinked = False
            .TickLabels.NumberFormat = mod_config.POS_DATE_AXIS_FMT
            .TickLabels.NumberFormatLocal = mod_config.POS_DATE_AXIS_FMT
        End With
        On Error GoTo 0
    End If

    ' Ensure date axis format is applied even when chart already exists
    On Error Resume Next
    With ch.Axes(xlCategory)
        .CategoryType = xlTimeScale
        .BaseUnit = xlDays
        .TickLabels.NumberFormatLinked = False
        .TickLabels.NumberFormat = mod_config.POS_DATE_AXIS_FMT
        .TickLabels.NumberFormatLocal = mod_config.POS_DATE_AXIS_FMT
    End With
    On Error GoTo 0

    ' ----- Thousand-rounded Y-axis bounds based on 3M NAV range
    Dim yMin As Double, yMax As Double, unit As Double
    unit = 1000#
    If navATL > 0 Then
        yMin = navATL * 0.95   ' 5% margin below
    Else
        yMin = 0#
    End If
    If navATH > 0 Then
        yMax = navATH * 1.05   ' 5 margin above
    Else
        yMax = 1#
    End If
    ' Round to thousand boundaries: floor for min, ceiling for max
    yMin = unit * Int(yMin / unit)
    yMax = unit * Int((yMax + unit - 1#) / unit)
    If yMax <= yMin Then yMax = yMin + unit

    On Error Resume Next
    With ch.Axes(xlValue)
        .MinimumScaleIsAuto = False
        .MaximumScaleIsAuto = False
        .MinimumScale = yMin
        .MaximumScale = yMax
    End With
    On Error GoTo 0
    ch.Axes(xlValue).TickLabels.NumberFormat = mod_config.MONEY_FMT
    On Error GoTo 0
    Exit Sub
Done:
End Sub
Private Sub UpdateCashCoinChart(wsP As Worksheet)
    On Error GoTo Done
    Dim co As ChartObject
    Set co = Nothing

    ' 1) Try by ChartObject.Name
    On Error Resume Next
    Set co = wsP.ChartObjects("Cash vs Coin")
    On Error GoTo 0

    ' 2) Try by Chart Title text (case-insensitive)
    If co Is Nothing Then
        Dim coIt As ChartObject
        For Each coIt In wsP.ChartObjects
            If Not coIt Is Nothing Then
                If coIt.Chart.HasTitle Then
                    If StrComp(coIt.Chart.ChartTitle.Text, "Cash vs Coin", vbTextCompare) = 0 Then
                        Set co = coIt
                        Exit For
                    End If
                End If
            End If
        Next coIt
    End If

    If co Is Nothing Then GoTo Done

    Dim cashVal As Double, coinVal As Double
    cashVal = 0#: coinVal = 0#
    If IsNumeric(wsP.Range(mod_config.CELL_CASH).Value) Then cashVal = CDbl(wsP.Range(mod_config.CELL_CASH).Value)
    If IsNumeric(wsP.Range(mod_config.CELL_COIN).Value) Then coinVal = CDbl(wsP.Range(mod_config.CELL_COIN).Value)

    ' Pie charts do not support negative values. Guard and skip if invalid.
    If cashVal < 0 Or coinVal < 0 Then GoTo Done

    Dim ch As Chart
    Set ch = co.Chart

    Dim s As Series
    If ch.SeriesCollection.Count = 0 Then
        Set s = ch.SeriesCollection.NewSeries
    Else
        Set s = ch.SeriesCollection(1)
    End If

    ' Ensure only one series (first) remains
    Dim i As Long
    For i = ch.SeriesCollection.Count To 2 Step -1
        ch.SeriesCollection(i).Delete
    Next i

    Dim vals As Variant, names As Variant
    vals = Array(cashVal, coinVal)
    names = Array("Cash", "Coin")
    s.Values = vals
    s.XValues = names
    s.HasDataLabels = True
    On Error Resume Next
    s.DataLabels.ShowCategoryName = True
    s.DataLabels.ShowPercentage = True
    s.DataLabels.ShowValue = False
    s.DataLabels.NumberFormat = "0%"
    On Error GoTo 0

Done:
End Sub

' Update the "Portfolio1" chart (pie) with portions by group: BTC, Alt.TOP, Alt.MID, Alt.LOW.
' Group definitions are read from sheet mod_config.SHEET_CATEGORY with headers: Coin | Group (row 1).
Public Sub Update_Portfolio1_FromCategory()
    On Error GoTo Clean

    Dim wsP As Worksheet, wsC As Worksheet
    Set wsP = SheetByName(mod_config.SHEET_PORTFOLIO)
    If wsP Is Nothing Then Err.Raise 1004, , "Sheet '" & mod_config.SHEET_PORTFOLIO & "' not found."
    ' Proactively normalize any legacy chart object/title to the new canonical name
    EnsureCoinCategoryChartName wsP, mod_config.CHART_PORTFOLIO1
    Set wsC = SheetByName(mod_config.SHEET_CATEGORY)
    ' Fallbacks for common spellings
    If wsC Is Nothing Then Set wsC = SheetByName("Category")
    If wsC Is Nothing Then Set wsC = SheetByName("Categories")
    If wsC Is Nothing Then Err.Raise 1004, , "Category sheet not found (tried '" & mod_config.SHEET_CATEGORY & "', 'Category', 'Categories')."

    ' 1) Build per-coin available balance from Position (Open rows)
    Dim hdrP As Long: hdrP = DetectPortfolioHeaderRow(wsP)
    If hdrP = 0 Then Err.Raise 1004, , "Cannot find header row on '" & mod_config.SHEET_PORTFOLIO & "'."
    Dim OUT_START As Long: OUT_START = hdrP + OUTPUT_OFFSET_ROWS
    Dim portCols As Object: Set portCols = MapPortfolioHeaders(wsP, hdrP)
    Dim lastPos As Long: lastPos = LastRowIn(wsP, portCols("Coin"), hdrP)

    Dim holdDict As Object: Set holdDict = CreateObject("Scripting.Dictionary")
    holdDict.CompareMode = vbTextCompare

    Dim r As Long, posTxt As String, coin As String
    Dim availBal As Double, qty As Double, mkt As Double

    For r = OUT_START To lastPos
        posTxt = CStr(wsP.Cells(r, portCols("Position")).Value)
        If LCase$(Trim$(posTxt)) = "open" Then
            coin = Trim$(CStr(wsP.Cells(r, portCols("Coin")).Value))
            If Len(coin) > 0 Then
                ' Prefer Available Balance if exists; else compute qty*price
                availBal = 0#
                If portCols("available balance") > 0 Then
                    If IsNumeric(wsP.Cells(r, portCols("available balance")).Value) Then
                        availBal = CDbl(wsP.Cells(r, portCols("available balance")).Value)
                    End If
                End If
                If availBal = 0# Then
                    qty = 0#: mkt = 0#
                    If portCols("available qty") > 0 And IsNumeric(wsP.Cells(r, portCols("available qty")).Value) Then _
                        qty = CDbl(wsP.Cells(r, portCols("available qty")).Value)
                    If portCols("market price") > 0 And IsNumeric(wsP.Cells(r, portCols("market price")).Value) Then _
                        mkt = CDbl(wsP.Cells(r, portCols("market price")).Value)
                    availBal = qty * mkt
                End If
                If availBal > 0 Then
                    If holdDict.Exists(coin) Then
                        holdDict(coin) = holdDict(coin) + availBal
                    Else
                        holdDict(coin) = availBal
                    End If
                End If
            End If
        End If
    Next r

    ' 2) Read coin -> group mapping from Category sheet
    Dim coinToGroup As Object
    Set coinToGroup = BuildCoinToGroupFromCategorySheet(wsC)

    ' 3) Aggregate into the four groups
    Dim gVals As Object: Set gVals = CreateObject("Scripting.Dictionary")
    gVals.CompareMode = vbTextCompare
    gVals("BTC") = 0#: gVals("Alt.TOP") = 0#: gVals("Alt.MID") = 0#: gVals("Alt.LOW") = 0#

    Dim k As Variant, grp As String, v As Double
    For Each k In holdDict.Keys
        v = CDbl(holdDict(k))
        If UCase$(CStr(k)) = "BTC" Then
            grp = "BTC"
        ElseIf coinToGroup.Exists(CStr(k)) Then
            grp = CStr(coinToGroup(k))
        Else
            ' Unmapped coin → skip (do not block chart update/rename)
            grp = vbNullString
        End If
        If gVals.Exists(grp) Then gVals(grp) = gVals(grp) + v
    Next k

    ' 4) Update chart "Portfolio1" on Position sheet or chart sheet
    Dim ok As Boolean
    ok = UpdatePortfolio1Chart(wsP, gVals)
    ' Silent on success/failure to avoid extra popups; top-level macro will report
Clean:
    End Sub

Private Function UpdatePortfolio1Chart(wsP As Worksheet, gVals As Object) As Boolean
    On Error GoTo Done
    Dim co As ChartObject, coIt As ChartObject
    Set co = Nothing
    On Error Resume Next
    Set co = wsP.ChartObjects(mod_config.CHART_PORTFOLIO1)
    On Error GoTo 0
    If co Is Nothing Then
        For Each coIt In wsP.ChartObjects
            If Not coIt Is Nothing Then
                If coIt.Chart.HasTitle Then
                    If StrComp(coIt.Chart.ChartTitle.Text, mod_config.CHART_PORTFOLIO1, vbTextCompare) = 0 Then
                        Set co = coIt
                        Exit For
                    End If
                End If
                ' Legacy names support: rename to new canonical name
                If co Is Nothing Then
                    If StrComp(coIt.Name, "Portfolio1", vbTextCompare) = 0 _
                       Or StrComp(coIt.Name, "Portfolio 1", vbTextCompare) = 0 _
                       Or StrComp(coIt.Name, "Portfolio_Category_Daily", vbTextCompare) = 0 _
                       Or (coIt.Chart.HasTitle And StrComp(coIt.Chart.ChartTitle.Text, "Portfolio1", vbTextCompare) = 0) _
                       Or (coIt.Chart.HasTitle And StrComp(coIt.Chart.ChartTitle.Text, "Portfolio 1", vbTextCompare) = 0) _
                       Or (coIt.Chart.HasTitle And StrComp(coIt.Chart.ChartTitle.Text, "Portfolio_Category_Daily", vbTextCompare) = 0) Then
                        On Error Resume Next
                        coIt.Name = mod_config.CHART_PORTFOLIO1
                        coIt.Chart.HasTitle = True
                        coIt.Chart.ChartTitle.Text = mod_config.CHART_PORTFOLIO1
                        On Error GoTo 0
                        Set co = coIt
                        Exit For
                    End If
                End If
            End If
        Next coIt
    End If
    If co Is Nothing Then
        ' Also search chart sheets by name or title
        Dim chs As Chart
        For Each chs In ThisWorkbook.Charts
            If StrComp(chs.Name, mod_config.CHART_PORTFOLIO1, vbTextCompare) = 0 Then
                UpdatePortfolio1Chart = ApplyPortfolioSeriesToChart(chs, gVals)
                Exit Function
            End If
            If chs.HasTitle Then
                If StrComp(chs.ChartTitle.Text, mod_config.CHART_PORTFOLIO1, vbTextCompare) = 0 Then
                    UpdatePortfolio1Chart = ApplyPortfolioSeriesToChart(chs, gVals)
                    Exit Function
                End If
            End If
            ' Legacy chart sheet names/titles -> rename then use
            If StrComp(chs.Name, "Portfolio1", vbTextCompare) = 0 Or StrComp(chs.Name, "Portfolio 1", vbTextCompare) = 0 Or StrComp(chs.Name, "Portfolio_Category_Daily", vbTextCompare) = 0 Then
                On Error Resume Next
                chs.Name = mod_config.CHART_PORTFOLIO1
                chs.HasTitle = True
                chs.ChartTitle.Text = mod_config.CHART_PORTFOLIO1
                On Error GoTo 0
                UpdatePortfolio1Chart = ApplyPortfolioSeriesToChart(chs, gVals)
                Exit Function
            End If
            If chs.HasTitle Then
                If StrComp(chs.ChartTitle.Text, "Portfolio1", vbTextCompare) = 0 Or StrComp(chs.ChartTitle.Text, "Portfolio 1", vbTextCompare) = 0 Or StrComp(chs.ChartTitle.Text, "Portfolio_Category_Daily", vbTextCompare) = 0 Then
                    On Error Resume Next
                    chs.HasTitle = True
                    chs.ChartTitle.Text = mod_config.CHART_PORTFOLIO1
                    On Error GoTo 0
                    UpdatePortfolio1Chart = ApplyPortfolioSeriesToChart(chs, gVals)
                    Exit Function
                End If
            End If
        Next chs
        ' Not found anywhere -> create an embedded pie chart on Position
        Set co = CreatePortfolio1Chart(wsP)
        If co Is Nothing Then GoTo Done
    End If

    UpdatePortfolio1Chart = ApplyPortfolioSeriesToChart(co.Chart, gVals)

Done:
End Function

' Ensure any legacy chart object or title is renamed to the canonical coin category name
Private Sub EnsureCoinCategoryChartName(wsP As Worksheet, ByVal canonical As String)
    On Error Resume Next
    Dim co As ChartObject
    ' Direct match by old names
    For Each co In wsP.ChartObjects
        If co Is Nothing Then GoTo NextOne
        If StrComp(co.Name, "Portfolio_Category_Daily", vbTextCompare) = 0 _
           Or StrComp(co.Name, "Portfolio1", vbTextCompare) = 0 _
           Or StrComp(co.Name, "Portfolio 1", vbTextCompare) = 0 Then
            co.Name = canonical
            co.Chart.HasTitle = True
            co.Chart.ChartTitle.Text = canonical
        ElseIf co.Chart.HasTitle Then
            If StrComp(co.Chart.ChartTitle.Text, "Portfolio_Category_Daily", vbTextCompare) = 0 _
               Or StrComp(co.Chart.ChartTitle.Text, "Portfolio1", vbTextCompare) = 0 _
               Or StrComp(co.Chart.ChartTitle.Text, "Portfolio 1", vbTextCompare) = 0 Then
                co.Name = canonical
                co.Chart.HasTitle = True
                co.Chart.ChartTitle.Text = canonical
            End If
        End If
NextOne:
    Next co
    On Error GoTo 0
End Sub

' Strong normalization: rename legacy charts to canonical or delete duplicates (embedded and chart sheets)
 ' Removed unused ForceNormalizeCoinCategoryCharts (legacy normalization no longer called)

Private Function ApplyPortfolioSeriesToChart(ByVal ch As Chart, ByVal gVals As Object) As Boolean
    On Error GoTo Fail
    Dim s As Series
    If ch.SeriesCollection.Count = 0 Then
        Set s = ch.SeriesCollection.NewSeries
    Else
        Set s = ch.SeriesCollection(1)
    End If
    Dim i As Long
    For i = ch.SeriesCollection.Count To 2 Step -1
        ch.SeriesCollection(i).Delete
    Next i

    Dim names As Variant, vals As Variant
    names = Array("BTC", "Alt.TOP", "Alt.MID", "Alt.LOW")
    vals = Array(CDbl(NzD(gVals("BTC"))), CDbl(NzD(gVals("Alt.TOP"))), CDbl(NzD(gVals("Alt.MID"))), CDbl(NzD(gVals("Alt.LOW"))))

    ' If all zeros, do not update (avoid pie issues)
    If (vals(0) + vals(1) + vals(2) + vals(3)) <= 0 Then
        ResetPieChartToNoHoldings ch
        ApplyPortfolioSeriesToChart = True
        Exit Function
    End If

    s.XValues = names
    s.Values = vals
    s.HasDataLabels = True
    On Error Resume Next
    s.DataLabels.ShowCategoryName = True
    s.DataLabels.ShowPercentage = True
    s.DataLabels.ShowValue = False
    s.DataLabels.NumberFormat = "0%"
    On Error GoTo 0
    ApplyPortfolioSeriesToChart = True
    Exit Function
Fail:
    ApplyPortfolioSeriesToChart = False
End Function

Private Function BuildCoinToGroupFromCategorySheet(wsC As Worksheet) As Object
    ' Supports two layouts:
    ' A) Two-column mapping with headers in row 1: Coin | Group (or Category/Catagory)
    ' B) Multi-column where row 1 cells are group names (BTC, Alt.TOP, Alt.MID, Alt.LOW)
    '    and each column lists coins under that group starting row 2.
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare

    Dim lastCol As Long: lastCol = wsC.Cells(1, wsC.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Set BuildCoinToGroupFromCategorySheet = d: Exit Function

    Dim coinCol As Long, groupCol As Long, c As Long
    coinCol = 0: groupCol = 0
    ' First, try detect two-column layout
    For c = 1 To lastCol
        Dim h As String: h = NormalizeHeader(CStr(wsC.Cells(1, c).Value))
        If h = "coin" Or h = "symbol" Or h = "asset" Then coinCol = c
        If h = "group" Or h = "category" Or h = "catagory" Then groupCol = c
    Next c

    Dim r As Long
    If coinCol > 0 And groupCol > 0 Then
        Dim lastRow As Long: lastRow = wsC.Cells(wsC.Rows.Count, coinCol).End(xlUp).Row
        For r = 2 To lastRow
            Dim cc As String, gg As String
            cc = Trim$(CStr(wsC.Cells(r, coinCol).Value))
            gg = Trim$(CStr(wsC.Cells(r, groupCol).Value))
            If Len(cc) > 0 And Len(gg) > 0 Then d(cc) = gg
        Next r
        Set BuildCoinToGroupFromCategorySheet = d
        Exit Function
    End If

    ' Otherwise, treat each header in row 1 as a group, coins listed below
    For c = 1 To lastCol
        Dim grp As String: grp = Trim$(CStr(wsC.Cells(1, c).Value))
        If Len(grp) > 0 Then
            Dim lastR As Long: lastR = wsC.Cells(wsC.Rows.Count, c).End(xlUp).Row
            For r = 2 To lastR
                Dim coin As String: coin = Trim$(CStr(wsC.Cells(r, c).Value))
                If Len(coin) > 0 Then d(coin) = grp
            Next r
        End If
    Next c

    Set BuildCoinToGroupFromCategorySheet = d
End Function

Private Function CreatePortfolio1Chart(wsP As Worksheet) As ChartObject
    On Error GoTo Fail
    Dim co As ChartObject
    ' Pick a default placement near the dashboard area
    Set co = wsP.ChartObjects.Add(Left:=300, Top:=20, Width:=360, Height:=220)
    co.Name = mod_config.CHART_PORTFOLIO1
    With co.Chart
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = mod_config.CHART_PORTFOLIO1
        ' Initialize with placeholders so chart renders
        Dim vals As Variant, names As Variant
        names = Array("BTC", "Alt.TOP", "Alt.MID", "Alt.LOW")
        vals = Array(1, 1, 1, 1)
        Dim s As Series
        If .SeriesCollection.Count = 0 Then
            Set s = .SeriesCollection.NewSeries
        Else
            Set s = .SeriesCollection(1)
        End If
        s.XValues = names
        s.Values = vals
        s.HasDataLabels = True
        On Error Resume Next
        s.DataLabels.ShowCategoryName = True
        s.DataLabels.ShowPercentage = True
        s.DataLabels.ShowValue = False
        s.DataLabels.NumberFormat = "0%"
        On Error GoTo 0
    End With
    Set CreatePortfolio1Chart = co
    Exit Function
Fail:
    Set CreatePortfolio1Chart = Nothing
End Function


' Create or update three pie charts on Position sheet showing per-coin breakdowns
' within Alt.TOP, Alt.MID, and Alt.LOW groups for the current cutoff day.
Private Sub UpdatePortfolioAltDailyPies(wsP As Worksheet, coinVals As Object)
    On Error GoTo Done
    Dim wsC As Worksheet
    Set wsC = SheetByName(mod_config.SHEET_CATEGORY)
    If wsC Is Nothing Then Set wsC = SheetByName("Categoty")
    If wsC Is Nothing Then Set wsC = SheetByName("Catagory")
    If wsC Is Nothing Then Set wsC = SheetByName("Category")
    Dim mapCG As Object: Set mapCG = BuildCoinToGroupFromCategorySheet(wsC)

    Dim topVals As Object, midVals As Object, lowVals As Object
    Set topVals = CreateObject("Scripting.Dictionary"): topVals.CompareMode = vbTextCompare
    Set midVals = CreateObject("Scripting.Dictionary"): midVals.CompareMode = vbTextCompare
    Set lowVals = CreateObject("Scripting.Dictionary"): lowVals.CompareMode = vbTextCompare

    If Not (coinVals Is Nothing) Then
        Dim k As Variant, grp As String
        For Each k In coinVals.Keys
            grp = ""
            If UCase$(CStr(k)) = "BTC" Then
                grp = "BTC"
            ElseIf Not (mapCG Is Nothing) And mapCG.Exists(CStr(k)) Then
                grp = CStr(mapCG(CStr(k)))
            End If
            Select Case NormalizeHeader(grp)
                Case NormalizeHeader("Alt.TOP")
                    topVals(CStr(k)) = CDbl(coinVals(k))
                Case NormalizeHeader("Alt.MID")
                    midVals(CStr(k)) = CDbl(coinVals(k))
                Case NormalizeHeader("Alt.LOW")
                    lowVals(CStr(k)) = CDbl(coinVals(k))
            End Select
        Next k
    End If

    ApplyPerCoinPie wsP, "Alt.TOP", topVals
    ApplyPerCoinPie wsP, "Alt.MID", midVals
    ApplyPerCoinPie wsP, "Alt.LOW", lowVals
Done:
End Sub

Private Sub ApplyPerCoinPie(wsP As Worksheet, ByVal chartName As String, ByVal vals As Object)
    On Error GoTo Fail
    Dim co As ChartObject
    Set co = Nothing
    On Error Resume Next
    Set co = wsP.ChartObjects(chartName)
    On Error GoTo 0
    ' If not found, try legacy names and rename to the new canonical chartName
    If co Is Nothing Then
        Dim alt As Variant, alts As Variant
        Select Case chartName
            Case "Alt.TOP": alts = Array("Portfolio_Alt.TOP_Daily", "Portfolio Alt.TOP Daily")
            Case "Alt.MID": alts = Array("Portfolio_Alt.MID_Daily", "Portfolio Alt.MID Daily")
            Case "Alt.LOW": alts = Array("Portfolio_Alt.LOW_Daily", "Portfolio Alt.LOW Daily")
            Case Else:      alts = Array()
        End Select
        For Each alt In alts
            On Error Resume Next
            Set co = wsP.ChartObjects(CStr(alt))
            On Error GoTo 0
            If Not co Is Nothing Then
                On Error Resume Next
                co.Name = chartName
                co.Chart.HasTitle = True
                co.Chart.ChartTitle.Text = chartName
                On Error GoTo 0
                Exit For
            End If
        Next alt
    End If
    If co Is Nothing Then
        Set co = wsP.ChartObjects.Add(Left:=20, Top:=260, Width:=360, Height:=220)
        co.Name = chartName
        With co.Chart
            .ChartType = xlPie
            .HasTitle = True
            .ChartTitle.Text = chartName
            .HasLegend = True
            Dim s As Series
            If .SeriesCollection.Count = 0 Then Set s = .SeriesCollection.NewSeries Else Set s = .SeriesCollection(1)
            s.XValues = Array("Coin A", "Coin B")
            s.Values = Array(1, 1)
            s.HasDataLabels = True
            On Error Resume Next
            s.DataLabels.ShowCategoryName = True
            s.DataLabels.ShowPercentage = True
            s.DataLabels.ShowValue = False
            s.DataLabels.NumberFormat = "0%"
            On Error GoTo 0
        End With
    End If

    With co.Chart
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = chartName
        .HasLegend = True
    End With

    ApplyPerCoinSeriesToChart co.Chart, vals
    Exit Sub
Fail:
End Sub

Private Sub ApplyPerCoinSeriesToChart(ByVal ch As Chart, ByVal coinVals As Object)
    On Error GoTo Fail
    If coinVals Is Nothing Then
        ResetPieChartToNoHoldings ch
        Exit Sub
    End If
    If coinVals.Count = 0 Then
        ResetPieChartToNoHoldings ch
        Exit Sub
    End If

    Dim s As Series
    If ch.SeriesCollection.Count = 0 Then
        Set s = ch.SeriesCollection.NewSeries
    Else
        Set s = ch.SeriesCollection(1)
    End If
    Dim i As Long
    For i = ch.SeriesCollection.Count To 2 Step -1
        ch.SeriesCollection(i).Delete
    Next i

    ' Build arrays from dictionary
    Dim n As Long: n = coinVals.Count
    Dim names() As Variant, vals() As Double
    ReDim names(1 To n)
    ReDim vals(1 To n)
    Dim k As Variant, idx As Long: idx = 1
    For Each k In coinVals.Keys
        names(idx) = CStr(k)
        vals(idx) = CDbl(coinVals(k))
        idx = idx + 1
    Next k

    ' If totals are zero, skip
    Dim sumV As Double
    For i = 1 To n
        sumV = sumV + vals(i)
    Next i
    If sumV <= 0 Then
        ResetPieChartToNoHoldings ch
        Exit Sub
    End If

    s.XValues = names
    s.Values = vals
    s.HasDataLabels = True
    On Error Resume Next
    s.DataLabels.ShowCategoryName = True
    s.DataLabels.ShowPercentage = True
    s.DataLabels.ShowValue = False
    s.DataLabels.NumberFormat = "0%"
    On Error GoTo 0
    Exit Sub
Fail:
End Sub


' When there are no open holdings, reset a pie chart to a single "No holdings" slice
Private Sub ResetPieChartToNoHoldings(ByVal ch As Chart)
    On Error Resume Next
    Dim i As Long
    For i = ch.SeriesCollection.Count To 1 Step -1
        ch.SeriesCollection(i).Delete
    Next i
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.XValues = Array("No holdings")
    s.Values = Array(1)
    s.HasDataLabels = True
    s.DataLabels.ShowCategoryName = True
    s.DataLabels.ShowPercentage = True
    s.DataLabels.ShowValue = False
    s.DataLabels.NumberFormat = "0%"
End Sub


' =========================== SESSION HELPERS =================================
Private Function NewSession(ByVal coin As String, ByVal openDt As Date) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    d("Coin") = coin
    d("OpenDate") = openDt
    d("CloseDate") = Empty
    d("BuyQty") = 0#
    d("Cost") = 0#
    d("SellQty") = 0#
    d("SellProceeds") = 0#
    d("AvailableQty") = 0#
    d("Storage") = ""
    d("StorageDate") = 0#
    Set NewSession = d
End Function

Private Sub UpdateLatestExchangeInSession(ByRef sess As Object, ByVal dt As Date, ByVal exch As String)
    If NzD(sess("StorageDate")) = 0# Or dt >= sess("StorageDate") Then
        sess("StorageDate") = dt
        sess("Storage") = exch
    End If
End Sub

' Convert an Order_History timestamp (UTC-4) to UTC+7
Private Function OrderToUTC7(ByVal dt As Date) As Date
    OrderToUTC7 = dt + (mod_config.ORDERS_TZ_OFFSET_HOURS / 24#)
End Function


' ====================== MAPPING / FORMATTING HELPERS =========================
Private Function SheetByName(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    Dim target As String, cand As String

    target = LCase$(Replace$(Replace$(Trim$(nm), " ", ""), "_", ""))

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.name, nm, vbTextCompare) = 0 Then
            Set SheetByName = ws
            Exit Function
        End If
    Next ws

    For Each ws In ThisWorkbook.Worksheets
        cand = LCase$(Replace$(Replace$(Trim$(ws.name), " ", ""), "_", ""))
        If cand = target Then
            Set SheetByName = ws
            Exit Function
        End If
    Next ws

    Set SheetByName = Nothing
End Function

Private Function DetectPortfolioHeaderRow(ws As Worksheet) As Long
    Dim r As Long, lastC As Long, raw As Object
    For r = 1 To Application.Min(30, ws.Rows.Count)
        lastC = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If lastC >= 1 Then
            Set raw = BuildHeaderRaw(ws, r)
            If (raw.Exists("position") Or raw.Exists("status") Or raw.Exists("state") Or raw.Exists("pos")) _
               And raw.Exists("coin") Then
                DetectPortfolioHeaderRow = r: Exit Function
            End If
        End If
    Next r
    DetectPortfolioHeaderRow = 0
End Function

' Auto-detect header row on Order_History (Date+Coin+Qty present)
Private Function DetectOrderHeaderRow(ws As Worksheet, Optional ByVal defaultHeaderRow As Long = 2) As Long
    Dim r As Long, raw As Object

    Set raw = BuildHeaderRaw(ws, defaultHeaderRow)
    If (raw.Exists("date") Or raw.Exists("time") Or raw.Exists("trade date") Or raw.Exists("open time")) _
       And (raw.Exists("coin") Or raw.Exists("symbol") Or raw.Exists("asset") Or raw.Exists("ticker")) _
       And (raw.Exists("qty") Or raw.Exists("quantity") Or raw.Exists("amount") Or raw.Exists("size")) Then
        DetectOrderHeaderRow = defaultHeaderRow: Exit Function
    End If

    For r = 1 To Application.Min(10, ws.Rows.Count)
        Set raw = BuildHeaderRaw(ws, r)
        If (raw.Exists("date") Or raw.Exists("time") Or raw.Exists("trade date") Or raw.Exists("open time")) _
           And (raw.Exists("coin") Or raw.Exists("symbol") Or raw.Exists("asset") Or raw.Exists("ticker")) _
           And (raw.Exists("qty") Or raw.Exists("quantity") Or raw.Exists("amount") Or raw.Exists("size")) Then
            DetectOrderHeaderRow = r: Exit Function
        End If
    Next r

    DetectOrderHeaderRow = defaultHeaderRow
End Function

Private Sub EnsureMapped(ByVal d As Object, ByVal key As String)
    If d Is Nothing Then Err.Raise 1004, , "Internal mapping is Nothing."
    If Not d.Exists(key) Then Err.Raise 1004, , "Missing header (one of): " & key
    If CLng(d(key)) < 1 Then Err.Raise 1004, , "Invalid column index for key: " & key
End Sub

Private Sub WriteCellSafe(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal v As Variant)
    If c >= 1 Then ws.Cells(r, c).Value = v
End Sub

Private Function BuildHeaderRaw(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Dim lastCol As Long: lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1
    Dim c As Long, key As String
    For c = 1 To lastCol
        key = NormalizeHeader(CStr(ws.Cells(headerRow, c).Value))
        If key <> "" Then d(key) = c
    Next c
    Set BuildHeaderRaw = d
End Function

Private Function MapPortfolioHeaders(wsP As Worksheet, ByVal headerRow As Long) As Object
    Dim raw As Object: Set raw = BuildHeaderRaw(wsP, headerRow)
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary"): map.CompareMode = vbTextCompare

    map("Position") = RequireAny(raw, Array("position", "status", "state", "pos"))
    map("Coin") = RequireAny(raw, Array("coin", "symbol", "asset", "ticker"))
    map("Open date") = RequireAny(raw, Array("open date", "open"))
    map("Close date") = RequireAny(raw, Array("close date", "close"))
    map("Buy Qty") = RequireAny(raw, Array("buy qty", "buy quantity", "buyqty", "buy", "buy q"))
    map("Cost") = RequireAny(raw, Array("cost", "buy cost", "total cost"))
    map("Avg. cost") = RequireAny(raw, Array("avg cost", "avg. cost", "average cost"))
    map("sell qty") = RequireAny(raw, Array("sell qty", "sell quantity", "sellqty", "sold qty", "sell q"))
    map("sell proceeds") = RequireAny(raw, Array("sell proceeds", "net proceeds", "sell money", "sell value"))
    map("avg sell price") = RequireAny(raw, Array("avg sell price", "average sell price"))
    map("available qty") = RequireAny(raw, Array("available qty", "netqty", "available", "remain qty", "remaining qty"))

    map("market price") = RequireAnyOptional(raw, Array("market price", "last price", "price"))
    map("available balance") = RequireAnyOptional(raw, Array("available balance", "balance", "unrealized value", "market value"))
    map("profit") = RequireAnyOptional(raw, Array("profit", "pnl", "p&l", "gain"))
    map("%PnL") = RequireAnyOptional(raw, Array("%pnl", "pnl%", "percent pnl", "% pnl", "roi", "return %", "return pct"))
    map("%NAV") = RequireAnyOptional(raw, Array("%nav", "nav%", "% of nav", "% nav", "weight", "%weight", "nav weight"))
    map("storage") = RequireAnyOptional(raw, Array("storage", "exchange", "venue", "wallet"))

    If Not map.Exists("market price") Then map("market price") = 0
    If Not map.Exists("available balance") Then map("available balance") = 0
    If Not map.Exists("profit") Then map("profit") = 0
    If Not map.Exists("%PnL") Then map("%PnL") = 0
    If Not map.Exists("storage") Then map("storage") = 0
    If Not map.Exists("%NAV") Then map("%NAV") = 0

    Set MapPortfolioHeaders = map
End Function

Private Function MapOrderHeaders(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim raw As Object: Set raw = BuildHeaderRaw(ws, headerRow)
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary"): map.CompareMode = vbTextCompare

    map("Date") = RequireAny(raw, Array("date", "time", "trade date", "open time"))
    map("Type") = RequireAny(raw, Array("type", "side", "action"))
    map("Coin") = RequireAny(raw, Array("coin", "symbol", "asset", "ticker", "pair base", "base"))
    map("Qty") = RequireAny(raw, Array("qty", "quantity", "amount", "size"))
    map("Price") = RequireAnyOptional(raw, Array("price", "unit price", "avg price", "fill price", "rate"))
    map("Fee") = RequireAnyOptional(raw, Array("fee", "commission"))
    map("Exchange") = RequireAnyOptional(raw, Array("exchange", "venue", "market", "wallet"))
    map("Total") = RequireAnyOptional(raw, Array("total", "amount", "gross", "value"))


    If Not map.Exists("Price") Then map("Price") = 0
    If Not map.Exists("Fee") Then map("Fee") = 0
    If Not map.Exists("Exchange") Then map("Exchange") = 0
    If Not map.Exists("Total") Then map("Total") = 0

    Set MapOrderHeaders = map
End Function

Private Function NormalizeHeader(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, Chr(160), " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    t = Replace(t, """", "")
    t = Replace(t, "'", "")
    t = Replace(t, ".", "")
    t = LCase$(Trim$(t))
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormalizeHeader = t
End Function

Private Function RequireAny(ByVal dict As Object, ByVal aliases As Variant) As Long
    Dim i As Long, k As String
    For i = LBound(aliases) To UBound(aliases)
        k = NormalizeHeader(CStr(aliases(i)))
        If dict.Exists(k) Then RequireAny = dict(k): Exit Function
    Next i
    Err.Raise 1004, , "Missing header (one of): " & JoinAliases(aliases)
End Function

Private Function RequireAnyOptional(ByVal dict As Object, ByVal aliases As Variant) As Long
    Dim i As Long, k As String
    For i = LBound(aliases) To UBound(aliases)
        k = NormalizeHeader(CStr(aliases(i)))
        If dict.Exists(k) Then RequireAnyOptional = dict(k): Exit Function
    Next i
    RequireAnyOptional = 0
End Function

Private Function JoinAliases(ByVal aliases As Variant) As String
    Dim i As Long, arr() As String
    ReDim arr(LBound(aliases) To UBound(aliases))
    For i = LBound(aliases) To UBound(aliases)
        arr(i) = CStr(aliases(i))
    Next i
    JoinAliases = Join(arr, " | ")
End Function

Private Sub ClearPortfolioOutput(ws As Worksheet, ByVal portCols As Object, ByVal headerRow As Long, ByVal outStart As Long)
    Dim lastR As Long: lastR = LastRowIn(ws, portCols("Coin"), headerRow)
    If lastR < outStart Then lastR = outStart
    Dim k As Variant, col As Long
    For Each k In portCols.Keys
        col = portCols(k)
        If col > 0 Then
            If LCase$(CStr(k)) = "market price" And Not mod_config.CLEAR_MARKET_PRICE Then
                ' keep
            Else
                ws.Range(ws.Cells(outStart, col), ws.Cells(lastR, col)).ClearContents
            End If
        End If
    Next k
End Sub

Private Function LastRowIn(ws As Worksheet, ByVal col As Long, ByVal headerRow As Long) As Long
    If col < 1 Then
        Dim f As Range
        Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If f Is Nothing Then LastRowIn = headerRow Else LastRowIn = f.Row
    ElseIf Application.WorksheetFunction.CountA(ws.Columns(col)) = 0 Then
        LastRowIn = headerRow
    Else
        LastRowIn = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    End If
End Function

Private Sub SafeFormat(ws As Worksheet, portCols As Object, lastRow As Long, ByVal headerRow As Long, ByVal outStart As Long)
    If lastRow < outStart Then Exit Sub
    On Error Resume Next
    Dim qtyFmt As String: qtyFmt = GetOrderQtyNumberFormat()
    If Len(qtyFmt) = 0 Then qtyFmt = "0." & String(mod_config.ROUND_QTY_DECIMALS, "0")
    ws.Range(ws.Cells(outStart, portCols("Buy Qty")), ws.Cells(lastRow, portCols("Buy Qty"))).NumberFormat = qtyFmt
    ws.Range(ws.Cells(outStart, portCols("sell qty")), ws.Cells(lastRow, portCols("sell qty"))).NumberFormat = qtyFmt
    ws.Range(ws.Cells(outStart, portCols("available qty")), ws.Cells(lastRow, portCols("available qty"))).NumberFormat = qtyFmt

    ws.Range(ws.Cells(outStart, portCols("Cost")), ws.Cells(lastRow, portCols("Cost"))).NumberFormat = mod_config.MONEY_FMT
    ws.Range(ws.Cells(outStart, portCols("sell proceeds")), ws.Cells(lastRow, portCols("sell proceeds"))).NumberFormat = mod_config.MONEY_FMT
    ws.Range(ws.Cells(outStart, portCols("Avg. cost")), ws.Cells(lastRow, portCols("Avg. cost"))).NumberFormat = mod_config.PRICE_FMT
    If portCols("available balance") > 0 Then _
        ws.Range(ws.Cells(outStart, portCols("available balance")), ws.Cells(lastRow, portCols("available balance"))).NumberFormat = mod_config.MONEY_FMT
    If portCols("profit") > 0 Then _
        ws.Range(ws.Cells(outStart, portCols("profit")), ws.Cells(lastRow, portCols("profit"))).NumberFormat = mod_config.MONEY_FMT

    If portCols("market price") > 0 Then _
        ws.Range(ws.Cells(outStart, portCols("market price")), ws.Cells(lastRow, portCols("market price"))).NumberFormat = mod_config.PRICE_FMT
    ws.Range(ws.Cells(outStart, portCols("avg sell price")), ws.Cells(lastRow, portCols("avg sell price"))).NumberFormat = mod_config.PRICE_FMT
    If portCols("%PnL") > 0 Then _
        ws.Range(ws.Cells(outStart, portCols("%PnL")), ws.Cells(lastRow, portCols("%PnL"))).NumberFormat = mod_config.PCT_FMT
    If portCols("%NAV") > 0 Then _
        ws.Range(ws.Cells(outStart, portCols("%NAV")), ws.Cells(lastRow, portCols("%NAV"))).NumberFormat = mod_config.PCT_FMT

    ws.Range(ws.Cells(outStart, portCols("Open date")), ws.Cells(lastRow, portCols("Open date"))).NumberFormat = mod_config.DATE_FMT
    ws.Range(ws.Cells(outStart, portCols("Close date")), ws.Cells(lastRow, portCols("Close date"))).NumberFormat = mod_config.DATE_FMT

    ' Respect config: keep existing column widths unless explicitly allowed
    If mod_config.AUTOFIT_POSITION_COLUMNS Then
        ' Avoid auto-fitting column A: start from column 2 if needed
        Dim startCol As Long: startCol = MinCol(portCols)
        If startCol < 2 Then startCol = 2
        ws.Range(ws.Cells(headerRow, startCol), ws.Cells(lastRow, MaxCol(portCols))).EntireColumn.AutoFit
    End If
    On Error GoTo 0
End Sub
' Return the NumberFormat for the Qty column in Order_History
Private Function GetOrderQtyNumberFormat() As String
    On Error GoTo Done
    Dim wsO As Worksheet: Set wsO = SheetByName(mod_config.SHEET_ORDERS)
    If wsO Is Nothing Then GoTo Done
    Dim hdrO As Long: hdrO = DetectOrderHeaderRow(wsO, mod_config.ORDERS_HEADER_ROW_DEFAULT)
    If hdrO = 0 Then GoTo Done
    Dim ordCols As Object: Set ordCols = MapOrderHeaders(wsO, hdrO)
    If Not ordCols.Exists("Qty") Then GoTo Done
    Dim c As Long: c = CLng(ordCols("Qty"))
    If c <= 0 Then GoTo Done
    ' Prefer the first data cell's NumberFormat below the header if available
    Dim lastR As Long: lastR = wsO.Cells(wsO.Rows.Count, c).End(xlUp).Row
    Dim r As Long
    For r = hdrO + 1 To lastR
        If IsNumeric(wsO.Cells(r, c).Value) Then
            GetOrderQtyNumberFormat = wsO.Cells(r, c).NumberFormat
            If Len(GetOrderQtyNumberFormat) > 0 Then Exit Function
        End If
    Next r
    ' Fallback to column format
    GetOrderQtyNumberFormat = wsO.Columns(c).NumberFormat
    Exit Function
Done:
    GetOrderQtyNumberFormat = vbNullString
End Function


' ========================== PRICE FETCH HELPERS ==============================
Private Function HttpGet(ByVal url As String) As String
    On Error GoTo Fail
    Dim o As Object: Set o = CreateObject("MSXML2.XMLHTTP")
    o.Open "GET", url, False
    o.setRequestHeader "Accept", "application/json"
    o.setRequestHeader "User-Agent", "ExcelVBA-Binance/1.0"
    o.send
    If o.readyState = 4 And o.Status = 200 Then
        HttpGet = CStr(o.responseText)
    Else
        HttpGet = ""
    End If
    Exit Function
Fail:
    HttpGet = ""
End Function

Private Function ParseTickerPriceFromJson(ByVal s As String) As Double
    Dim i As Long, j As Long
    i = InStr(1, s, """price""", vbTextCompare): If i = 0 Then Exit Function
    i = InStr(i, s, ":", vbTextCompare): If i = 0 Then Exit Function
    i = InStr(i, s, """", vbTextCompare): If i = 0 Then Exit Function
    j = InStr(i + 1, s, """", vbTextCompare): If j = 0 Then Exit Function
    ParseTickerPriceFromJson = Val(Mid$(s, i + 1, j - i - 1))
End Function

' Parse first kline's close from Binance klines JSON
Private Function ParseFirstKlineCloseFromJson(ByVal s As String) As Double
    Dim a As Long, b As Long, inner As String, parts() As String
    a = InStr(1, s, "[[", vbTextCompare): If a = 0 Then Exit Function
    b = InStr(a + 2, s, "]]", vbTextCompare): If b = 0 Then Exit Function
    inner = Mid$(s, a + 2, b - a - 2)
    parts = Split(inner, ",")
    If UBound(parts) >= 4 Then ParseFirstKlineCloseFromJson = Val(Replace$(Replace$(parts(4), """", ""), " ", ""))
End Function

Private Function MapCoinToBinanceSymbol(ByVal coin As String) As String
    Dim c As String: c = UCase$(Trim$(coin))
    c = Replace$(c, "/", "")
    c = Replace$(c, "-", "")
    If Right$(c, 4) = "USDT" Or Right$(c, 4) = "USDC" Or Right$(c, 4) = "BUSD" Then
        MapCoinToBinanceSymbol = c
    Else
        MapCoinToBinanceSymbol = c & "USDT"
    End If
End Function

Private Function IsStableCoin(ByVal coin As String) As Boolean
    Dim c As String: c = UCase$(Trim$(coin))
    IsStableCoin = (c = "USDT" Or c = "USDC" Or c = "BUSD" Or c = "FDUSD" Or c = "TUSD")
End Function

Private Function GetBinanceRealtimePrice(ByVal symbol As String) As Variant
    On Error GoTo Fail
    Dim s As String: s = HttpGet("https://api.binance.com/api/v3/ticker/price?symbol=" & symbol)
    If Len(s) = 0 Then GoTo Fail
    Dim p As Double: p = ParseTickerPriceFromJson(s)
    If p > 0 Then GetBinanceRealtimePrice = p Else GetBinanceRealtimePrice = Empty
    Exit Function
Fail:
    GetBinanceRealtimePrice = Empty
End Function

' Epoch milliseconds as string (prevents 32-bit overflow)
Private Function MsSinceEpochUTC(ByVal dt As Date) As String
    MsSinceEpochUTC = Format$(CDbl((dt - #1/1/1970#) * 86400000#), "0")
End Function

' Get the D1 close for a given UTC calendar day (Binance D1 is UTC-aligned).
Private Function GetBinanceDailyCloseUTC(ByVal symbol As String, ByVal dayUTC As Date) As Variant
    On Error GoTo Fail

    ' UTC window: [dayUTC 00:00, next 00:00)
    Dim startUtc As Date, endUtc As Date
    startUtc = DateSerial(Year(dayUTC), Month(dayUTC), Day(dayUTC))
    endUtc = DateSerial(Year(dayUTC), Month(dayUTC), Day(dayUTC) + 1)

    Dim startMs As String, endMs As String
    startMs = MsSinceEpochUTC(startUtc)
    endMs = MsSinceEpochUTC(endUtc)

    Dim url As String, s As String, closeP As Double

    ' Preferred: startTime + limit=1
    url = "https://api.binance.com/api/v3/klines?symbol=" & symbol & _
          "&interval=1d&startTime=" & startMs & "&limit=1"
    s = HttpGet(url)
    closeP = ParseFirstKlineCloseFromJson(s)
    If closeP > 0 Then GetBinanceDailyCloseUTC = closeP: Exit Function

    ' Fallback: endTime + limit=1
    url = "https://api.binance.com/api/v3/klines?symbol=" & symbol & _
          "&interval=1d&endTime=" & endMs & "&limit=1"
    s = HttpGet(url)
    closeP = ParseFirstKlineCloseFromJson(s)
    If closeP > 0 Then GetBinanceDailyCloseUTC = closeP: Exit Function

    GetBinanceDailyCloseUTC = Empty
    Exit Function
Fail:
    GetBinanceDailyCloseUTC = Empty
End Function

' Fallback for USDT->USDC/BUSD using same UTC-day logic
Private Function GetFallbackRealtimeOrCloseUTC(ByVal usdtSym As String, ByVal dayCutoffUTC7 As Date, ByVal todayUTC7 As Date) As Variant
    Dim base As String: base = Left$(usdtSym, Len(usdtSym) - 4)
    Dim px As Variant, qv As Variant, alt As String

    For Each qv In Array("USDC", "BUSD")
        alt = base & CStr(qv)
        If dayCutoffUTC7 < todayUTC7 Then
            Dim dayUTC As Date: dayUTC = dayCutoffUTC7   ' interpret as UTC day
            px = GetBinanceDailyCloseUTC(alt, dayUTC)
        Else
            px = GetBinanceRealtimePrice(alt)
        End If
        If IsNumeric(px) And px > 0 Then GetFallbackRealtimeOrCloseUTC = px: Exit Function
    Next qv

    GetFallbackRealtimeOrCloseUTC = Empty
End Function


' ========================= CUTOFF PARSER HELPERS =============================
' Read cutoff from Position!B3 (UTC+7). Returns True if a valid datetime was parsed.
Private Function GetCutoffFromPositionB3(ByRef outDt As Date) As Boolean
    Dim v As Variant
    v = ThisWorkbook.Worksheets(mod_config.SHEET_PORTFOLIO).Range(mod_config.CELL_CUTOFF).Value
    GetCutoffFromPositionB3 = TryParseDateTimeFlexible(v, outDt)
End Function

' Flexible datetime parser (keeps time if present).
' Accepts: yyyy-mm-dd[ hh:mm[:ss]], dd/mm/yyyy[ hh:mm], dd-mm-yyyy, 2025.08.31, Excel serial, etc.
Private Function TryParseDateTimeFlexible(ByVal v As Variant, ByRef dt As Date) As Boolean
    On Error GoTo Nope

    If IsDate(v) Then
        dt = CDate(v): TryParseDateTimeFlexible = True: Exit Function
    End If

    If IsNumeric(v) Then
        dt = CDate(v): TryParseDateTimeFlexible = True: Exit Function
    End If

    If VarType(v) = vbString Then
        Dim t As String: t = Trim$(CStr(v))
        If t = "" Then GoTo Nope

        t = Replace$(t, "\", "/")
        t = Replace$(t, "-", "/")
        t = Replace$(t, ".", "/")

        On Error Resume Next
        dt = CDate(t)
        If Err.Number = 0 Then TryParseDateTimeFlexible = True: Exit Function
        Err.Clear

        Dim a() As String: a = Split(Split(t, " ")(0), "/")
        If UBound(a) = 2 Then
            Dim yy As Long, mm As Long, dd As Long, tm As String
            tm = ""
            If InStr(t, " ") > 0 Then tm = Trim$(Mid$(t, InStr(t, " ") + 1))

            If Len(a(0)) = 4 Then
                yy = CLng(a(0)): mm = CLng(a(1)): dd = CLng(a(2))   ' yyyy/mm/dd
            Else
                yy = CLng(a(2)): mm = CLng(a(1)): dd = CLng(a(0))   ' dd/mm/yyyy
            End If

            If tm <> "" Then
                dt = CDate(Format$(DateSerial(yy, mm, dd), "yyyy-mm-dd") & " " & tm)
            Else
                dt = DateSerial(yy, mm, dd)
            End If
            TryParseDateTimeFlexible = True: Exit Function
        End If
        On Error GoTo 0
    End If

Nope:
    TryParseDateTimeFlexible = False
End Function


' ============================== MISC HELPERS =================================
Private Function NzD(v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Or v = "" Then NzD = 0# Else NzD = CDbl(v)
End Function

Private Function MinCol(d As Object) As Long
    Dim c As Long, k As Variant: c = 1000000
    For Each k In d.Keys
        If d(k) > 0 Then If d(k) < c Then c = d(k)
    Next k
    MinCol = c
End Function

Private Function MaxCol(d As Object) As Long
    Dim c As Long, k As Variant: c = 0
    For Each k In d.Keys
        If d(k) > c Then c = d(k)
    Next k
    MaxCol = c
End Function

' Excel ROUND: half away from zero
Private Function RoundN(ByVal v As Variant, ByVal n As Long) As Variant
    If IsNumeric(v) Then
        RoundN = Application.WorksheetFunction.Round(CDbl(v), n)
    Else
        RoundN = v
    End If
End Function



Public Sub Take_Daily_Snapshot()
    On Error GoTo Fail

    Dim wsP As Worksheet, wsS As Worksheet, wsC As Worksheet
    Set wsP = SheetByName(mod_config.SHEET_PORTFOLIO)
    If wsP Is Nothing Then Err.Raise 1004, , "Sheet '" & mod_config.SHEET_PORTFOLIO & "' not found."

    ' Ensure target sheet exists with correct name
    Set wsS = SheetByName(mod_config.SHEET_SNAPSHOT)
    If wsS Is Nothing Then
        Set wsS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsS.name = mod_config.SHEET_SNAPSHOT
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Snapshot date from Position!B3 (UTC+7); fallback to today if invalid
    Dim snapDt As Date
    If Not GetCutoffFromPositionB3(snapDt) Then snapDt = Date
    snapDt = DateValue(snapDt) ' only date

    ' Read dashboard totals from Position (d�ng constants d� khai b�o trong module)
    Dim cashVal As Variant, coinVal As Variant, navVal As Variant
    Dim depVal As Variant, wdrVal As Variant, pnlVal As Variant

    cashVal = wsP.Range(mod_config.CELL_CASH).Value
    coinVal = wsP.Range(mod_config.CELL_COIN).Value
    navVal = wsP.Range(mod_config.CELL_NAV).Value
    depVal = wsP.Range(mod_config.CELL_SUM_DEPOSIT).Value
    wdrVal = wsP.Range(mod_config.CELL_SUM_WITHDRAW).Value
    pnlVal = wsP.Range(mod_config.CELL_TOTAL_PNL).Value

    ' ---------- Build holdings from Position (Open rows) ----------
    Dim hdrP As Long: hdrP = DetectPortfolioHeaderRow(wsP)
    If hdrP = 0 Then Err.Raise 1004, , "Cannot find header row on '" & mod_config.SHEET_PORTFOLIO & "'."
    Dim OUT_START As Long: OUT_START = hdrP + OUTPUT_OFFSET_ROWS
    Dim portCols As Object: Set portCols = MapPortfolioHeaders(wsP, hdrP)

    Dim lastPos As Long: lastPos = LastRowIn(wsP, portCols("Coin"), hdrP)

    Dim holdDict As Object: Set holdDict = CreateObject("Scripting.Dictionary")
    holdDict.CompareMode = vbTextCompare

    Dim r As Long, posTxt As String, coin As String
    Dim availBal As Double, qty As Double, mkt As Double

    For r = OUT_START To lastPos
        posTxt = Trim$(CStr(wsP.Cells(r, portCols("Position")).Value))
        If LCase$(posTxt) = "open" Then
            coin = Trim$(CStr(wsP.Cells(r, portCols("Coin")).Value))
            If coin <> "" Then
                ' Try available balance; if blank, compute qty*price
                availBal = 0#
                If portCols("available balance") > 0 Then
                    If IsNumeric(wsP.Cells(r, portCols("available balance")).Value) Then
                        availBal = CDbl(wsP.Cells(r, portCols("available balance")).Value)
                    End If
                End If
                If availBal = 0# Then
                    qty = 0#: mkt = 0#
                    If portCols("available qty") > 0 And IsNumeric(wsP.Cells(r, portCols("available qty")).Value) Then _
                        qty = CDbl(wsP.Cells(r, portCols("available qty")).Value)
                    If portCols("market price") > 0 And IsNumeric(wsP.Cells(r, portCols("market price")).Value) Then _
                        mkt = CDbl(wsP.Cells(r, portCols("market price")).Value)
                    availBal = qty * mkt
                End If

                If availBal > 0 Then
                    If holdDict.Exists(coin) Then
                        holdDict(coin) = holdDict(coin) + availBal
                    Else
                        holdDict(coin) = availBal
                    End If
                End If
            End If
        End If
    Next r

    ' Compose "COIN: value" string with thousand separators (sorted by value desc)
    Dim holdingsStr As String: holdingsStr = ""
    If holdDict.Count > 0 Then
        Dim nHold As Long: nHold = holdDict.Count
        Dim keysArr() As Variant, valsArr() As Double
        ReDim keysArr(1 To nHold)
        ReDim valsArr(1 To nHold)
        Dim k As Variant, idxH As Long: idxH = 1
        For Each k In holdDict.Keys
            keysArr(idxH) = CStr(UCase$(k))
            valsArr(idxH) = CDbl(holdDict(k))
            idxH = idxH + 1
        Next k
        ' Simple selection sort by value desc
        Dim iH As Long, jH As Long
        For iH = 1 To nHold - 1
            Dim maxIdx As Long: maxIdx = iH
            For jH = iH + 1 To nHold
                If valsArr(jH) > valsArr(maxIdx) Then maxIdx = jH
            Next jH
            If maxIdx <> iH Then
                Dim tv As Double, ts As Variant
                tv = valsArr(iH): valsArr(iH) = valsArr(maxIdx): valsArr(maxIdx) = tv
                ts = keysArr(iH): keysArr(iH) = keysArr(maxIdx): keysArr(maxIdx) = ts
            End If
        Next iH
        For iH = 1 To nHold
            If Len(holdingsStr) > 0 Then holdingsStr = holdingsStr & "; "
            holdingsStr = holdingsStr & keysArr(iH) & ": " & Format(valsArr(iH), "#,##0")
        Next iH
    End If

    ' ---------- Build group totals (BTC/Alt.TOP/Alt.MID/Alt.LOW) ----------
    Dim coinToGroup As Object
    Set wsC = SheetByName(mod_config.SHEET_CATEGORY)
    If wsC Is Nothing Then Set wsC = SheetByName("Category")
    If wsC Is Nothing Then Set wsC = SheetByName("Categories")
    If Not wsC Is Nothing Then
        Set coinToGroup = BuildCoinToGroupFromCategorySheet(wsC)
    Else
        Set coinToGroup = CreateObject("Scripting.Dictionary"): coinToGroup.CompareMode = vbTextCompare
    End If

    Dim gVals As Object: Set gVals = CreateObject("Scripting.Dictionary"): gVals.CompareMode = vbTextCompare
    gVals("BTC") = 0#: gVals("Alt.TOP") = 0#: gVals("Alt.MID") = 0#: gVals("Alt.LOW") = 0#

    If holdDict.Count > 0 Then
        Dim ck As Variant, grp As String, val As Double
        For Each ck In holdDict.Keys
            val = CDbl(holdDict(ck))
            If UCase$(CStr(ck)) = "BTC" Then
                grp = "BTC"
            ElseIf Not (coinToGroup Is Nothing) And coinToGroup.Exists(CStr(ck)) Then
                grp = CStr(coinToGroup(ck))
            Else
                grp = "Alt.LOW"
            End If
            If gVals.Exists(grp) Then gVals(grp) = gVals(grp) + val
        Next ck
    End If

    ' ---------- Ensure headers (standardize to A:L layout) ----------
    wsS.Cells(1, 1).Value = "Date"            ' A1
    wsS.Cells(1, 2).Value = "Cash"            ' B1
    wsS.Cells(1, 3).Value = "Coin"            ' C1 (total coin value)
    wsS.Cells(1, 4).Value = "NAV"             ' D1
    wsS.Cells(1, 5).Value = "Total deposit"   ' E1
    wsS.Cells(1, 6).Value = "Total withdraw"  ' F1
    wsS.Cells(1, 7).Value = "Total profit"    ' G1
    wsS.Cells(1, 8).Value = "BTC"             ' H1
    wsS.Cells(1, 9).Value = "Alt.TOP"         ' I1
    wsS.Cells(1, 10).Value = "Alt.MID"        ' J1
    wsS.Cells(1, 11).Value = "Alt.LOW"        ' K1
    wsS.Cells(1, 12).Value = "Holdings"       ' L1 (Holdings string)
    wsS.Range("A1:L1").Font.Bold = True
    wsS.Columns("A").NumberFormat = mod_config.SNAPSHOT_DATE_FMT
    wsS.Columns("B:G").NumberFormat = mod_config.SNAPSHOT_NUMBER_FMT
    wsS.Columns("H:K").NumberFormat = mod_config.SNAPSHOT_NUMBER_FMT
    wsS.Columns("L").NumberFormat = "@"       ' text

    ' ---------- UPSERT: t�m d�ng c� Date = snapDt ----------
    Dim lastRow As Long, writeRow As Long, found As Boolean
    lastRow = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
    found = False
    If lastRow >= 2 Then
        For r = 2 To lastRow
            If IsDate(wsS.Cells(r, 1).Value) Then
                If DateValue(wsS.Cells(r, 1).Value) = snapDt Then
                    writeRow = r
                    found = True
                    Exit For
                End If
            End If
        Next r
    End If
    If Not found Then writeRow = lastRow + 1

    ' ---------- Write snapshot row ----------
    ' Clear group/holdings cells for this row to avoid stale text from older layout
    wsS.Range(wsS.Cells(writeRow, 8), wsS.Cells(writeRow, 12)).ClearContents
    wsS.Cells(writeRow, 1).Value = snapDt
    wsS.Cells(writeRow, 2).Value = Round(cashVal, 0)
    wsS.Cells(writeRow, 3).Value = Round(coinVal, 0)
    wsS.Cells(writeRow, 4).Value = Round(navVal, 0)
    wsS.Cells(writeRow, 5).Value = Round(depVal, 0)
    wsS.Cells(writeRow, 6).Value = Round(wdrVal, 0)
    wsS.Cells(writeRow, 7).Value = Round(pnlVal, 0)
    wsS.Cells(writeRow, 8).Value = Round(NzD(gVals("BTC")), 0)
    wsS.Cells(writeRow, 9).Value = Round(NzD(gVals("Alt.TOP")), 0)
    wsS.Cells(writeRow, 10).Value = Round(NzD(gVals("Alt.MID")), 0)
    wsS.Cells(writeRow, 11).Value = Round(NzD(gVals("Alt.LOW")), 0)
    wsS.Cells(writeRow, 12).Value = holdingsStr

    ' ---------- Sort by Date ascending ----------
    lastRow = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
    If lastRow > 2 Then
        wsS.Sort.SortFields.Clear
        wsS.Sort.SortFields.Add key:=wsS.Range("A2:A" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With wsS.Sort
            .SetRange wsS.Range("A1:L" & lastRow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If

    wsS.Columns("A:L").AutoFit
    MsgBox "Daily snapshot saved for " & Format$(snapDt, "yyyy-mm-dd"), vbInformation

Clean:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error (Take_Daily_Snapshot): " & Err.Description, vbExclamation
End Sub

' ============================= SNAPSHOT HELPER ===============================
 ' Removed unused SafeRead helper (no references)

Private Function GetRealtimePriceByExchange(ByVal exchangeName As String, ByVal coin As String) As Variant
    Dim ex As String: ex = LCase$(Trim$(exchangeName))
    If ex = "" Or ex = "binance" Then
        GetRealtimePriceByExchange = GetBinanceRealtimePrice(MapCoinToBinanceSymbol(coin))
        Exit Function
    End If
    If ex = "okx" Or ex = "okex" Then
        GetRealtimePriceByExchange = GetOkxRealtimePrice(MapCoinToOkxInstId(coin))
        Exit Function
    End If
    If ex = "bybit" Then
        GetRealtimePriceByExchange = GetBybitRealtimePrice(MapCoinToBybitSymbol(coin))
        Exit Function
    End If
    ' Unknown exchange -> try Binance
    GetRealtimePriceByExchange = GetBinanceRealtimePrice(MapCoinToBinanceSymbol(coin))
End Function

Private Function MapCoinToOkxInstId(ByVal coin As String) As String
    Dim c As String: c = UCase$(Trim$(coin))
    c = Replace$(c, "/", "")
    c = Replace$(c, "-", "")
    If Right$(c, 4) = "USDT" Then
        MapCoinToOkxInstId = Left$(c, Len(c) - 4) & "-USDT"
    Else
        MapCoinToOkxInstId = c & "-USDT"
    End If
End Function

Private Function MapCoinToBybitSymbol(ByVal coin As String) As String
    MapCoinToBybitSymbol = MapCoinToBinanceSymbol(coin) ' same format: BTCUSDT
End Function

Private Function GetOkxRealtimePrice(ByVal instId As String) As Variant
    On Error GoTo Fail
    Dim url As String: url = "https://www.okx.com/api/v5/market/ticker?instId=" & instId
    Dim s As String: s = HttpGet(url)
    If Len(s) = 0 Then GoTo Fail
    ' Find "last":"<price>"
    Dim i As Long, j As Long
    i = InStr(1, s, "last", vbTextCompare): If i = 0 Then GoTo Fail
    i = InStr(i, s, ":", vbTextCompare): If i = 0 Then GoTo Fail
    i = InStr(i, s, """", vbTextCompare): If i = 0 Then GoTo Fail
    j = InStr(i + 1, s, """", vbTextCompare): If j = 0 Then GoTo Fail
    GetOkxRealtimePrice = Val(Mid$(s, i + 1, j - i - 1))
    Exit Function
Fail:
    GetOkxRealtimePrice = Empty
End Function

Private Function GetBybitRealtimePrice(ByVal symbol As String) As Variant
    On Error GoTo Fail
    Dim url As String
    url = "https://api.bybit.com/v5/market/tickers?category=spot&symbol=" & symbol
    Dim s As String: s = HttpGet(url)
    If Len(s) = 0 Then GoTo Fail
    ' Find "lastPrice":"<price>"
    Dim i As Long, j As Long
    i = InStr(1, s, "lastPrice", vbTextCompare): If i = 0 Then GoTo Fail
    i = InStr(i, s, ":", vbTextCompare): If i = 0 Then GoTo Fail
    i = InStr(i, s, """", vbTextCompare): If i = 0 Then GoTo Fail
    j = InStr(i + 1, s, """", vbTextCompare): If j = 0 Then GoTo Fail
    GetBybitRealtimePrice = Val(Mid$(s, i + 1, j - i - 1))
    Exit Function
Fail:
    GetBybitRealtimePrice = Empty
End Function

Private Function NzStr(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

' =============================================================================
' ========================= BATCH SNAPSHOT UPDATER ============================
' =============================================================================
' Update all missing daily snapshots from Daily_Snapshot!A2 (start date)
' up to the current cutoff date on Position!B3. Creates rows only for
' missing dates; existing dates are left unchanged.
Public Sub Update_All_Snapshot()
    On Error GoTo Fail

    Dim wsP As Worksheet, wsS As Worksheet
    Set wsP = SheetByName(mod_config.SHEET_PORTFOLIO)
    Set wsS = SheetByName(mod_config.SHEET_SNAPSHOT)
    If wsP Is Nothing Then Err.Raise 1004, , "Sheet '" & mod_config.SHEET_PORTFOLIO & "' not found."
    If wsS Is Nothing Then Err.Raise 1004, , "Sheet '" & mod_config.SHEET_SNAPSHOT & "' not found."

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim cutoff As Date
    If Not GetCutoffFromPositionB3(cutoff) Then Err.Raise 1004, , "Invalid cutoff at Position!" & mod_config.CELL_CUTOFF
    cutoff = DateValue(cutoff)

    Dim startDate As Date
    If IsDate(wsS.Cells(2, 1).Value) Then
        startDate = DateValue(wsS.Cells(2, 1).Value)
    Else
        Err.Raise 1004, , "Daily_Snapshot!A2 must contain a start date."
    End If
    If startDate > cutoff Then GoTo Clean ' nothing to do

    ' Build a set of existing dates to avoid overwriting
    Dim exists As Object: Set exists = CreateObject("Scripting.Dictionary")
    exists.CompareMode = vbTextCompare
    Dim lastRow As Long: lastRow = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If IsDate(wsS.Cells(r, 1).Value) Then exists(AddDateKey(DateValue(wsS.Cells(r, 1).Value))) = True
    Next r

    Dim originalCutoff As Variant
    originalCutoff = wsP.Range(mod_config.CELL_CUTOFF).Value

    Dim d As Date
    gSuppressPositionMsg = True
    For d = startDate To cutoff
        If Not exists.Exists(AddDateKey(d)) Then
            ' Rebuild Position for date d by setting cutoff and running full update
            wsP.Range(mod_config.CELL_CUTOFF).Value = d
            Update_All_Position
            ' Take snapshot for this date
            Take_Daily_Snapshot
        End If
    Next d
    gSuppressPositionMsg = False

    ' Restore original cutoff
    wsP.Range(mod_config.CELL_CUTOFF).Value = originalCutoff

Clean:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    gSuppressPositionMsg = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error (Update_All_Snapshot): " & Err.Description, vbExclamation
End Sub

Private Function AddDateKey(ByVal d As Date) As String
    AddDateKey = Format$(d, "yyyy-mm-dd")
End Function

