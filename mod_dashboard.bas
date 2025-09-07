Attribute VB_Name = "mod_dashboard"
Option Explicit
' Last Modified (UTC): 2025-09-07T08:15:56Z

Public Sub Update_Dashboard()
    On Error GoTo Fail

    Dim wsDash As Worksheet, wsSnap As Worksheet
    Set wsDash = SheetByName(mod_config.SHEET_DASHBOARD)
    If wsDash Is Nothing Then Err.Raise 1004, , "Sheet '" & mod_config.SHEET_DASHBOARD & "' not found."
    Set wsSnap = SheetByName(mod_config.SHEET_SNAPSHOT)
    If wsSnap Is Nothing Then Err.Raise 1004, , "Sheet '" & mod_config.SHEET_SNAPSHOT & "' not found."

    Dim dStart As Date, dEnd As Date
    If Not IsDate(wsDash.Range("B2").Value) Then Err.Raise 1004, , mod_config.SHEET_DASHBOARD & "!B2 must be a date (start)."
    If Not IsDate(wsDash.Range("B3").Value) Then Err.Raise 1004, , mod_config.SHEET_DASHBOARD & "!B3 must be a date (end)."
    dStart = DateValue(wsDash.Range("B2").Value)
    dEnd = DateValue(wsDash.Range("B3").Value)
    If dEnd < dStart Then Err.Raise 1004, , "Dashboard date range invalid (end < start)."

    ' Performance toggles
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalc As XlCalculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Build NAV series from Daily_Snapshot (Date in col A, NAV in col D)
    Dim lastR As Long: lastR = wsSnap.Cells(wsSnap.Rows.Count, 1).End(xlUp).Row
    Dim datesArr() As Variant, navArr() As Double, pnlArr() As Double, depArr() As Double, wdrArr() As Double, ratioArr() As Double
    Dim tv As Double, ts As String
    Dim count As Long: count = 0

    Dim r As Long, dt As Date
    ' Reused header temp variables across multiple loops
    Dim hname As String, norm As String
    ' Pass 1: count rows in range
    For r = 2 To lastR
        If IsDate(wsSnap.Cells(r, 1).Value) Then
            dt = DateValue(wsSnap.Cells(r, 1).Value)
            If dt >= dStart And dt <= dEnd Then count = count + 1
        End If
    Next r

    If count > 0 Then
        ReDim datesArr(1 To count)
        ReDim navArr(1 To count)
        ReDim pnlArr(1 To count)
        ReDim depArr(1 To count)
        ReDim wdrArr(1 To count)
        ReDim ratioArr(1 To count)
    End If

    ' Pass 2: fill arrays
    Dim idx As Long: idx = 0
    For r = 2 To lastR
        If IsDate(wsSnap.Cells(r, 1).Value) Then
            dt = DateValue(wsSnap.Cells(r, 1).Value)
            If dt >= dStart And dt <= dEnd Then
                idx = idx + 1
                datesArr(idx) = dt
                If IsNumeric(wsSnap.Cells(r, 4).Value) Then navArr(idx) = CDbl(wsSnap.Cells(r, 4).Value) Else navArr(idx) = 0#
                If IsNumeric(wsSnap.Cells(r, 7).Value) Then pnlArr(idx) = CDbl(wsSnap.Cells(r, 7).Value) Else pnlArr(idx) = 0#
                If IsNumeric(wsSnap.Cells(r, 5).Value) Then depArr(idx) = CDbl(wsSnap.Cells(r, 5).Value) Else depArr(idx) = 0#
                If IsNumeric(wsSnap.Cells(r, 6).Value) Then wdrArr(idx) = CDbl(wsSnap.Cells(r, 6).Value) Else wdrArr(idx) = 0#
                Dim cashVal As Double
                If IsNumeric(wsSnap.Cells(r, 2).Value) Then cashVal = CDbl(wsSnap.Cells(r, 2).Value) Else cashVal = 0#
                If Abs(navArr(idx)) > mod_config.EPS_CLOSE Then ratioArr(idx) = cashVal / navArr(idx) Else ratioArr(idx) = 0#
            End If
        End If
    Next r

    Dim co As ChartObject, ch As Chart
    Set co = GetOrCreateChart(wsDash, "NAV")
    Set ch = co.Chart

    ' Apply series
    Dim s As Series
    If ch.SeriesCollection.Count = 0 Then
        Set s = ch.SeriesCollection.NewSeries
    Else
        Set s = ch.SeriesCollection(1)
    End If

    If count > 0 Then
        s.Values = navArr
        s.XValues = datesArr
    Else
        ' No data in range: clear series
        On Error Resume Next
        ch.SeriesCollection(1).Delete
        On Error GoTo 0
    End If

    ch.ChartType = xlLine
    ch.HasTitle = True
    ch.ChartTitle.Text = "NAV"
    On Error Resume Next
    ch.Axes(xlCategory).CategoryType = xlTimeScale
    ch.Axes(xlCategory).TickLabels.NumberFormat = mod_config.SNAPSHOT_DATE_FMT
    ch.Axes(xlValue).TickLabels.NumberFormat = mod_config.MONEY_FMT
    On Error GoTo 0

    ' Compute drawdowns and annotate on NAV chart (before switching charts)
    Dim mddPct As Double: mddPct = ComputeMaxDrawdownPct(navArr)
    Dim curDDPct As Double: curDDPct = ComputeCurrentDrawdownPct(navArr)
    AnnotateDrawdown ch, mddPct, curDDPct

    ' Build/Apply combined Deposit & Withdraw chart
    Set co = GetOrCreateChart(wsDash, "Deposit")
    Set ch = co.Chart
    Dim sDep As Series, sWdr As Series
    If ch.SeriesCollection.Count = 0 Then
        Set sDep = ch.SeriesCollection.NewSeries
        Set sWdr = ch.SeriesCollection.NewSeries
    ElseIf ch.SeriesCollection.Count = 1 Then
        Set sDep = ch.SeriesCollection(1)
        Set sWdr = ch.SeriesCollection.NewSeries
    Else
        Set sDep = ch.SeriesCollection(1)
        Set sWdr = ch.SeriesCollection(2)
    End If
    If count > 0 Then
        sDep.Name = "Deposit"
        sDep.Values = depArr
        sDep.XValues = datesArr
        sWdr.Name = "Withdraw"
        sWdr.Values = wdrArr
        sWdr.XValues = datesArr
    Else
        On Error Resume Next
        Do While ch.SeriesCollection.Count > 0
            ch.SeriesCollection(1).Delete
        Loop
        On Error GoTo 0
    End If
    ch.ChartType = xlLine
    ch.HasTitle = True
    ch.ChartTitle.Text = "Deposit & Withdraw"
    ' Ensure no drawdown annotation appears on this chart
    On Error Resume Next
    ch.Shapes("MDD_NAV").Delete
    On Error GoTo 0
    ch.HasLegend = True
    On Error Resume Next
    ch.Axes(xlCategory).CategoryType = xlTimeScale
    ch.Axes(xlCategory).TickLabels.NumberFormat = mod_config.SNAPSHOT_DATE_FMT
    ch.Axes(xlValue).TickLabels.NumberFormat = mod_config.MONEY_FMT
    With ch.Axes(xlValue)
        .Crosses = xlAxisCrossesMinimum
    End With
    ' Remove old standalone Withdraw chart if it exists
    On Error Resume Next
    wsDash.ChartObjects("Withdraw").Delete
    On Error GoTo 0

    ' Build/Apply PnL chart (Total profit)
    Set co = GetOrCreateChart(wsDash, "PnL")
    Set ch = co.Chart
    If ch.SeriesCollection.Count = 0 Then
        Set s = ch.SeriesCollection.NewSeries
    Else
        Set s = ch.SeriesCollection(1)
    End If
    If count > 0 Then
        s.Values = pnlArr
        s.XValues = datesArr
    Else
        On Error Resume Next
        ch.SeriesCollection(1).Delete
        On Error GoTo 0
    End If
    ch.ChartType = xlLine
    ch.HasTitle = True
    ch.ChartTitle.Text = "PnL"
    On Error Resume Next
    ch.Axes(xlCategory).CategoryType = xlTimeScale
    ch.Axes(xlCategory).TickLabels.NumberFormat = mod_config.SNAPSHOT_DATE_FMT
    ch.Axes(xlValue).TickLabels.NumberFormat = mod_config.MONEY_FMT
    ' Always place X axis at bottom
    With ch.Axes(xlValue)
        .Crosses = xlAxisCrossesMinimum
    End With
    On Error GoTo 0

    ' Do not annotate drawdown on PnL chart

    ' Build/Apply Cash vs NAV chart (percent = Cash/NAV)
    ' Remove legacy chart if previously named "Cash vs Coin"
    On Error Resume Next: wsDash.ChartObjects("Cash vs Coin").Delete: On Error GoTo 0
    Set co = GetOrCreateChart(wsDash, "Cash vs NAV")
    Set ch = co.Chart
    If ch.SeriesCollection.Count = 0 Then
        Set s = ch.SeriesCollection.NewSeries
    Else
        Set s = ch.SeriesCollection(1)
    End If
    If count > 0 Then
        s.Name = "Cash/NAV"
        s.Values = ratioArr
        s.XValues = datesArr
    Else
        On Error Resume Next
        ch.SeriesCollection(1).Delete
        On Error GoTo 0
    End If
    ch.ChartType = xlLine
    ch.HasTitle = True
    ch.ChartTitle.Text = "Cash vs NAV"
    ch.HasLegend = False
    On Error Resume Next
    ch.Axes(xlCategory).CategoryType = xlTimeScale
    ch.Axes(xlCategory).TickLabels.NumberFormat = mod_config.SNAPSHOT_DATE_FMT
    ' Show percentage (e.g., 10%, 21%)
    ch.Axes(xlValue).TickLabels.NumberFormat = "0%"
    With ch.Axes(xlValue)
        .Crosses = xlAxisCrossesMinimum
    End With
    On Error GoTo 0

    ' Build/Apply Portfolio_Catagory chart (stacked column of category amounts; no labels)
    ' Assumed Daily_Snapshot layout:
    '   Col A = Date, B = Cash, C = Coin (total), D = NAV, E = Deposit, F = Withdraw, G = PnL,
    '   Col H.. = one column per Category (amount in same units as Coin, e.g., USD value of holdings)
    Dim lastCol As Long
    lastCol = wsSnap.Cells(1, wsSnap.Columns.Count).End(xlToLeft).Column
    If lastCol > 7 Then
        ' Prefer computing Portfolio_Catagory from Holdings + Catagory mapping if available
        Dim holdColPG As Long: holdColPG = 0
        Dim cPG As Long, hdrPG As String
        For cPG = 1 To lastCol
            hdrPG = CStr(wsSnap.Cells(1, cPG).Value)
            If InStr(1, hdrPG, "holding", vbTextCompare) > 0 Then
                holdColPG = cPG: Exit For
            End If
        Next cPG
        If holdColPG > 0 And count > 0 Then
            Dim wsCat As Worksheet: Set wsCat = SheetByName(mod_config.SHEET_CATEGORY)
            If wsCat Is Nothing Then Set wsCat = SheetByName("Categoty")
            If wsCat Is Nothing Then Set wsCat = SheetByName("Catagory")
            If wsCat Is Nothing Then Set wsCat = SheetByName("Category")
            Dim mapCG As Object: Set mapCG = BuildCoinToGroupMap(wsCat)

            Dim grpVals As Object: Set grpVals = CreateObject("Scripting.Dictionary"): grpVals.CompareMode = vbTextCompare
            Dim grpNamesDyn As Object: Set grpNamesDyn = CreateObject("Scripting.Dictionary"): grpNamesDyn.CompareMode = vbTextCompare

            Dim idxPG As Long: idxPG = 0
            Dim holdsPG As Object, ckPG As Variant
            Dim coinNamePG As String, amtPG As Double, grpPG As String, ovrPG As String
            For r = 2 To lastR
                If IsDate(wsSnap.Cells(r, 1).Value) Then
                    dt = DateValue(wsSnap.Cells(r, 1).Value)
                    If dt >= dStart And dt <= dEnd Then
                        idxPG = idxPG + 1
                        Set holdsPG = ParseHoldingsString(CStr(wsSnap.Cells(r, holdColPG).Value))
                        For Each ckPG In holdsPG.Keys
                            coinNamePG = CStr(ckPG)
                            amtPG = CDbl(holdsPG(ckPG))
                            grpPG = ""
                            ovrPG = GroupOverride(coinNamePG)
                            If Len(ovrPG) > 0 Then
                                grpPG = ovrPG
                            ElseIf UCase$(coinNamePG) = "BTC" Then
                                grpPG = "BTC"
                            ElseIf Not (mapCG Is Nothing) And mapCG.Exists(coinNamePG) Then
                                grpPG = CStr(mapCG(coinNamePG))
                            Else
                                grpPG = "Other"
                            End If
                            If Not grpVals.Exists(grpPG) Then
                                Dim arrPG() As Double
                                ReDim arrPG(1 To count)
                                grpVals(grpPG) = arrPG
                                grpNamesDyn(grpPG) = 1
                            End If
                            Dim curArrPG() As Double
                            curArrPG = grpVals(grpPG)
                            curArrPG(idxPG) = curArrPG(idxPG) + amtPG
                            grpVals(grpPG) = curArrPG
                        Next ckPG
                    End If
                End If
            Next r

            ' Build a stable order: BTC, Alt.TOP, Alt.MID, Alt.LOW, then others A..Z
            Dim namesPG() As String, nPG As Long: nPG = grpNamesDyn.Count
            If nPG > 0 Then ReDim namesPG(1 To nPG)
            Dim iiPG As Long: iiPG = 0
            Dim kPG As Variant
            For Each kPG In grpNamesDyn.Keys
                iiPG = iiPG + 1: namesPG(iiPG) = CStr(kPG)
            Next kPG
            ' Simple selection sort A..Z
            Dim iPG As Long, jPG As Long, minIdxPG As Long, tmpS As String
            For iPG = 1 To nPG - 1
                minIdxPG = iPG
                For jPG = iPG + 1 To nPG
                    If StrComp(namesPG(jPG), namesPG(minIdxPG), vbTextCompare) < 0 Then minIdxPG = jPG
                Next jPG
                If minIdxPG <> iPG Then tmpS = namesPG(iPG): namesPG(iPG) = namesPG(minIdxPG): namesPG(minIdxPG) = tmpS
            Next iPG
            ' Move priority groups to front in desired order
            Dim desired As Variant: desired = Array("BTC", "Alt.TOP", "Alt.MID", "Alt.LOW")
            Dim outNames() As String: ReDim outNames(1 To nPG)
            Dim outPos As Long: outPos = 0
            For Each kPG In desired
                For iPG = 1 To nPG
                    If StrComp(namesPG(iPG), CStr(kPG), vbTextCompare) = 0 Then
                        outPos = outPos + 1: outNames(outPos) = namesPG(iPG): namesPG(iPG) = "\u0000"
                    End If
                Next iPG
            Next kPG
            For iPG = 1 To nPG
                If namesPG(iPG) <> "\u0000" Then outPos = outPos + 1: outNames(outPos) = namesPG(iPG)
            Next iPG

            ' Apply to chart
            Set co = GetOrCreateChart(wsDash, "Portfolio_Catagory")
            Set ch = co.Chart
            ch.ChartType = xlColumnStacked
            ch.HasTitle = True
            ch.ChartTitle.Text = "Portfolio_Catagory"
            ch.HasLegend = True

            ' Ensure series count equals groups
            On Error Resume Next
            Do While ch.SeriesCollection.Count > nPG: ch.SeriesCollection(ch.SeriesCollection.Count).Delete: Loop
            For iPG = ch.SeriesCollection.Count + 1 To nPG: ch.SeriesCollection.NewSeries: Next iPG
            On Error GoTo 0

            For iPG = 1 To nPG
                Set s = ch.SeriesCollection(iPG)
                s.Name = outNames(iPG)
                s.Values = grpVals(outNames(iPG))
                s.XValues = datesArr
                s.HasDataLabels = False
            Next iPG

            On Error Resume Next
            ch.Axes(xlCategory).CategoryType = xlTimeScale
            ch.Axes(xlCategory).TickLabels.NumberFormat = mod_config.SNAPSHOT_DATE_FMT
            ch.Axes(xlValue).TickLabels.NumberFormat = mod_config.MONEY_FMT
            On Error GoTo 0

        Else
            Dim catCount As Long: catCount = 0
            Dim catNames() As String, catCols() As Long
            Dim iCat As Long, iCol As Long
            Dim hdr As String
            ' Build list of category columns, skipping any header named "Holdings"/"Holding"
            For iCol = 8 To lastCol
                hdr = Trim$(CStr(wsSnap.Cells(1, iCol).Value))
                If LCase$(hdr) <> "holdings" And LCase$(hdr) <> "holding" Then
                    catCount = catCount + 1
                    ReDim Preserve catNames(1 To catCount)
                    ReDim Preserve catCols(1 To catCount)
                    If hdr = vbNullString Then
                        catNames(catCount) = "Cat" & CStr(catCount)
                    Else
                        catNames(catCount) = hdr
                    End If
                    catCols(catCount) = iCol
                End If
            Next iCol

        If count > 0 And catCount > 0 Then
            ' Prepare per-category arrays of values and percents sized to filtered date count
            Dim catVals As Object: Set catVals = CreateObject("Scripting.Dictionary")
            Dim catPct As Object:  Set catPct  = CreateObject("Scripting.Dictionary")
            catVals.CompareMode = vbTextCompare: catPct.CompareMode = vbTextCompare
            Dim tmp() As Double, tmp2() As Double
            For iCat = 1 To catCount
                ReDim tmp(1 To count)
                ReDim tmp2(1 To count)
                catVals(CStr(iCat)) = tmp
                catPct(CStr(iCat)) = tmp2
            Next iCat

            ' Second pass to fill arrays aligned with datesArr
            idx = 0
            For r = 2 To lastR
                If IsDate(wsSnap.Cells(r, 1).Value) Then
                    dt = DateValue(wsSnap.Cells(r, 1).Value)
                    If dt >= dStart And dt <= dEnd Then
                        idx = idx + 1
                        Dim coinTotal As Double
                        If IsNumeric(wsSnap.Cells(r, 3).Value) Then coinTotal = CDbl(wsSnap.Cells(r, 3).Value) Else coinTotal = 0#
                        For iCat = 1 To catCount
                            tmp = catVals(CStr(iCat))
                            tmp2 = catPct(CStr(iCat))
                            Dim vAmt As Double
                            If IsNumeric(wsSnap.Cells(r, catCols(iCat)).Value) Then vAmt = CDbl(wsSnap.Cells(r, catCols(iCat)).Value) Else vAmt = 0#
                            tmp(idx) = vAmt
                            If Abs(coinTotal) > mod_config.EPS_CLOSE Then
                                tmp2(idx) = vAmt / coinTotal
                            Else
                                tmp2(idx) = 0#
                            End If
                            catVals(CStr(iCat)) = tmp
                            catPct(CStr(iCat)) = tmp2
                        Next iCat
                    End If
                End If
            Next r

            ' Create/update stacked column chart
            Set co = GetOrCreateChart(wsDash, "Portfolio_Catagory")
            Set ch = co.Chart
            ch.ChartType = xlColumnStacked
            ch.HasTitle = True
            ch.ChartTitle.Text = "Portfolio_Catagory"
            ch.HasLegend = True
            ' Remove any pre-existing series named like "Holdings"
            On Error Resume Next
            For iCol = ch.SeriesCollection.Count To 1 Step -1
                If InStr(1, ch.SeriesCollection(iCol).Name, "holding", vbTextCompare) > 0 Then
                    ch.SeriesCollection(iCol).Delete
                End If
            Next iCol
            On Error GoTo 0
            ' Ensure enough series
            Dim sc As Long: sc = ch.SeriesCollection.Count
            If sc < catCount Then
                For iCat = sc + 1 To catCount
                    ch.SeriesCollection.NewSeries
                Next iCat
            End If
            ' Remove extra series
            On Error Resume Next
            Do While ch.SeriesCollection.Count > catCount
                ch.SeriesCollection(catCount + 1).Delete
            Loop
            On Error GoTo 0

            ' Assign data per category
            Dim sCat As Series
            For iCat = 1 To catCount
                Set sCat = ch.SeriesCollection(iCat)
                sCat.Name = catNames(iCat)
                sCat.Values = catVals(CStr(iCat))
                sCat.XValues = datesArr
                ' Data labels: show amount and percent
                ' No data labels (remove amount/percent)
                sCat.HasDataLabels = False
            Next iCat

            ' Remove legend entry named "Holdings" (case-insensitive), if present
            On Error Resume Next
            If ch.HasLegend Then
                Dim iLE As Long
                For iLE = ch.Legend.LegendEntries.Count To 1 Step -1
                    Dim cap As String
                    cap = ch.Legend.LegendEntries(iLE).Caption
                    If InStr(1, cap, "holding", vbTextCompare) > 0 Then
                        ch.Legend.LegendEntries(iLE).Delete
                    End If
                Next iLE
            End If
            On Error GoTo 0

            On Error Resume Next
            ch.Axes(xlCategory).TickLabels.NumberFormat = mod_config.SNAPSHOT_DATE_FMT
            ch.Axes(xlValue).TickLabels.NumberFormat = mod_config.MONEY_FMT
            On Error GoTo 0
        Else
            ' No rows in range: clear chart
            Set co = GetOrCreateChart(wsDash, "Portfolio_Catagory")
            Set ch = co.Chart
            On Error Resume Next
            Do While ch.SeriesCollection.Count > 0
                ch.SeriesCollection(1).Delete
            Loop
            On Error GoTo 0
            ch.ChartType = xlColumnStacked
            ch.HasTitle = True
            ch.ChartTitle.Text = "Portfolio_Catagory"
        End If
        End If
    Else
        ' No category columns present in Daily_Snapshot
        On Error Resume Next
        Set co = GetOrCreateChart(wsDash, "Portfolio_Catagory")
        co.Chart.ChartTitle.Text = "Portfolio_Catagory (no categories)"
        On Error GoTo 0
    End If

    ' Build/Apply Portfolio_Alt.TOP chart (percent breakdown by coin within Alt.TOP)
    Dim altTopCol As Long: altTopCol = 0
    Dim holdCol As Long: holdCol = 0
    Dim c As Long
    For c = 1 To lastCol
        hname = LCase$(Trim$(CStr(wsSnap.Cells(1, c).Value)))
        norm = Replace(Replace(hname, ".", ""), " ", "")
        If c >= 8 And (norm = "alttop") Then altTopCol = c
        If InStr(1, hname, "holding", vbTextCompare) > 0 Then holdCol = c
    Next c
    Set co = GetOrCreateChart(wsDash, "Portfolio_Alt.TOP")
    Set ch = co.Chart
    If altTopCol > 0 And count > 0 Then
        ' Build coin -> amount and percent arrays per date from Holdings using Category sheet mapping
        Dim wsC As Worksheet: Set wsC = SheetByName(mod_config.SHEET_CATEGORY)
        If wsC Is Nothing Then Set wsC = SheetByName("Categoty")
        If wsC Is Nothing Then Set wsC = SheetByName("Catagory")
        If wsC Is Nothing Then Set wsC = SheetByName("Category")
        Dim coinToGroup As Object: Set coinToGroup = BuildCoinToGroupMap(wsC)

        Dim coinPct As Object: Set coinPct = CreateObject("Scripting.Dictionary"): coinPct.CompareMode = vbTextCompare
        Dim coinAmt As Object: Set coinAmt = CreateObject("Scripting.Dictionary"): coinAmt.CompareMode = vbTextCompare
        Dim idx2 As Long: idx2 = 0
        Dim altTopTotal As Double, holds As Object
        For r = 2 To lastR
            If IsDate(wsSnap.Cells(r, 1).Value) Then
                dt = DateValue(wsSnap.Cells(r, 1).Value)
                If dt >= dStart And dt <= dEnd Then
                    idx2 = idx2 + 1
                    If IsNumeric(wsSnap.Cells(r, altTopCol).Value) Then altTopTotal = CDbl(wsSnap.Cells(r, altTopCol).Value) Else altTopTotal = 0#
                    If holdCol > 0 Then
                        Set holds = ParseHoldingsString(CStr(wsSnap.Cells(r, holdCol).Value))
                    Else
                        Set holds = CreateObject("Scripting.Dictionary"): holds.CompareMode = vbTextCompare
                    End If
                    Dim ck As Variant, coinName As String, amt As Double, arr() As Double, arrAmt() As Double
                    For Each ck In holds.Keys
                        coinName = CStr(ck)
                        amt = CDbl(holds(ck))
                        Dim grp As String: grp = ""
                        Dim ovr As String: ovr = GroupOverride(coinName)
                        If Len(ovr) > 0 Then
                            grp = ovr
                        ElseIf UCase$(coinName) = "BTC" Then
                            grp = "BTC"
                        ElseIf Not (coinToGroup Is Nothing) And coinToGroup.Exists(coinName) Then
                            grp = CStr(coinToGroup(coinName))
                        End If
                        If StrComp(grp, "Alt.TOP", vbTextCompare) = 0 Then
                            If Not coinPct.Exists(coinName) Then
                                ReDim arr(1 To count)
                                coinPct(coinName) = arr
                            End If
                            If Not coinAmt.Exists(coinName) Then
                                ReDim arrAmt(1 To count)
                                coinAmt(coinName) = arrAmt
                            End If
                            arr = coinPct(coinName)
                            arrAmt = coinAmt(coinName)
                            arrAmt(idx2) = amt
                            If Abs(altTopTotal) > mod_config.EPS_CLOSE Then
                                arr(idx2) = amt / altTopTotal
                            Else
                                arr(idx2) = 0#
                            End If
                            coinPct(coinName) = arr
                            coinAmt(coinName) = arrAmt
                        End If
                    Next ck
                End If
            End If
        Next r

        ' Fallback: if no Alt.TOP coins found via Category mapping, distribute Alt.TOP across non-BTC holdings proportionally
        If coinPct.Count = 0 Then
            Dim fbPct As Object: Set fbPct = CreateObject("Scripting.Dictionary"): fbPct.CompareMode = vbTextCompare
            Dim fbAmt As Object: Set fbAmt = CreateObject("Scripting.Dictionary"): fbAmt.CompareMode = vbTextCompare
            idx2 = 0
            For r = 2 To lastR
                If IsDate(wsSnap.Cells(r, 1).Value) Then
                    dt = DateValue(wsSnap.Cells(r, 1).Value)
                    If dt >= dStart And dt <= dEnd Then
                        idx2 = idx2 + 1
                        If IsNumeric(wsSnap.Cells(r, altTopCol).Value) Then altTopTotal = CDbl(wsSnap.Cells(r, altTopCol).Value) Else altTopTotal = 0#
                        If holdCol > 0 Then
                            Set holds = ParseHoldingsString(CStr(wsSnap.Cells(r, holdCol).Value))
                        Else
                            Set holds = CreateObject("Scripting.Dictionary"): holds.CompareMode = vbTextCompare
                        End If
                        Dim sumNonBTC As Double: sumNonBTC = 0#
                        For Each ck In holds.Keys
                            If StrComp(CStr(ck), "BTC", vbTextCompare) <> 0 Then sumNonBTC = sumNonBTC + CDbl(holds(ck))
                        Next ck
                        If sumNonBTC > mod_config.EPS_CLOSE And altTopTotal > mod_config.EPS_CLOSE Then
                            For Each ck In holds.Keys
                                coinName = CStr(ck)
                                If StrComp(coinName, "BTC", vbTextCompare) <> 0 Then
                                    Dim arrP() As Double, arrA() As Double
                                    If Not fbPct.Exists(coinName) Then ReDim arrP(1 To count): fbPct(coinName) = arrP
                                    If Not fbAmt.Exists(coinName) Then ReDim arrA(1 To count): fbAmt(coinName) = arrA
                                    arrP = fbPct(coinName): arrA = fbAmt(coinName)
                                    Dim frac As Double: frac = CDbl(holds(ck)) / sumNonBTC
                                    arrP(idx2) = frac
                                    arrA(idx2) = frac * altTopTotal
                                    fbPct(coinName) = arrP: fbAmt(coinName) = arrA
                                End If
                            Next ck
                        End If
                    End If
                End If
            Next r
            Set coinPct = fbPct
            Set coinAmt = fbAmt
        End If

        ' Group tiny slices (<1% avg) into Other and sort by average share
        Dim thresh As Double: thresh = 0.01
        Dim nCoins As Long: nCoins = coinPct.Count
        Dim names() As String, avgs() As Double
        ReDim names(1 To IIf(nCoins = 0, 1, nCoins))
        ReDim avgs(1 To IIf(nCoins = 0, 1, nCoins))
        Dim nKeys As Long: nKeys = 0
        Dim k As Variant, i As Long
        For Each k In coinPct.Keys
            nKeys = nKeys + 1
            names(nKeys) = CStr(k)
            Dim pArr() As Double: pArr = coinPct(k)
            Dim sumP As Double: sumP = 0#: Dim cntP As Long: cntP = 0
            For i = 1 To count
                If pArr(i) > 0 Then sumP = sumP + pArr(i): cntP = cntP + 1
            Next i
            If cntP > 0 Then avgs(nKeys) = sumP / cntP Else avgs(nKeys) = 0#
        Next k

        Dim keep As Object: Set keep = CreateObject("Scripting.Dictionary"): keep.CompareMode = vbTextCompare
        Dim otherAmt() As Double, otherPct() As Double
        ReDim otherAmt(1 To count)
        ReDim otherPct(1 To count)
        For i = 1 To nKeys
            If avgs(i) >= thresh Then
                keep(names(i)) = 1
            Else
                Dim aAmt() As Double: aAmt = coinAmt(names(i))
                Dim aPct() As Double: aPct = coinPct(names(i))
                For idx = 1 To count
                    otherAmt(idx) = otherAmt(idx) + aAmt(idx)
                    otherPct(idx) = otherPct(idx) + aPct(idx)
                Next idx
            End If
        Next i

        ' Sort kept by average percent desc
        Dim keptCount As Long: keptCount = keep.Count
        Dim keptNames() As String, keptAvgs() As Double
        If keptCount > 0 Then
            ReDim keptNames(1 To keptCount)
            ReDim keptAvgs(1 To keptCount)
            Dim t As Long: t = 0
            For i = 1 To nKeys
                If keep.Exists(names(i)) Then
                    t = t + 1
                    keptNames(t) = names(i)
                    keptAvgs(t) = avgs(i)
                End If
            Next i
            Dim ii As Long, jj As Long, maxIdx As Long
            For ii = 1 To keptCount - 1
                maxIdx = ii
                For jj = ii + 1 To keptCount
                    If keptAvgs(jj) > keptAvgs(maxIdx) Then maxIdx = jj
                Next jj
                If maxIdx <> ii Then
                    tv = keptAvgs(ii): keptAvgs(ii) = keptAvgs(maxIdx): keptAvgs(maxIdx) = tv
                    ts = keptNames(ii): keptNames(ii) = keptNames(maxIdx): keptNames(maxIdx) = ts
                End If
            Next ii
        End If

        Dim finalCount As Long: finalCount = keptCount
        Dim hasOther As Boolean: hasOther = False
        For idx = 1 To count
            If otherAmt(idx) <> 0 Then hasOther = True: Exit For
        Next idx
        If hasOther Then finalCount = finalCount + 1

        ' Apply normal stacked columns of AMOUNTS, no data labels
        ch.ChartType = xlColumnStacked
        ch.HasTitle = True
        ch.ChartTitle.Text = "Portfolio_Alt.TOP"
        ch.HasLegend = True

        ' Remove any previous AxisHelper
        On Error Resume Next
        For i = ch.SeriesCollection.Count To 1 Step -1
            If StrComp(ch.SeriesCollection(i).Name, "AxisHelper_Pct", vbTextCompare) = 0 Then ch.SeriesCollection(i).Delete
        Next i
        On Error GoTo 0

        Dim scnt As Long: scnt = ch.SeriesCollection.Count
        If scnt < finalCount Then
            For i = scnt + 1 To finalCount
                ch.SeriesCollection.NewSeries
            Next i
        End If
        On Error Resume Next
        Do While ch.SeriesCollection.Count > finalCount
            ch.SeriesCollection(ch.SeriesCollection.Count).Delete
        Loop
        On Error GoTo 0

        Dim pos As Long: pos = 0
        Dim sCoin As Series
        Dim j As Long
        For i = 1 To keptCount
            pos = pos + 1
            Set sCoin = ch.SeriesCollection(pos)
            sCoin.Name = keptNames(i)
            sCoin.Values = coinAmt(keptNames(i))
            sCoin.XValues = datesArr
            sCoin.HasDataLabels = False
        Next i
        If hasOther Then
            pos = pos + 1
            Set sCoin = ch.SeriesCollection(pos)
            sCoin.Name = "Other"
            sCoin.Values = otherAmt
            sCoin.XValues = datesArr
            sCoin.HasDataLabels = False
        End If

        On Error Resume Next
        ch.Axes(xlCategory).CategoryType = xlTimeScale
        ch.Axes(xlCategory).TickLabels.NumberFormat = mod_config.SNAPSHOT_DATE_FMT
        ch.Axes(xlValue).TickLabels.NumberFormat = mod_config.MONEY_FMT
        On Error GoTo 0

        ' Ensure no secondary percent axis/helper series is present
        On Error Resume Next
        For j = ch.SeriesCollection.Count To 1 Step -1
            If StrComp(ch.SeriesCollection(j).Name, "AxisHelper_Pct", vbTextCompare) = 0 Then ch.SeriesCollection(j).Delete
        Next j
        On Error GoTo 0
    Else
        On Error Resume Next
        Do While ch.SeriesCollection.Count > 0
            ch.SeriesCollection(1).Delete
        Loop
        ch.ChartType = xlColumnStacked
        ch.HasTitle = True
        ch.ChartTitle.Text = "Portfolio_Alt.TOP"
        On Error GoTo 0
    End If

    ' ===== Portfolio_Alt.MID =====
    Dim altMidCol As Long: altMidCol = 0
    For c = 1 To lastCol
        hname = LCase$(Trim$(CStr(wsSnap.Cells(1, c).Value)))
        norm = Replace(Replace(hname, ".", ""), " ", "")
        If c >= 8 And (norm = "altmid") Then altMidCol = c
    Next c
    Set co = GetOrCreateChart(wsDash, "Portfolio_Alt.MID")
    Set ch = co.Chart
    If altMidCol > 0 And count > 0 Then
        Dim coinPctM As Object: Set coinPctM = CreateObject("Scripting.Dictionary"): coinPctM.CompareMode = vbTextCompare
        Dim coinAmtM As Object: Set coinAmtM = CreateObject("Scripting.Dictionary"): coinAmtM.CompareMode = vbTextCompare
        idx2 = 0
        Dim altMidTotal As Double
        For r = 2 To lastR
            If IsDate(wsSnap.Cells(r, 1).Value) Then
                dt = DateValue(wsSnap.Cells(r, 1).Value)
                If dt >= dStart And dt <= dEnd Then
                    idx2 = idx2 + 1
                    If IsNumeric(wsSnap.Cells(r, altMidCol).Value) Then altMidTotal = CDbl(wsSnap.Cells(r, altMidCol).Value) Else altMidTotal = 0#
                    If holdCol > 0 Then Set holds = ParseHoldingsString(CStr(wsSnap.Cells(r, holdCol).Value)) Else Set holds = CreateObject("Scripting.Dictionary"): holds.CompareMode = vbTextCompare
                    For Each ck In holds.Keys
                        coinName = CStr(ck)
                        amt = CDbl(holds(ck))
                        grp = ""
                        ovr = GroupOverride(coinName)
                        If Len(ovr) > 0 Then
                            grp = ovr
                        ElseIf UCase$(coinName) = "BTC" Then
                            grp = "BTC"
                        ElseIf Not (coinToGroup Is Nothing) And coinToGroup.Exists(coinName) Then
                            grp = CStr(coinToGroup(coinName))
                        End If
                        If StrComp(grp, "Alt.MID", vbTextCompare) = 0 Then
                            If Not coinPctM.Exists(coinName) Then ReDim arr(1 To count): coinPctM(coinName) = arr
                            If Not coinAmtM.Exists(coinName) Then ReDim arrAmt(1 To count): coinAmtM(coinName) = arrAmt
                            arr = coinPctM(coinName): arrAmt = coinAmtM(coinName)
                            arrAmt(idx2) = amt
                            If Abs(altMidTotal) > mod_config.EPS_CLOSE Then arr(idx2) = amt / altMidTotal Else arr(idx2) = 0#
                            coinPctM(coinName) = arr: coinAmtM(coinName) = arrAmt
                        End If
                    Next ck
                End If
            End If
        Next r
        ' Fallback proportional across non-BTC if no mapped coins
        If coinPctM.Count = 0 Then
            Dim fbPctM As Object: Set fbPctM = CreateObject("Scripting.Dictionary"): fbPctM.CompareMode = vbTextCompare
            Dim fbAmtM As Object: Set fbAmtM = CreateObject("Scripting.Dictionary"): fbAmtM.CompareMode = vbTextCompare
            idx2 = 0
            For r = 2 To lastR
                If IsDate(wsSnap.Cells(r, 1).Value) Then
                    dt = DateValue(wsSnap.Cells(r, 1).Value)
                    If dt >= dStart And dt <= dEnd Then
                        idx2 = idx2 + 1
                        If IsNumeric(wsSnap.Cells(r, altMidCol).Value) Then altMidTotal = CDbl(wsSnap.Cells(r, altMidCol).Value) Else altMidTotal = 0#
                        If holdCol > 0 Then Set holds = ParseHoldingsString(CStr(wsSnap.Cells(r, holdCol).Value)) Else Set holds = CreateObject("Scripting.Dictionary"): holds.CompareMode = vbTextCompare
                        sumNonBTC = 0#: For Each ck In holds.Keys: If StrComp(CStr(ck), "BTC", vbTextCompare) <> 0 Then sumNonBTC = sumNonBTC + CDbl(holds(ck))
                        Next ck
                        If sumNonBTC > mod_config.EPS_CLOSE And altMidTotal > mod_config.EPS_CLOSE Then
                            For Each ck In holds.Keys
                                coinName = CStr(ck)
                                If StrComp(coinName, "BTC", vbTextCompare) <> 0 Then
                                    If Not fbPctM.Exists(coinName) Then ReDim arr(1 To count): fbPctM(coinName) = arr
                                    If Not fbAmtM.Exists(coinName) Then ReDim arrAmt(1 To count): fbAmtM(coinName) = arrAmt
                                    arr = fbPctM(coinName): arrAmt = fbAmtM(coinName)
                                    frac = CDbl(holds(ck)) / sumNonBTC
                                    arr(idx2) = frac: arrAmt(idx2) = frac * altMidTotal
                                    fbPctM(coinName) = arr: fbAmtM(coinName) = arrAmt
                                End If
                            Next ck
                        End If
                    End If
                End If
            Next r
            Set coinPctM = fbPctM: Set coinAmtM = fbAmtM
        End If

        ' Build sorted kept list and Other
        thresh = 0.01: nCoins = coinPctM.Count
        ReDim names(1 To IIf(nCoins = 0, 1, nCoins))
        ReDim avgs(1 To IIf(nCoins = 0, 1, nCoins))
        nKeys = 0
        For Each k In coinPctM.Keys
            nKeys = nKeys + 1
            names(nKeys) = CStr(k)
            pArr = coinPctM(k)
            sumP = 0#: cntP = 0
            For i = 1 To count: If pArr(i) > 0 Then sumP = sumP + pArr(i): cntP = cntP + 1
            Next i
            If cntP > 0 Then avgs(nKeys) = sumP / cntP Else avgs(nKeys) = 0#
        Next k
        Set keep = CreateObject("Scripting.Dictionary"): keep.CompareMode = vbTextCompare
        ReDim otherAmt(1 To count): ReDim otherPct(1 To count)
        For i = 1 To nKeys
            If avgs(i) >= thresh Then
                keep(names(i)) = 1
            Else
                aAmt = coinAmtM(names(i))
                aPct = coinPctM(names(i))
                For idx = 1 To count
                    otherAmt(idx) = otherAmt(idx) + aAmt(idx)
                    otherPct(idx) = otherPct(idx) + aPct(idx)
                Next idx
            End If
        Next i
        keptCount = keep.Count: If keptCount > 0 Then ReDim keptNames(1 To keptCount): ReDim keptAvgs(1 To keptCount)
        t = 0: For i = 1 To nKeys: If keep.Exists(names(i)) Then t = t + 1: keptNames(t) = names(i): keptAvgs(t) = avgs(i)
        Next i
        For ii = 1 To keptCount - 1
            maxIdx = ii
            For jj = ii + 1 To keptCount
                If keptAvgs(jj) > keptAvgs(maxIdx) Then maxIdx = jj
            Next jj
            If maxIdx <> ii Then
                tv = keptAvgs(ii)
                keptAvgs(ii) = keptAvgs(maxIdx)
                keptAvgs(maxIdx) = tv
                ts = keptNames(ii)
                keptNames(ii) = keptNames(maxIdx)
                keptNames(maxIdx) = ts
            End If
        Next ii
        finalCount = keptCount
        hasOther = False
        For idx = 1 To count
            If otherAmt(idx) <> 0 Then
                hasOther = True
                Exit For
            End If
        Next idx
        If hasOther Then finalCount = finalCount + 1

        ch.ChartType = xlColumnStacked: ch.HasTitle = True: ch.ChartTitle.Text = "Portfolio_Alt.MID": ch.HasLegend = True
        On Error Resume Next
        For j = ch.SeriesCollection.Count To 1 Step -1
            If StrComp(ch.SeriesCollection(j).Name, "AxisHelper_Pct", vbTextCompare) = 0 Then
                ch.SeriesCollection(j).Delete
            End If
        Next j
        On Error GoTo 0
        scnt = ch.SeriesCollection.Count: If scnt < finalCount Then For i = scnt + 1 To finalCount: ch.SeriesCollection.NewSeries: Next i
        On Error Resume Next: Do While ch.SeriesCollection.Count > finalCount: ch.SeriesCollection(ch.SeriesCollection.Count).Delete: Loop: On Error GoTo 0
        pos = 0: For i = 1 To keptCount: pos = pos + 1: Set sCoin = ch.SeriesCollection(pos): sCoin.Name = keptNames(i): sCoin.Values = coinAmtM(keptNames(i)): sCoin.XValues = datesArr: sCoin.HasDataLabels = False: Next i
        If hasOther Then pos = pos + 1: Set sCoin = ch.SeriesCollection(pos): sCoin.Name = "Other": sCoin.Values = otherAmt: sCoin.XValues = datesArr: sCoin.HasDataLabels = False
        On Error Resume Next: ch.Axes(xlCategory).CategoryType = xlTimeScale: ch.Axes(xlCategory).TickLabels.NumberFormat = mod_config.SNAPSHOT_DATE_FMT: ch.Axes(xlValue).TickLabels.NumberFormat = mod_config.MONEY_FMT: On Error GoTo 0
    Else
        On Error Resume Next: Do While ch.SeriesCollection.Count > 0: ch.SeriesCollection(1).Delete: Loop: ch.ChartType = xlColumnStacked: ch.HasTitle = True: ch.ChartTitle.Text = "Portfolio_Alt.MID": On Error GoTo 0
    End If

    ' ===== Portfolio_Alt.LOW =====
    Dim altLowCol As Long: altLowCol = 0
    For c = 1 To lastCol
        hname = LCase$(Trim$(CStr(wsSnap.Cells(1, c).Value)))
        norm = Replace(Replace(hname, ".", ""), " ", "")
        If c >= 8 And (norm = "altlow") Then altLowCol = c
    Next c
    Set co = GetOrCreateChart(wsDash, "Portfolio_Alt.LOW")
    Set ch = co.Chart
    If altLowCol > 0 And count > 0 Then
        Dim coinPctL As Object: Set coinPctL = CreateObject("Scripting.Dictionary"): coinPctL.CompareMode = vbTextCompare
        Dim coinAmtL As Object: Set coinAmtL = CreateObject("Scripting.Dictionary"): coinAmtL.CompareMode = vbTextCompare
        idx2 = 0
        Dim altLowTotal As Double
        For r = 2 To lastR
            If IsDate(wsSnap.Cells(r, 1).Value) Then
                dt = DateValue(wsSnap.Cells(r, 1).Value)
                If dt >= dStart And dt <= dEnd Then
                    idx2 = idx2 + 1
                    If IsNumeric(wsSnap.Cells(r, altLowCol).Value) Then altLowTotal = CDbl(wsSnap.Cells(r, altLowCol).Value) Else altLowTotal = 0#
                    If holdCol > 0 Then Set holds = ParseHoldingsString(CStr(wsSnap.Cells(r, holdCol).Value)) Else Set holds = CreateObject("Scripting.Dictionary"): holds.CompareMode = vbTextCompare
                    For Each ck In holds.Keys
                        coinName = CStr(ck)
                        amt = CDbl(holds(ck))
                        grp = ""
                        ovr = GroupOverride(coinName)
                        If Len(ovr) > 0 Then
                            grp = ovr
                        ElseIf UCase$(coinName) = "BTC" Then
                            grp = "BTC"
                        ElseIf Not (coinToGroup Is Nothing) And coinToGroup.Exists(coinName) Then
                            grp = CStr(coinToGroup(coinName))
                        End If
                        If StrComp(grp, "Alt.LOW", vbTextCompare) = 0 Then
                            If Not coinPctL.Exists(coinName) Then ReDim arr(1 To count): coinPctL(coinName) = arr
                            If Not coinAmtL.Exists(coinName) Then ReDim arrAmt(1 To count): coinAmtL(coinName) = arrAmt
                            arr = coinPctL(coinName): arrAmt = coinAmtL(coinName)
                            arrAmt(idx2) = amt
                            If Abs(altLowTotal) > mod_config.EPS_CLOSE Then arr(idx2) = amt / altLowTotal Else arr(idx2) = 0#
                            coinPctL(coinName) = arr: coinAmtL(coinName) = arrAmt
                        End If
                    Next ck
                End If
            End If
        Next r
        If coinPctL.Count = 0 Then
            Dim fbPctL As Object: Set fbPctL = CreateObject("Scripting.Dictionary"): fbPctL.CompareMode = vbTextCompare
            Dim fbAmtL As Object: Set fbAmtL = CreateObject("Scripting.Dictionary"): fbAmtL.CompareMode = vbTextCompare
            idx2 = 0
            For r = 2 To lastR
                If IsDate(wsSnap.Cells(r, 1).Value) Then
                    dt = DateValue(wsSnap.Cells(r, 1).Value)
                    If dt >= dStart And dt <= dEnd Then
                        idx2 = idx2 + 1
                        If IsNumeric(wsSnap.Cells(r, altLowCol).Value) Then altLowTotal = CDbl(wsSnap.Cells(r, altLowCol).Value) Else altLowTotal = 0#
                        If holdCol > 0 Then Set holds = ParseHoldingsString(CStr(wsSnap.Cells(r, holdCol).Value)) Else Set holds = CreateObject("Scripting.Dictionary"): holds.CompareMode = vbTextCompare
                        sumNonBTC = 0#: For Each ck In holds.Keys: If StrComp(CStr(ck), "BTC", vbTextCompare) <> 0 Then sumNonBTC = sumNonBTC + CDbl(holds(ck))
                        Next ck
                        If sumNonBTC > mod_config.EPS_CLOSE And altLowTotal > mod_config.EPS_CLOSE Then
                            For Each ck In holds.Keys
                                coinName = CStr(ck)
                                If StrComp(coinName, "BTC", vbTextCompare) <> 0 Then
                                    If Not fbPctL.Exists(coinName) Then ReDim arr(1 To count): fbPctL(coinName) = arr
                                    If Not fbAmtL.Exists(coinName) Then ReDim arrAmt(1 To count): fbAmtL(coinName) = arrAmt
                                    arr = fbPctL(coinName): arrAmt = fbAmtL(coinName)
                                    frac = CDbl(holds(ck)) / sumNonBTC
                                    arr(idx2) = frac: arrAmt(idx2) = frac * altLowTotal
                                    fbPctL(coinName) = arr: fbAmtL(coinName) = arrAmt
                                End If
                            Next ck
                        End If
                    End If
                End If
            Next r
            Set coinPctL = fbPctL: Set coinAmtL = fbAmtL
        End If

        ' Build sorted kept list and Other for LOW
        thresh = 0.01: nCoins = coinPctL.Count
        ReDim names(1 To IIf(nCoins = 0, 1, nCoins))
        ReDim avgs(1 To IIf(nCoins = 0, 1, nCoins))
        nKeys = 0
        For Each k In coinPctL.Keys
            nKeys = nKeys + 1: names(nKeys) = CStr(k): pArr = coinPctL(k)
            sumP = 0#: cntP = 0
            For i = 1 To count
                If pArr(i) > 0 Then
                    sumP = sumP + pArr(i)
                    cntP = cntP + 1
                End If
            Next i
            If cntP > 0 Then avgs(nKeys) = sumP / cntP Else avgs(nKeys) = 0#
        Next k
        Set keep = CreateObject("Scripting.Dictionary"): keep.CompareMode = vbTextCompare
        ReDim otherAmt(1 To count): ReDim otherPct(1 To count)
        For i = 1 To nKeys
            If avgs(i) >= thresh Then
                keep(names(i)) = 1
            Else
                aAmt = coinAmtL(names(i))
                aPct = coinPctL(names(i))
                For idx = 1 To count
                    otherAmt(idx) = otherAmt(idx) + aAmt(idx)
                    otherPct(idx) = otherPct(idx) + aPct(idx)
                Next idx
            End If
        Next i
        keptCount = keep.Count: If keptCount > 0 Then ReDim keptNames(1 To keptCount): ReDim keptAvgs(1 To keptCount)
        t = 0: For i = 1 To nKeys: If keep.Exists(names(i)) Then t = t + 1: keptNames(t) = names(i): keptAvgs(t) = avgs(i)
        Next i
        For ii = 1 To keptCount - 1
            maxIdx = ii
            For jj = ii + 1 To keptCount
                If keptAvgs(jj) > keptAvgs(maxIdx) Then maxIdx = jj
            Next jj
            If maxIdx <> ii Then
                tv = keptAvgs(ii)
                keptAvgs(ii) = keptAvgs(maxIdx)
                keptAvgs(maxIdx) = tv
                ts = keptNames(ii)
                keptNames(ii) = keptNames(maxIdx)
                keptNames(maxIdx) = ts
            End If
        Next ii
        finalCount = keptCount
        hasOther = False
        For idx = 1 To count
            If otherAmt(idx) <> 0 Then
                hasOther = True
                Exit For
            End If
        Next idx
        If hasOther Then finalCount = finalCount + 1

        ch.ChartType = xlColumnStacked: ch.HasTitle = True: ch.ChartTitle.Text = "Portfolio_Alt.LOW": ch.HasLegend = True
        On Error Resume Next
        For j = ch.SeriesCollection.Count To 1 Step -1
            If StrComp(ch.SeriesCollection(j).Name, "AxisHelper_Pct", vbTextCompare) = 0 Then
                ch.SeriesCollection(j).Delete
            End If
        Next j
        On Error GoTo 0
        scnt = ch.SeriesCollection.Count: If scnt < finalCount Then For i = scnt + 1 To finalCount: ch.SeriesCollection.NewSeries: Next i
        On Error Resume Next: Do While ch.SeriesCollection.Count > finalCount: ch.SeriesCollection(ch.SeriesCollection.Count).Delete: Loop: On Error GoTo 0
        pos = 0: For i = 1 To keptCount: pos = pos + 1: Set sCoin = ch.SeriesCollection(pos): sCoin.Name = keptNames(i): sCoin.Values = coinAmtL(keptNames(i)): sCoin.XValues = datesArr: sCoin.HasDataLabels = False: Next i
        If hasOther Then pos = pos + 1: Set sCoin = ch.SeriesCollection(pos): sCoin.Name = "Other": sCoin.Values = otherAmt: sCoin.XValues = datesArr: sCoin.HasDataLabels = False
        On Error Resume Next: ch.Axes(xlCategory).CategoryType = xlTimeScale: ch.Axes(xlCategory).TickLabels.NumberFormat = mod_config.SNAPSHOT_DATE_FMT: ch.Axes(xlValue).TickLabels.NumberFormat = mod_config.MONEY_FMT: On Error GoTo 0
    Else
        On Error Resume Next: Do While ch.SeriesCollection.Count > 0: ch.SeriesCollection(1).Delete: Loop: ch.ChartType = xlColumnStacked: ch.HasTitle = True: ch.ChartTitle.Text = "Portfolio_Alt.LOW": On Error GoTo 0
    End If

    ' Refactored Alt group charts (helper-based)
    Dim headerMap As Object: Set headerMap = GetHeaderMap(wsSnap)
    BuildAltGroupChart wsDash, wsSnap, headerMap, dStart, dEnd, "Alt.TOP", "Portfolio_Alt.TOP", datesArr, count
    BuildAltGroupChart wsDash, wsSnap, headerMap, dStart, dEnd, "Alt.MID", "Portfolio_Alt.MID", datesArr, count
    BuildAltGroupChart wsDash, wsSnap, headerMap, dStart, dEnd, "Alt.LOW", "Portfolio_Alt.LOW", datesArr, count

CleanExit:
    On Error Resume Next
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Application.Calculation = prevCalc
    On Error GoTo 0
    Exit Sub
Fail:
    MsgBox "Error (Update_Dashboard): " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Function GetOrCreateChart(ws As Worksheet, ByVal chartName As String) As ChartObject
    On Error Resume Next
    Dim co As ChartObject
    Set co = ws.ChartObjects(chartName)
    On Error GoTo 0

    ' Try common alias/typo names and normalize the name if found
    If co Is Nothing Then
        Dim alt As Variant
        Dim alts As Variant
        alts = Array("Portfoilo_" & Mid$(chartName, InStr(1, chartName, "_") + 1), _
                     Replace(chartName, "_", " "), _
                     Replace("Portfoilo_" & Mid$(chartName, InStr(1, chartName, "_") + 1), "_", " "))
        For Each alt In alts
            On Error Resume Next
            Set co = ws.ChartObjects(CStr(alt))
            On Error GoTo 0
            If Not co Is Nothing Then
                On Error Resume Next
                co.Name = chartName
                On Error GoTo 0
                Exit For
            End If
        Next alt
        ' Legacy rename support: if requesting Portfolio_Catagory, also look for Portfolio_Group
        If co Is Nothing Then
            If StrComp(chartName, "Portfolio_Catagory", vbTextCompare) = 0 Then
                On Error Resume Next
                Set co = ws.ChartObjects("Portfolio_Group")
                If co Is Nothing Then Set co = ws.ChartObjects("Portfolio Group")
                On Error GoTo 0
                If Not co Is Nothing Then
                    On Error Resume Next
                    co.Name = chartName
                    On Error GoTo 0
                End If
            End If
        End If
    End If

    If co Is Nothing Then
        Set co = ws.ChartObjects.Add(Left:=20, Top:=20, Width:=520, Height:=260)
        co.Name = chartName
        co.Chart.ChartType = xlLine
        co.Chart.HasTitle = True
        co.Chart.ChartTitle.Text = chartName
    End If
    Set GetOrCreateChart = co
End Function

Private Function BuildCoinToGroupMap(wsC As Worksheet) As Object
    On Error GoTo Done
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    ' Accept both "Categoty" (current), "Catagory" and "Category" sheet names
    If wsC Is Nothing Then
        Set wsC = SheetByName("Categoty")
        If wsC Is Nothing Then Set wsC = SheetByName("Catagory")
        If wsC Is Nothing Then Set wsC = SheetByName("Category")
    End If
    If wsC Is Nothing Then Set BuildCoinToGroupMap = d: Exit Function

    Dim lastR As Long, lastCol As Long
    lastR = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
    lastCol = wsC.Cells(1, wsC.Columns.Count).End(xlToLeft).Column
    If lastR < 2 And lastCol < 1 Then Set BuildCoinToGroupMap = d: Exit Function

    ' Try two-column mapping: headers include Coin and Group/Category/Catagory
    Dim coinCol As Long: coinCol = 0
    Dim grpCol As Long: grpCol = 0
    Dim c As Long, h As String
    For c = 1 To lastCol
        h = LCase$(Trim$(CStr(wsC.Cells(1, c).Value)))
        If InStr(h, "coin") > 0 Then coinCol = c
        If InStr(h, "group") > 0 Or InStr(h, "category") > 0 Or InStr(h, "catagory") > 0 Then grpCol = c
    Next c
    If coinCol > 0 And grpCol > 0 Then
        Dim r As Long, coin As String, grp As String
        For r = 2 To lastR
            coin = Trim$(CStr(wsC.Cells(r, coinCol).Value))
            grp = Trim$(CStr(wsC.Cells(r, grpCol).Value))
            If Len(coin) > 0 And Len(grp) > 0 Then d(coin) = grp
        Next r
        Set BuildCoinToGroupMap = d: Exit Function
    End If

    ' Fallback: multi-column with group names in row 1; coins listed below
    Dim g As Long, groupName As String
    For g = 1 To lastCol
        groupName = Trim$(CStr(wsC.Cells(1, g).Value))
        If Len(groupName) > 0 Then
            For r = 2 To wsC.Cells(wsC.Rows.Count, g).End(xlUp).Row
                coin = Trim$(CStr(wsC.Cells(r, g).Value))
                If Len(coin) > 0 Then d(coin) = groupName
            Next r
        End If
    Next g

    Set BuildCoinToGroupMap = d
    Exit Function
Done:
    Set BuildCoinToGroupMap = CreateObject("Scripting.Dictionary")
End Function

Private Function NormalizeHeader(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace(t, ".", "")
    t = Replace(t, " ", "")
    NormalizeHeader = t
End Function

Private Function GetHeaderMap(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    If ws Is Nothing Then Set GetHeaderMap = d: Exit Function
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, raw As String, key As String
    For c = 1 To lastCol
        raw = CStr(ws.Cells(1, c).Value)
        key = NormalizeHeader(raw)
        If Len(key) > 0 Then d(key) = c
    Next c
    ' Also map generic holdings if header contains "holding"
    For c = 1 To lastCol
        raw = CStr(ws.Cells(1, c).Value)
        If InStr(1, raw, "holding", vbTextCompare) > 0 Then d("holdings") = c
    Next c
    Set GetHeaderMap = d
End Function

Private Sub BuildAltGroupChart(wsDash As Worksheet, wsSnap As Worksheet, headerMap As Object, _
    ByVal dStart As Date, ByVal dEnd As Date, ByVal groupName As String, ByVal chartName As String, _
    ByRef datesArr() As Variant, ByVal n As Long)

    On Error GoTo Done
    Dim chObj As ChartObject, ch As Chart
    Set chObj = GetOrCreateChart(wsDash, chartName)
    Set ch = chObj.Chart

    Dim groupKey As String: groupKey = NormalizeHeader(groupName)
    Dim groupCol As Long
    If Not (headerMap Is Nothing) And headerMap.Exists(groupKey) Then groupCol = CLng(headerMap(groupKey)) Else groupCol = 0
    Dim holdCol As Long
    If Not (headerMap Is Nothing) And headerMap.Exists("holdings") Then holdCol = CLng(headerMap("holdings")) Else holdCol = 0

    Dim lastR As Long: lastR = wsSnap.Cells(wsSnap.Rows.Count, 1).End(xlUp).Row
    If groupCol = 0 Or n = 0 Then
        On Error Resume Next
        Do While ch.SeriesCollection.Count > 0
            ch.SeriesCollection(1).Delete
        Loop
        On Error GoTo 0
        ch.ChartType = xlColumnStacked
        ch.HasTitle = True
        ch.ChartTitle.Text = chartName
        Exit Sub
    End If

    Dim wsC As Worksheet: Set wsC = SheetByName(mod_config.SHEET_CATEGORY)
    If wsC Is Nothing Then Set wsC = SheetByName("Categoty")
    If wsC Is Nothing Then Set wsC = SheetByName("Catagory")
    If wsC Is Nothing Then Set wsC = SheetByName("Category")
    Dim coinToGroup As Object: Set coinToGroup = BuildCoinToGroupMap(wsC)

    Dim coinPct As Object: Set coinPct = CreateObject("Scripting.Dictionary"): coinPct.CompareMode = vbTextCompare
    Dim coinAmt As Object: Set coinAmt = CreateObject("Scripting.Dictionary"): coinAmt.CompareMode = vbTextCompare

    Dim idx As Long: idx = 0
    Dim r As Long, dt As Date, altTotal As Double
    Dim holds As Object
    For r = 2 To lastR
        If IsDate(wsSnap.Cells(r, 1).Value) Then
            dt = DateValue(wsSnap.Cells(r, 1).Value)
            If dt >= dStart And dt <= dEnd Then
                idx = idx + 1
                If IsNumeric(wsSnap.Cells(r, groupCol).Value) Then altTotal = CDbl(wsSnap.Cells(r, groupCol).Value) Else altTotal = 0#
                If holdCol > 0 Then Set holds = ParseHoldingsString(CStr(wsSnap.Cells(r, holdCol).Value)) Else Set holds = CreateObject("Scripting.Dictionary"): holds.CompareMode = vbTextCompare
                Dim ck As Variant, coinName As String, amt As Double
                For Each ck In holds.Keys
                    coinName = CStr(ck)
                    amt = CDbl(holds(ck))
                    Dim grp As String: grp = ""
                    Dim ovr As String: ovr = GroupOverride(coinName)
                    If Len(ovr) > 0 Then
                        grp = ovr
                    ElseIf UCase$(coinName) = "BTC" Then
                        grp = "BTC"
                    ElseIf Not (coinToGroup Is Nothing) And coinToGroup.Exists(coinName) Then
                        grp = CStr(coinToGroup(coinName))
                    End If
                    ' Compare normalized group labels to tolerate punctuation/spacing differences
                    If StrComp(NormalizeHeader(grp), NormalizeHeader(groupName), vbTextCompare) = 0 Then
                        Dim aP() As Double, aA() As Double
                        If Not coinPct.Exists(coinName) Then ReDim aP(1 To n): coinPct(coinName) = aP
                        If Not coinAmt.Exists(coinName) Then ReDim aA(1 To n): coinAmt(coinName) = aA
                        aP = coinPct(coinName): aA = coinAmt(coinName)
                        aA(idx) = amt
                        If Abs(altTotal) > mod_config.EPS_CLOSE Then aP(idx) = amt / altTotal Else aP(idx) = 0#
                        coinPct(coinName) = aP: coinAmt(coinName) = aA
                    End If
                Next ck
            End If
        End If
    Next r

    ' Fallback proportional split if nothing mapped
    If coinPct.Count = 0 Then
        idx = 0
        For r = 2 To lastR
            If IsDate(wsSnap.Cells(r, 1).Value) Then
                dt = DateValue(wsSnap.Cells(r, 1).Value)
                If dt >= dStart And dt <= dEnd Then
                    idx = idx + 1
                    If IsNumeric(wsSnap.Cells(r, groupCol).Value) Then altTotal = CDbl(wsSnap.Cells(r, groupCol).Value) Else altTotal = 0#
                    If holdCol > 0 Then Set holds = ParseHoldingsString(CStr(wsSnap.Cells(r, holdCol).Value)) Else Set holds = CreateObject("Scripting.Dictionary"): holds.CompareMode = vbTextCompare
                    Dim sumNonBTC As Double: sumNonBTC = 0#
                    For Each ck In holds.Keys
                        If StrComp(CStr(ck), "BTC", vbTextCompare) <> 0 Then sumNonBTC = sumNonBTC + CDbl(holds(ck))
                    Next ck
                    If sumNonBTC > mod_config.EPS_CLOSE And altTotal > mod_config.EPS_CLOSE Then
                        For Each ck In holds.Keys
                            coinName = CStr(ck)
                            If StrComp(coinName, "BTC", vbTextCompare) <> 0 Then
                                Dim bP() As Double, bA() As Double
                                If Not coinPct.Exists(coinName) Then ReDim bP(1 To n): coinPct(coinName) = bP
                                If Not coinAmt.Exists(coinName) Then ReDim bA(1 To n): coinAmt(coinName) = bA
                                bP = coinPct(coinName): bA = coinAmt(coinName)
                                Dim frac As Double: frac = CDbl(holds(ck)) / sumNonBTC
                                bP(idx) = frac: bA(idx) = frac * altTotal
                                coinPct(coinName) = bP: coinAmt(coinName) = bA
                            End If
                        Next ck
                    End If
                End If
            End If
        Next r
    End If

    ' Group tiny slices and sort by average share
    Dim thresh As Double: thresh = 0.01
    Dim nCoins As Long: nCoins = coinPct.Count
    Dim names() As String, avgs() As Double
    If nCoins > 0 Then
        ReDim names(1 To nCoins): ReDim avgs(1 To nCoins)
    Else
        ReDim names(1 To 1): ReDim avgs(1 To 1)
    End If
    Dim i As Long, k As Variant, nKeys As Long: nKeys = 0
    For Each k In coinPct.Keys
        nKeys = nKeys + 1
        names(nKeys) = CStr(k)
        Dim pArr() As Double: pArr = coinPct(k)
        Dim sumP As Double: sumP = 0#: Dim cntP As Long: cntP = 0
        For i = 1 To n
            If pArr(i) > 0 Then sumP = sumP + pArr(i): cntP = cntP + 1
        Next i
        If cntP > 0 Then avgs(nKeys) = sumP / cntP Else avgs(nKeys) = 0#
    Next k

    Dim keep As Object: Set keep = CreateObject("Scripting.Dictionary"): keep.CompareMode = vbTextCompare
    Dim otherAmt() As Double, otherPct() As Double
    ReDim otherAmt(1 To IIf(n = 0, 1, n))
    ReDim otherPct(1 To IIf(n = 0, 1, n))
    Dim idx2 As Long
    For i = 1 To nKeys
        If avgs(i) >= thresh Then
            keep(names(i)) = 1
        Else
            Dim aAmt() As Double: aAmt = coinAmt(names(i))
            Dim aPct() As Double: aPct = coinPct(names(i))
            For idx2 = 1 To n
                otherAmt(idx2) = otherAmt(idx2) + aAmt(idx2)
                otherPct(idx2) = otherPct(idx2) + aPct(idx2)
            Next idx2
        End If
    Next i

    ' If everything is below threshold (keep empty), then keep all coins to avoid a lone "Other" series
    If keep.Count = 0 And nKeys > 0 Then
        For i = 1 To nKeys
            keep(names(i)) = 1
        Next i
        ' Clear Other arrays since we are keeping all
        For idx2 = 1 To n
            otherAmt(idx2) = 0#: otherPct(idx2) = 0#
        Next idx2
    End If

    Dim keptCount As Long: keptCount = keep.Count
    Dim keptNames() As String, keptAvgs() As Double
    If keptCount > 0 Then ReDim keptNames(1 To keptCount): ReDim keptAvgs(1 To keptCount)
    Dim t As Long: t = 0
    For i = 1 To nKeys
        If keep.Exists(names(i)) Then t = t + 1: keptNames(t) = names(i): keptAvgs(t) = avgs(i)
    Next i
    Dim ii As Long, jj As Long, maxIdx As Long
    For ii = 1 To keptCount - 1
        maxIdx = ii
        For jj = ii + 1 To keptCount
            If keptAvgs(jj) > keptAvgs(maxIdx) Then maxIdx = jj
        Next jj
        If maxIdx <> ii Then
            Dim tv As Double, ts As String
            tv = keptAvgs(ii): keptAvgs(ii) = keptAvgs(maxIdx): keptAvgs(maxIdx) = tv
            ts = keptNames(ii): keptNames(ii) = keptNames(maxIdx): keptNames(maxIdx) = ts
        End If
    Next ii

    Dim hasOther As Boolean: hasOther = False
    For idx2 = 1 To n
        If otherAmt(idx2) <> 0 Then hasOther = True: Exit For
    Next idx2
    Dim finalCount As Long: finalCount = keptCount + IIf(hasOther, 1, 0)

    ' Apply chart
    ch.ChartType = xlColumnStacked
    ch.HasTitle = True
    ch.ChartTitle.Text = chartName
    ch.HasLegend = True
    On Error Resume Next
    Do While ch.SeriesCollection.Count > finalCount And ch.SeriesCollection.Count > 0
        ch.SeriesCollection(ch.SeriesCollection.Count).Delete
    Loop
    If ch.SeriesCollection.Count < finalCount Then
        For i = ch.SeriesCollection.Count + 1 To finalCount
            ch.SeriesCollection.NewSeries
        Next i
    End If
    On Error GoTo 0

    Dim pos As Long: pos = 0
    Dim s As Series
    For i = 1 To keptCount
        pos = pos + 1
        Set s = ch.SeriesCollection(pos)
        s.Name = keptNames(i)
        s.Values = coinAmt(keptNames(i))
        s.XValues = datesArr
        s.HasDataLabels = False
    Next i
    If hasOther Then
        pos = pos + 1
        Set s = ch.SeriesCollection(pos)
        s.Name = "Other"
        s.Values = otherAmt
        s.XValues = datesArr
        s.HasDataLabels = False
    End If
    On Error Resume Next
    ch.Axes(xlCategory).CategoryType = xlTimeScale
    ch.Axes(xlCategory).TickLabels.NumberFormat = mod_config.SNAPSHOT_DATE_FMT
    ch.Axes(xlValue).TickLabels.NumberFormat = mod_config.MONEY_FMT
    On Error GoTo 0
    Exit Sub
Done:
    ' If anything goes wrong, keep chart scaffolded
    On Error Resume Next
    ch.ChartType = xlColumnStacked
    ch.HasTitle = True
    ch.ChartTitle.Text = chartName
    On Error GoTo 0
End Sub

Private Function ParseHoldingsString(ByVal s As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Dim part As Variant, kv As Variant, key As String, valStr As String, valD As Double
    If Len(Trim$(s)) = 0 Then Set ParseHoldingsString = d: Exit Function
    Dim items() As String
    items = Split(s, ";")
    For Each part In items
        kv = Split(part, ":")
        If UBound(kv) >= 1 Then
            key = Trim$(CStr(kv(0)))
            valStr = Trim$(CStr(kv(1)))
            valStr = Replace(valStr, ",", "")
            If IsNumeric(valStr) Then
                valD = CDbl(valStr)
                d(UCase$(key)) = valD
            End If
        End If
    Next part
    Set ParseHoldingsString = d
End Function

Private Function GroupOverride(ByVal coin As String) As String
    ' Override disabled: rely entirely on Category sheet mapping
    GroupOverride = vbNullString
End Function

Private Function CollectHoldingCoins(wsSnap As Worksheet, ByVal holdCol As Long, ByVal dStart As Date, ByVal dEnd As Date) As Object
    Dim coins As Object: Set coins = CreateObject("Scripting.Dictionary"): coins.CompareMode = vbTextCompare
    If wsSnap Is Nothing Or holdCol <= 0 Then Set CollectHoldingCoins = coins: Exit Function
    Dim lastR As Long: lastR = wsSnap.Cells(wsSnap.Rows.Count, 1).End(xlUp).Row
    Dim r As Long, dt As Date, holds As Object, k As Variant
    For r = 2 To lastR
        If IsDate(wsSnap.Cells(r, 1).Value) Then
            dt = DateValue(wsSnap.Cells(r, 1).Value)
            If dt >= dStart And dt <= dEnd Then
                Set holds = ParseHoldingsString(CStr(wsSnap.Cells(r, holdCol).Value))
                For Each k In holds.Keys
                    If Not coins.Exists(CStr(k)) Then coins(CStr(k)) = 1
                Next k
            End If
        End If
    Next r
    Set CollectHoldingCoins = coins
End Function

Private Function ComputeMaxDrawdownPct(vals() As Double) As Double
    On Error GoTo Done
    Dim n As Long
    n = UBound(vals) - LBound(vals) + 1
    If n <= 1 Then GoTo Done

    Dim peak As Double: peak = vals(LBound(vals))
    Dim maxDD As Double: maxDD = 0#
    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        If vals(i) > peak Then peak = vals(i)
        If peak > 0 Then
            Dim dd As Double: dd = (peak - vals(i)) / peak
            If dd > maxDD Then maxDD = dd
        End If
    Next i
    ComputeMaxDrawdownPct = maxDD
    Exit Function
Done:
    ComputeMaxDrawdownPct = 0#
End Function

Private Function ComputeCurrentDrawdownPct(vals() As Double) As Double
    On Error GoTo Done
    Dim n As Long
    n = UBound(vals) - LBound(vals) + 1
    If n <= 1 Then GoTo Done

    Dim peak As Double: peak = vals(LBound(vals))
    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        If vals(i) > peak Then peak = vals(i)
    Next i
    If peak > 0 Then
        Dim lastVal As Double
        lastVal = vals(UBound(vals))
        ' If last value is at all-time high (within tolerance), drawdown is zero
        If lastVal >= peak - mod_config.EPS_CLOSE Then
            ComputeCurrentDrawdownPct = 0#
        Else
            ComputeCurrentDrawdownPct = (peak - lastVal) / peak
        End If
    Else
        ComputeCurrentDrawdownPct = 0#
    End If
    Exit Function
Done:
    ComputeCurrentDrawdownPct = 0#
End Function

Private Sub AnnotateDrawdown(ch As Chart, ByVal mddPct As Double, ByVal curDDPct As Double)
    ' Annotate all charts except Deposit/Withdraw/PnL
    On Error Resume Next
    Dim co As Object
    Set co = ch.Parent
    If Not co Is Nothing Then
        Dim cname As String: cname = LCase$(co.Name)
        If InStr(cname, "deposit") > 0 Or InStr(cname, "withdraw") > 0 Or InStr(cname, "pnl") > 0 Then Exit Sub
    End If
    On Error GoTo 0
    On Error Resume Next
    Dim shp As Shape
    ' Remove old annotation if any
    ch.Shapes("MDD_NAV").Delete
    On Error GoTo 0

    Dim txt As String
    ' Show integer percent only
    txt = "Drawdown: Max. " & Format(mddPct, "0%") & ", Current: " & Format(curDDPct, "0%")

    ' Position from config
    Dim w As Single: w = NAV_MDD_WIDTH
    Dim h As Single: h = NAV_MDD_HEIGHT
    Dim leftPos As Single, topPos As Single
    Dim anchor As String: anchor = LCase$(NAV_MDD_ANCHOR)

    If anchor = "undertitle" And ch.HasTitle Then
        Dim ct As ChartTitle
        Set ct = ch.ChartTitle
        topPos = ct.Top + ct.Height + NAV_MDD_OFFSET_Y
        ' Center under title then apply X offset
        leftPos = ct.Left + (ct.Width - w) / 2 + NAV_MDD_OFFSET_X
    ElseIf anchor = "plottopright" Then
        Dim pa1 As PlotArea
        Set pa1 = ch.PlotArea
        leftPos = pa1.InsideLeft + pa1.InsideWidth - w - 6 + NAV_MDD_OFFSET_X
        topPos = pa1.InsideTop + 6 + NAV_MDD_OFFSET_Y
    ElseIf anchor = "plottopleft" Then
        Dim pa2 As PlotArea
        Set pa2 = ch.PlotArea
        leftPos = pa2.InsideLeft + 6 + NAV_MDD_OFFSET_X
        topPos = pa2.InsideTop + 6 + NAV_MDD_OFFSET_Y
    Else
        ' Fallback: under title if exists, otherwise plot top-left
        If ch.HasTitle Then
            Dim ct2 As ChartTitle
            Set ct2 = ch.ChartTitle
            topPos = ct2.Top + ct2.Height + NAV_MDD_OFFSET_Y
            leftPos = ct2.Left + (ct.Width - w) / 2 + NAV_MDD_OFFSET_X
        Else
            Dim pa3 As PlotArea
            Set pa3 = ch.PlotArea
            leftPos = pa3.InsideLeft + 6 + NAV_MDD_OFFSET_X
            topPos = pa3.InsideTop + 6 + NAV_MDD_OFFSET_Y
        End If
    End If

    Set shp = ch.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos, w, h)
    shp.Name = "MDD_NAV"
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.Size = 9
        .TextRange.Font.Fill.ForeColor.RGB = RGB(80, 80, 80)
        Select Case LCase$(NAV_MDD_ALIGN)
            Case "left":   .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            Case "right":  .TextRange.ParagraphFormat.Alignment = msoAlignRight
            Case Else:      .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End Select
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 1: .MarginBottom = 1
    End With
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse
End Sub

Private Function SheetByName(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function
