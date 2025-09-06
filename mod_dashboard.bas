Attribute VB_Name = "mod_dashboard"
Option Explicit
'
' Last Modified (UTC): 2025-09-06T10:53:51Z

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

    ' Build NAV series from Daily_Snapshot (Date in col A, NAV in col D)
    Dim lastR As Long: lastR = wsSnap.Cells(wsSnap.Rows.Count, 1).End(xlUp).Row
    Dim datesArr() As Variant, navArr() As Double, pnlArr() As Double, depArr() As Double, wdrArr() As Double, ratioArr() As Double
    Dim count As Long: count = 0

    Dim r As Long, dt As Date
    For r = 2 To lastR
        If IsDate(wsSnap.Cells(r, 1).Value) Then
            dt = DateValue(wsSnap.Cells(r, 1).Value)
            If dt >= dStart And dt <= dEnd Then
                count = count + 1
                ReDim Preserve datesArr(1 To count)
                ReDim Preserve navArr(1 To count)
                ReDim Preserve pnlArr(1 To count)
                ReDim Preserve depArr(1 To count)
                ReDim Preserve wdrArr(1 To count)
                ReDim Preserve ratioArr(1 To count)
                datesArr(count) = dt
                If IsNumeric(wsSnap.Cells(r, 4).Value) Then
                    navArr(count) = CDbl(wsSnap.Cells(r, 4).Value)
                Else
                    navArr(count) = 0#
                End If
                If IsNumeric(wsSnap.Cells(r, 7).Value) Then
                    pnlArr(count) = CDbl(wsSnap.Cells(r, 7).Value)
                Else
                    pnlArr(count) = 0#
                End If
                If IsNumeric(wsSnap.Cells(r, 5).Value) Then
                    depArr(count) = CDbl(wsSnap.Cells(r, 5).Value)
                Else
                    depArr(count) = 0#
                End If
                If IsNumeric(wsSnap.Cells(r, 6).Value) Then
                    wdrArr(count) = CDbl(wsSnap.Cells(r, 6).Value)
                Else
                    wdrArr(count) = 0#
                End If
                ' Cash vs NAV percentage from snapshot columns: Cash (col 2) / NAV (col 4)
                Dim cashVal As Double
                If IsNumeric(wsSnap.Cells(r, 2).Value) Then cashVal = CDbl(wsSnap.Cells(r, 2).Value) Else cashVal = 0#
                If Abs(navArr(count)) > mod_config.EPS_CLOSE Then
                    ratioArr(count) = cashVal / navArr(count) ' fraction, format axis as percent
                Else
                    ratioArr(count) = 0#
                End If
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

    ' Build/Apply Portfolio_Group chart (stacked column of category amounts; no labels)
    ' Assumed Daily_Snapshot layout:
    '   Col A = Date, B = Cash, C = Coin (total), D = NAV, E = Deposit, F = Withdraw, G = PnL,
    '   Col H.. = one column per Category (amount in same units as Coin, e.g., USD value of holdings)
    Dim lastCol As Long
    lastCol = wsSnap.Cells(1, wsSnap.Columns.Count).End(xlToLeft).Column
    If lastCol > 7 Then
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
            Dim idx As Long: idx = 0
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
            Set co = GetOrCreateChart(wsDash, "Portfolio_Group")
            Set ch = co.Chart
            ch.ChartType = xlColumnStacked
            ch.HasTitle = True
            ch.ChartTitle.Text = "Portfolio_Group"
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
            Set co = GetOrCreateChart(wsDash, "Portfolio_Group")
            Set ch = co.Chart
            On Error Resume Next
            Do While ch.SeriesCollection.Count > 0
                ch.SeriesCollection(1).Delete
            Loop
            On Error GoTo 0
            ch.ChartType = xlColumnStacked
            ch.HasTitle = True
            ch.ChartTitle.Text = "Portfolio_Group"
        End If
    Else
        ' No category columns present in Daily_Snapshot
        On Error Resume Next
        Set co = GetOrCreateChart(wsDash, "Portfolio_Group")
        co.Chart.ChartTitle.Text = "Portfolio_Group (no categories)"
        On Error GoTo 0
    End If

    Exit Sub
Fail:
    MsgBox "Error (Update_Dashboard): " & Err.Description, vbExclamation
End Sub

Private Function GetOrCreateChart(ws As Worksheet, ByVal chartName As String) As ChartObject
    On Error Resume Next
    Dim co As ChartObject
    Set co = ws.ChartObjects(chartName)
    On Error GoTo 0
    If co Is Nothing Then
        Set co = ws.ChartObjects.Add(Left:=20, Top:=20, Width:=520, Height:=260)
        co.Name = chartName
        co.Chart.ChartType = xlLine
        co.Chart.HasTitle = True
        co.Chart.ChartTitle.Text = chartName
    End If
    Set GetOrCreateChart = co
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
