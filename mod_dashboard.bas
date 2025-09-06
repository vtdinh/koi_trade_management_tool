Attribute VB_Name = "mod_dashboard"
Option Explicit

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
    Dim datesArr() As Variant, navArr() As Double
    Dim count As Long: count = 0

    Dim r As Long, dt As Date
    For r = 2 To lastR
        If IsDate(wsSnap.Cells(r, 1).Value) Then
            dt = DateValue(wsSnap.Cells(r, 1).Value)
            If dt >= dStart And dt <= dEnd Then
                count = count + 1
                ReDim Preserve datesArr(1 To count)
                ReDim Preserve navArr(1 To count)
                datesArr(count) = dt
                If IsNumeric(wsSnap.Cells(r, 4).Value) Then
                    navArr(count) = CDbl(wsSnap.Cells(r, 4).Value)
                Else
                    navArr(count) = 0#
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

    ' Compute drawdowns and annotate
    Dim mddPct As Double: mddPct = ComputeMaxDrawdownPct(navArr)
    Dim curDDPct As Double: curDDPct = ComputeCurrentDrawdownPct(navArr)
    AnnotateDrawdown ch, mddPct, curDDPct

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
        ComputeCurrentDrawdownPct = (peak - vals(UBound(vals))) / peak
    Else
        ComputeCurrentDrawdownPct = 0#
    End If
    Exit Function
Done:
    ComputeCurrentDrawdownPct = 0#
End Function

Private Sub AnnotateDrawdown(ch As Chart, ByVal mddPct As Double, ByVal curDDPct As Double)
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
