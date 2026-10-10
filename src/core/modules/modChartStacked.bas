Attribute VB_Name = "modChartStacked"
Option Explicit
' Colunas empilhadas / agrupadas / 100% / linha.
' Bloco (corpo): linha 1 = cabeçalhos; colunas: 1 Categoria | 2 Total | 3 TotalPos (=Total, ou 1 no 100%).

Public Function Stacked_BuildBlock(wb As Workbook, id As String, data As clsChartData, kind As String) As Range
    Dim blk As Range, body As Range, i As Long, rowCells As Range
    Set blk = CalcSheet_AllocateBlock(wb, id, kind, Mid$(ChartData_ExternalAddress(data.Source), 2), data.CategoryCount + 1, 3)
    Set body = Block_Body(blk)
    body.Cells(1, 1).Value = "Categoria": body.Cells(1, 2).Value = "Total": body.Cells(1, 3).Value = "TotalPos"
    For i = 1 To data.CategoryCount
        body.Cells(i + 1, 1).Formula = ChartData_ExternalAddress(data.Source.Cells(i + 1, 1))
        Set rowCells = data.Source.Cells(i + 1, 2).Resize(1, data.SeriesCount)
        body.Cells(i + 1, 2).Formula = "=AGGREGATE(9,6," & Mid$(ChartData_ExternalAddress(rowCells), 2) & ")"   ' ignora erros (#N/A) e texto
        If kind = "stacked100" Then
            body.Cells(i + 1, 3).Formula = "=1"
        Else
            body.Cells(i + 1, 3).Formula = "=" & body.Cells(i + 1, 2).Address(False, False)
        End If
    Next
    Set Stacked_BuildBlock = blk
End Function

Public Function Stacked_Insert(ws As Worksheet, data As clsChartData, kind As String) As ChartObject
    Dim id As String, co As ChartObject, cht As Chart, s As Series, j As Long, anchor As Range
    Dim en As Long, es As String, ed As String
    On Error GoTo Fail
    id = Registry_NewId()
    Stacked_BuildBlock ws.Parent, id, data, kind
    Util_TestHookAfterBlock
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, Chart_FreeTop(ws, anchor.Left, anchor.Top, 460, 280), 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    For j = 1 To data.SeriesCount
        Set s = cht.SeriesCollection.NewSeries
        s.Name = ChartData_ExternalAddress(data.SeriesNameCell(j))
        s.Values = ChartData_ExternalAddress(data.SeriesValueCells(j))
        s.XValues = ChartData_ExternalAddress(data.CategoryCells())
    Next
    Select Case kind
        Case "clustered": cht.ChartType = xlColumnClustered
        Case "stacked100": cht.ChartType = xlColumnStacked100
        Case "line": cht.ChartType = xlLineMarkers
        Case Else: cht.ChartType = xlColumnStacked
    End Select
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlZero
    Style_ApplyBase cht
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionRight
    cht.Legend.Font.Size = 9
    Registry_SetOption cht, "numfmt", "int"
    Registry_SetOption cht, "palette", "padrao"
    If kind <> "line" Then
        cht.ChartGroups(1).GapWidth = IIf(kind = "clustered", 80, 50)
        Registry_SetOption cht, "gap", IIf(kind = "clustered", "80", "50")
    End If
    Style_ApplySeriesColors cht, "padrao"
    Stacked_SetLabels cht, True
    Stacked_SetTotals cht, (kind = "stacked" Or kind = "stacked100")
    Stacked_SetConnectors cht, False
    Set Stacked_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

Public Sub Stacked_SetTotals(cht As Chart, turnOn As Boolean)
    Dim s As Series, body As Range, n As Long, fmt As String
    Set s = Chart_FindSeries(cht, "tqq_total")
    If turnOn Then
        If Registry_ChartType(cht) <> "stacked" And Registry_ChartType(cht) <> "stacked100" Then Exit Sub
    End If
    If Not turnOn Then
        If Not s Is Nothing Then s.Delete
        Registry_SetOption cht, "totals", "0"
        Exit Sub
    End If
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    If s Is Nothing Then
        Set s = cht.SeriesCollection.NewSeries
        s.Name = "tqq_total"
        s.Values = "='" & CALC_SHEET & "'!" & body.Cells(2, 3).Resize(n, 1).Address(True, True)
        s.ChartType = xlLine
        s.Format.Line.Visible = msoFalse
        s.MarkerStyle = xlMarkerStyleNone
    End If
    Stacked_HideTotalLegendEntry cht
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    Labels_LinkPoints s, body.Cells(2, 2).Resize(n, 1), fmt
    s.DataLabels.Font.Bold = True
    Registry_SetOption cht, "totals", "1"
End Sub

Public Sub Stacked_SetLabels(cht As Chart, turnOn As Boolean)
    Dim s As Series, fmt As String
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            s.HasDataLabels = turnOn
            If turnOn Then
                With s.DataLabels
                    .ShowValue = True: .ShowSeriesName = False: .ShowCategoryName = False
                    Style_AssignNumberFormat s.DataLabels, fmt
                    .Font.Size = 9
                    If s.ChartType = xlLine Or s.ChartType = xlLineMarkers Then
                        .Position = xlLabelPositionAbove
                        .Font.Color = RGB(64, 64, 64)
                    ElseIf s.ChartType = xlColumnClustered Then
                        .Position = xlLabelPositionOutsideEnd
                        .Font.Color = RGB(64, 64, 64)
                    Else
                        .Position = xlLabelPositionCenter
                        .Font.Color = Style_ContrastFont(s.Format.Fill.ForeColor.RGB)
                    End If
                End With
            End If
        End If
    Next
    Registry_SetOption cht, "labels", IIf(turnOn, "1", "0")
End Sub

' Linhas de série (conectores entre segmentos). Grupos sem suporte (linha/agrupadas) são ignorados.
Public Sub Stacked_SetConnectors(cht As Chart, turnOn As Boolean)
    Dim g As ChartGroup, accepted As Boolean
    On Error Resume Next
    For Each g In cht.ChartGroups
        Err.Clear
        If IsStackedColumnGroup(g) Then g.HasSeriesLines = turnOn Else Err.Raise 5
        If Err.Number = 0 Then
            If turnOn Then
                If g.HasSeriesLines Then
                    accepted = True
                    g.SeriesLines.Format.Line.ForeColor.RGB = RGB(166, 166, 166)
                    g.SeriesLines.Format.Line.Weight = 0.75
                End If
            End If
        End If
    Next
    On Error GoTo 0
    Registry_SetOption cht, "connectors", IIf(turnOn And accepted, "1", "0")
End Sub

' Esconde da legenda a entrada da série auxiliar tqq_total. A posição da entrada é calculada pela ordem
' dos grupos de gráfico (colunas antes de linhas) e das séries dentro do grupo. Idempotente: se há menos
' entradas que séries, a entrada já foi removida e nada é feito.
Public Sub Stacked_HideTotalLegendEntry(cht As Chart)
    Dim g As ChartGroup, k As Long, idx As Long
    If Not cht.HasLegend Then Exit Sub
    If Chart_FindSeries(cht, "tqq_total") Is Nothing Then Exit Sub
    If cht.Legend.LegendEntries.Count < cht.SeriesCollection.Count Then Exit Sub
    For Each g In cht.ChartGroups
        For k = 1 To g.SeriesCollection.Count
            idx = idx + 1
            If g.SeriesCollection(k).Name = "tqq_total" Then
                cht.Legend.LegendEntries(idx).Delete
                Exit Sub
            End If
        Next
    Next
End Sub

Private Function IsStackedColumnGroup(g As ChartGroup) As Boolean
    Dim ct As Long
    If g.SeriesCollection.Count = 0 Then Exit Function
    ct = g.SeriesCollection(1).ChartType
    IsStackedColumnGroup = (ct = xlColumnStacked Or ct = xlColumnStacked100)
End Function