Attribute VB_Name = "modChartStacked"
Option Explicit
Private Const TOTAL_LABEL_OFFSET As Double = 6      ' o rotulo "Right" de um ponto XY comeca ~5,5 pt depois do ponto (medido)
Private Const TOTAL_LABEL_MARGIN As Double = 4      ' respiro entre o fim do rotulo do total e a borda livre
' Colunas e barras empilhadas / agrupadas / 100% / linha (kinds stacked, clustered, stacked100, line, barstacked,
' barclustered, barstacked100).
' Bloco (corpo): linha 1 = cabeçalhos; colunas: 1 Categoria | 2 Total | 3 TotalPos | (só 100%) 4..3+S Participação de cada série |
' P = Pos | P+1 = VLine (as duas últimas, em todos os tipos da família; P = Stacked_PosCol, P+1 = Stacked_VLineCol).
' TotalPos = posição plotada do rótulo do total: em empilhadas é a extensão positiva da pilha (soma dos valores > 0;
' o RÓTULO continua mostrando o total líquido da coluna 2; sem valores positivos, a extensão negativa); em 100% é 1.
' Participação = valor / total da linha (0 se o total é 0 ou a célula não é número).
' Pos = posição da categoria no eixo Y secundário das auxiliares XY das barras (n - i + 0,5: a primeira categoria no topo).
' VLine = média da coluna Total (linha de valor; #N/D em 100%, onde não há linha).

Public Function Stacked_BuildBlock(wb As Workbook, id As String, data As clsChartData, kind As String) As Range
    Dim blk As Range, body As Range, i As Long, j As Long, rowCells As Range, srcAddr As String, totAddr As String
    Dim is100 As Boolean, n As Long, cols As Long, posCol As Long, vCol As Long, totRange As String
    is100 = Kinds_Is100(kind)
    n = data.CategoryCount
    cols = 3 + IIf(is100, data.SeriesCount, 0) + 2
    Set blk = CalcSheet_AllocateBlock(wb, id, kind, Mid$(ChartData_ExternalAddress(data.Source), 2), n + 1, cols)
    Set body = Block_Body(blk)
    posCol = Stacked_PosCol(blk): vCol = Stacked_VLineCol(blk)
    body.Cells(1, 1).Value = "Categoria": body.Cells(1, 2).Value = "Total": body.Cells(1, 3).Value = "TotalPos"
    If is100 Then
        For j = 1 To data.SeriesCount: body.Cells(1, 3 + j).Value = "Part" & j: Next
    End If
    body.Cells(1, posCol).Value = "Pos": body.Cells(1, vCol).Value = "VLine"
    totRange = body.Cells(2, 2).Resize(n, 1).Address(True, True)     ' mesma planilha: sem nome de planilha
    For i = 1 To n
        body.Cells(i + 1, 1).Formula = ChartData_ExternalAddress(data.Source.Cells(i + 1, 1))
        Set rowCells = data.Source.Cells(i + 1, 2).Resize(1, data.SeriesCount)
        body.Cells(i + 1, 2).Formula = "=AGGREGATE(9,6," & Mid$(ChartData_ExternalAddress(rowCells), 2) & ")"   ' ignora erros (#N/A) e texto
        If is100 Then
            body.Cells(i + 1, 3).Formula = "=1"
            totAddr = body.Cells(i + 1, 2).Address(False, False)
            For j = 1 To data.SeriesCount
                srcAddr = Mid$(ChartData_ExternalAddress(data.Source.Cells(i + 1, 1 + j)), 2)
                body.Cells(i + 1, 3 + j).Formula = "=IF(AND(ISNUMBER(" & srcAddr & ")," & totAddr & "<>0)," & srcAddr & "/" & totAddr & ",0)"
            Next
        ElseIf kind = "clustered" Or kind = "barclustered" Or kind = "line" Then
            body.Cells(i + 1, 3).Formula = "=" & body.Cells(i + 1, 2).Address(False, False)
        Else
            ' empilhada: posição do total = extensão positiva da pilha (com sinais mistos o total líquido ficaria dentro da pilha)
            srcAddr = Mid$(ChartData_ExternalAddress(rowCells), 2)
            body.Cells(i + 1, 3).Formula = "=IF(SUMIF(" & srcAddr & ","">0"")>0,SUMIF(" & srcAddr & ","">0""),SUMIF(" & srcAddr & ",""<0""))"
        End If
        body.Cells(i + 1, posCol).Formula = "=" & (n - i) & ".5"           ' independe do separador decimal da localidade
        If is100 Then
            body.Cells(i + 1, vCol).Formula = "=NA()"
        Else
            body.Cells(i + 1, vCol).Formula = "=AVERAGE(" & totRange & ")"
        End If
    Next
    Set Stacked_BuildBlock = blk
End Function

' Coluna Pos (posição da categoria) e coluna VLine (linha de valor): as duas últimas do bloco.
Public Function Stacked_PosCol(blk As Range) As Long
    Stacked_PosCol = blk.Columns.Count - 1
End Function

Public Function Stacked_VLineCol(blk As Range) As Long
    Stacked_VLineCol = blk.Columns.Count
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
        Case "barstacked": cht.ChartType = xlBarStacked
        Case "barclustered": cht.ChartType = xlBarClustered
        Case "barstacked100": cht.ChartType = xlBarStacked100
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
    Registry_SetOption cht, "n", CStr(data.CategoryCount)
    If kind <> "line" Then
        cht.ChartGroups(1).GapWidth = IIf(kind = "clustered" Or kind = "barclustered", 80, 50)
        Registry_SetOption cht, "gap", IIf(kind = "clustered" Or kind = "barclustered", "80", "50")
    End If
    Style_ApplySeriesColors cht, "padrao"
    Stacked_SetLabels cht, True
    Stacked_SetTotals cht, Kinds_Supports(kind, "totals")
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
    Dim s As Series, blk As Range, body As Range, n As Long, fmt As String, kind As String
    kind = Registry_ChartType(cht)
    Set s = Chart_FindSeries(cht, "tqq_total")
    If turnOn Then
        If Not Kinds_Supports(kind, "totals") Then Exit Sub
    End If
    If Not turnOn Then
        If Not s Is Nothing Then s.Delete
        Registry_SetOption cht, "totals", "0"
        Exit Sub
    End If
    Set blk = Registry_Block(cht)
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    If s Is Nothing Then
        If Kinds_IsHorizontal(kind) Then
            ' barras: auxiliar XY no grupo secundário, X = TotalPos (fim da pilha) e Y = Pos (centro da faixa da categoria)
            Set s = Chart_AddHelperSeries(cht, "tqq_total", body.Cells(2, 3).Resize(n, 1), body.Cells(2, Stacked_PosCol(blk)).Resize(n, 1))
        Else
            Set s = Chart_AddHelperSeries(cht, "tqq_total", body.Cells(2, 3).Resize(n, 1))
        End If
    End If
    Chart_HideHelperLegendEntries cht
    fmt = Labels_TotalFormat(cht)
    Labels_LinkPoints s, body.Cells(2, 2).Resize(n, 1), fmt, , , Kinds_IsHorizontal(kind)
    Labels_RefreshPositions cht                    ' acima/abaixo (barras: direita/esquerda) pela posição plotada (TotalPos)
    s.DataLabels.Font.Bold = True
    If Kinds_IsHorizontal(kind) Then Stacked_EnsureTotalsRoom cht, body.Cells(2, 3).Resize(n, 1)
    Registry_SetOption cht, "totals", "1"
End Sub

' Barras: garante que o rótulo do total mais à direita caiba antes da borda livre (a do gráfico ou, com a legenda à
' direita, a borda esquerda dela). O rótulo começa TOTAL_LABEL_OFFSET depois do fim da barra e ocupa a sua largura
' (estimada pelo texto: Labels_MaxTotalWidth). Em 100% (a barra termina no fim do eixo) e com a maior barra perto do
' máximo do eixo a folga automática não basta e o rótulo ficaria sobre a legenda. Só reduz a área de plotagem (nunca
' a aumenta); 3 passadas porque a largura interna nem sempre acompanha 1:1 a externa. Nunca propaga erro. A folga vale
' para a largura dos rótulos e a escala do momento (criação do gráfico e religar os totais).
Private Sub Stacked_EnsureTotalsRoom(cht As Chart, totalPos As Range)
    Dim k As Long, best As Double, mx As Double, need As Double, delta As Double
    On Error Resume Next
    best = Labels_MaxTotalWidth(cht)
    mx = Application.WorksheetFunction.Max(totalPos)
    If Err.Number <> 0 Or best <= 0 Or mx <= 0 Then Err.Clear: Exit Sub
    need = TOTAL_LABEL_OFFSET + best + TOTAL_LABEL_MARGIN
    For k = 1 To 3
        delta = need - (Geom_ChartRightEdge(cht) - Geom_ChartValCoord(cht, mx))
        If delta <= 0.01 Then Exit For
        cht.PlotArea.Width = cht.PlotArea.Width - delta
    Next
    If Err.Number <> 0 Then Errors_Log "Stacked_EnsureTotalsRoom" & vbTab & Err.Number & vbTab & Err.Description
End Sub

Public Sub Stacked_SetLabels(cht As Chart, turnOn As Boolean)
    Dim s As Series, fmt As String, j As Long, body As Range, n As Long, is100 As Boolean
    fmt = Labels_SegmentFormat(cht)
    is100 = Kinds_Is100(Registry_ChartType(cht))
    If is100 Then Set body = Block_Body(Registry_Block(cht)): n = body.Rows.Count - 1
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            j = j + 1
            s.HasDataLabels = turnOn
            If turnOn Then
                With s.DataLabels
                    .ShowValue = True: .ShowSeriesName = False: .ShowCategoryName = False
                    Style_AssignNumberFormat s.DataLabels, fmt
                    .Font.Size = 9
                    If s.ChartType = xlLine Or s.ChartType = xlLineMarkers Then
                        .Position = xlLabelPositionAbove
                        .Font.Color = RGB(64, 64, 64)
                    ElseIf s.ChartType = xlColumnClustered Or s.ChartType = xlBarClustered Then
                        .Position = xlLabelPositionOutsideEnd
                        .Font.Color = RGB(64, 64, 64)
                    Else
                        .Position = xlLabelPositionCenter
                        .Font.Color = Style_ContrastFont(s.Format.Fill.ForeColor.RGB)
                    End If
                End With
                ' 100%: o rótulo do segmento é a participação (coluna de participação do bloco), não o valor absoluto
                If is100 Then Labels_LinkPoints s, body.Cells(2, 3 + j).Resize(n, 1), fmt, xlLabelPositionCenter, Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB))
            End If
        End If
    Next
    Registry_SetOption cht, "labels", IIf(turnOn, "1", "0")
End Sub

' Linhas de série (conectores entre segmentos), em colunas e barras empilhadas e 100%. Grupos sem suporte
' (linha, agrupadas, a auxiliar XY das barras) são ignorados.
Public Sub Stacked_SetConnectors(cht As Chart, turnOn As Boolean)
    Dim g As ChartGroup, accepted As Boolean
    On Error Resume Next
    For Each g In cht.ChartGroups
        Err.Clear
        If IsStackedGroup(g) Then g.HasSeriesLines = turnOn Else Err.Raise 5
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

' Esconde da legenda as entradas das séries auxiliares (tqq_total e as demais tqq_*); idempotente. A lógica mora em
' Chart_HideHelperLegendEntries (modChartUtil), que também cobre várias auxiliares e grupos secundários.
Public Sub Stacked_HideTotalLegendEntry(cht As Chart)
    Chart_HideHelperLegendEntries cht
End Sub

Private Function IsStackedGroup(g As ChartGroup) As Boolean
    Dim ct As Long
    If g.SeriesCollection.Count = 0 Then Exit Function
    ct = g.SeriesCollection(1).ChartType
    IsStackedGroup = (ct = xlColumnStacked Or ct = xlColumnStacked100 Or ct = xlBarStacked Or ct = xlBarStacked100)
End Function