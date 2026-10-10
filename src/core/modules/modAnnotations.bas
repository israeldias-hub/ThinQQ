Attribute VB_Name = "modAnnotations"
Option Explicit
' Anotações = formas dentro do gráfico (movem com ele). Parâmetros no AlternativeText.
Private mRefreshing As Boolean
Private mLastError As String

Public Function Ann_TryCagr(startVal As Double, endVal As Double, periods As Double, ByRef rate As Double) As Boolean
    rate = 0
    If periods <= 0 Or startVal = 0 Then Exit Function
    If (startVal < 0) <> (endVal < 0) Or endVal = 0 Then Exit Function
    rate = (endVal / startVal) ^ (1 / periods) - 1
    Ann_TryCagr = True
End Function

Public Function Ann_ParseParam(altText As String, key As String) As String
    Dim parts() As String, i As Long, kv() As String
    parts = Split(altText, ";")
    For i = LBound(parts) To UBound(parts)
        kv = Split(parts(i), "=")
        If UBound(kv) = 1 Then
            If kv(0) = key Then Ann_ParseParam = kv(1): Exit Function
        End If
    Next
End Function

' values = número usado nos rótulos (total / End / 1ª série); tops = valor plotado onde a seta encosta.
' Em empilhadas, segmentos negativos entram em values mas são ignorados em tops (a seta encosta no topo positivo).
' Séries mais curtas que n contam como zero; com n = 0 devolve arrays 0..0.
Public Sub Ann_GetCategoryData(cht As Chart, ByRef values() As Double, ByRef tops() As Double)
    Dim kind As String, n As Long, i As Long, s As Series, vals As Variant, v As Double, ends() As Double
    kind = Registry_ChartType(cht)
    n = Geom_ChartCategoryCount(cht)
    If n = 0 Then
        ReDim values(0 To 0): ReDim tops(0 To 0)
        Exit Sub
    End If
    ReDim values(1 To n): ReDim tops(1 To n)
    Select Case kind
        Case "waterfall"
            ends = Waterfall_EndValues(cht)
            For i = 1 To n
                If i <= UBound(ends) Then values(i) = ends(i): tops(i) = ends(i)
            Next
        Case "clustered", "line"
            For i = 1 To n
                v = Ann_SeriesValue(cht.SeriesCollection(1), i)
                values(i) = v: tops(i) = v
            Next
        Case Else
            For Each s In cht.SeriesCollection
                If Not Style_IsHelperSeries(s) Then
                    vals = s.Values
                    For i = 1 To n
                        v = Ann_ArrayValue(vals, i)
                        values(i) = values(i) + v
                        If v > 0 Then tops(i) = tops(i) + v
                    Next
                End If
            Next
            If kind = "stacked100" Then
                For i = 1 To n: tops(i) = 1: Next
            End If
    End Select
End Sub

Private Function Ann_SeriesValue(s As Series, i As Long) As Double
    Ann_SeriesValue = Ann_ArrayValue(s.Values, i)
End Function

Private Function Ann_ArrayValue(vals As Variant, i As Long) As Double
    If Not IsArray(vals) Then Exit Function
    If LBound(vals) + i - 1 > UBound(vals) Then Exit Function
    Ann_ArrayValue = Util_ToDouble(vals(LBound(vals) + i - 1))
End Function

Public Function Ann_AddCagr(cht As Chart, fromIdx As Long, toIdx As Long) As Shape
    Dim nm As String, ln As Shape
    nm = Ann_NewName(cht, "cagr")
    Set ln = cht.Shapes.AddLine(0, 0, 10, 10)
    ln.Name = nm
    ln.AlternativeText = "kind=cagr;from=" & fromIdx & ";to=" & toIdx
    With ln.Line
        .ForeColor.RGB = RGB(89, 89, 89): .Weight = 1.25
        .EndArrowheadStyle = msoArrowheadTriangle
        .EndArrowheadLength = msoArrowheadShort: .EndArrowheadWidth = msoArrowheadNarrow
    End With
    Ann_AddLabel cht, nm
    Ann_Refresh cht
    Set Ann_AddCagr = ln
End Function

Public Function Ann_AddDifference(cht As Chart, fromIdx As Long, toIdx As Long) As Shape
    Dim nm As String, ln As Shape, lvl As Shape
    If Registry_ChartType(cht) = "stacked100" Then Err.Raise vbObjectError + 7001, "Ann_AddDifference", "A seta de diferença não está disponível para gráficos 100%."
    nm = Ann_NewName(cht, "diff")
    Set ln = cht.Shapes.AddLine(0, 0, 0, 10)
    ln.Name = nm
    ln.AlternativeText = "kind=diff;from=" & fromIdx & ";to=" & toIdx
    With ln.Line
        .ForeColor.RGB = RGB(89, 89, 89): .Weight = 1.25
        .BeginArrowheadStyle = msoArrowheadTriangle: .EndArrowheadStyle = msoArrowheadTriangle
        .BeginArrowheadLength = msoArrowheadShort: .EndArrowheadLength = msoArrowheadShort
    End With
    Set lvl = cht.Shapes.AddLine(0, 0, 10, 0)
    lvl.Name = nm & "_lvl"
    lvl.AlternativeText = "kind=level;owner=" & nm
    lvl.Line.ForeColor.RGB = RGB(166, 166, 166): lvl.Line.Weight = 0.75: lvl.Line.DashStyle = msoLineDash
    Ann_AddLabel cht, nm
    Ann_Refresh cht
    Set Ann_AddDifference = ln
End Function

Public Sub Ann_SetWaterfallConnectors(cht As Chart, turnOn As Boolean)
    Dim n As Long, i As Long, shp As Shape
    For i = cht.Shapes.Count To 1 Step -1
        If Left$(cht.Shapes(i).Name, Len(ANN_PREFIX & "conn_")) = ANN_PREFIX & "conn_" Then cht.Shapes(i).Delete
    Next
    Registry_SetOption cht, "connectors", IIf(turnOn, "1", "0")
    If Not turnOn Then Exit Sub
    n = Geom_ChartCategoryCount(cht)
    For i = 1 To n - 1
        Set shp = cht.Shapes.AddLine(0, 0, 10, 0)
        shp.Name = ANN_PREFIX & "conn_" & i
        shp.AlternativeText = "kind=conn;from=" & i & ";to=" & (i + 1)
        shp.Line.ForeColor.RGB = RGB(127, 127, 127): shp.Line.Weight = 0.75
    Next
    Ann_Refresh cht
End Sub

Public Sub Ann_RemoveAll(cht As Chart, Optional includeConnectors As Boolean = False)
    Dim i As Long, nm As String
    For i = cht.Shapes.Count To 1 Step -1
        nm = cht.Shapes(i).Name
        If Left$(nm, Len(ANN_PREFIX)) = ANN_PREFIX Then
            If includeConnectors Or Left$(nm, Len(ANN_PREFIX & "conn_")) <> ANN_PREFIX & "conn_" Then cht.Shapes(i).Delete
        End If
    Next
End Sub

' Reposiciona todas as anotações e reescreve os textos. Seguro para chamar em eventos: cada forma é isolada
' (um erro em uma não impede as outras e fica registrado em Ann_LastError); formas inválidas são ocultadas.
Public Sub Ann_Refresh(cht As Chart)
    Dim values() As Double, tops() As Double, n As Long, i As Long, shp As Shape, kind As String, errs As String
    Dim fmt As String, fromIdx As Long, toIdx As Long, valid As Boolean, chartKind As String, nm As String
    If mRefreshing Then Exit Sub
    mRefreshing = True
    On Error GoTo Done
    n = Geom_ChartCategoryCount(cht)
    If n = 0 Then GoTo Done
    chartKind = Registry_ChartType(cht)
    ' conectores da cascata acompanham o número de categorias (a guarda impede recursão via Ann_SetWaterfallConnectors)
    If chartKind = "waterfall" And Registry_GetOption(cht, "connectors", "0") = "1" Then
        If Ann_CountConnectors(cht) <> n - 1 Then Ann_SetWaterfallConnectors cht, True
    End If
    Ann_GetCategoryData cht, values, tops
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    For i = 1 To cht.Shapes.Count
        Set shp = cht.Shapes(i)
        nm = shp.Name
        If Left$(nm, Len(ANN_PREFIX)) = ANN_PREFIX Then
            kind = Ann_ParseParam(shp.AlternativeText, "kind")
            If kind = "cagr" Or kind = "diff" Or kind = "conn" Then
                fromIdx = Val(Ann_ParseParam(shp.AlternativeText, "from"))
                toIdx = Val(Ann_ParseParam(shp.AlternativeText, "to"))
                valid = (fromIdx >= 1 And toIdx <= n And fromIdx < toIdx)
                If kind = "diff" And chartKind = "stacked100" Then valid = False
                Ann_SetShapeVisible cht, nm, valid
                If valid Then
                    On Error Resume Next
                    Ann_LayoutOne cht, shp, kind, fromIdx, toIdx, values, tops, fmt
                    If Err.Number <> 0 Then
                        mLastError = "Ann_Refresh " & nm & " " & Err.Number & " " & Err.Description
                        Errors_Log mLastError
                        Err.Clear
                    End If
                    On Error GoTo Done
                End If
            End If
        End If
    Next
Done:
    mRefreshing = False
End Sub

' Último erro registrado por Ann_Refresh (vazio se nenhum). Substituível por Errors_Log quando existir.
Public Function Ann_LastError() As String
    Ann_LastError = mLastError
End Function

' ---------- privados ----------

Private Function Ann_NewName(cht As Chart, kind As String) As String
    Dim k As Long, nm As String
    Do
        k = k + 1
        nm = ANN_PREFIX & kind & "_" & k
    Loop While Not Ann_Shape(cht, nm) Is Nothing Or Not Ann_Shape(cht, nm & "_lbl") Is Nothing Or Not Ann_Shape(cht, nm & "_lvl") Is Nothing
    Ann_NewName = nm
End Function

Private Function Ann_Shape(cht As Chart, nm As String) As Shape
    On Error Resume Next
    Set Ann_Shape = cht.Shapes(nm)
    On Error GoTo 0
End Function

Private Sub Ann_AddLabel(cht As Chart, ownerName As String)
    Dim lbl As Shape
    Set lbl = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 70, 16)
    lbl.Name = ownerName & "_lbl"
    lbl.AlternativeText = "kind=label;owner=" & ownerName
    With lbl.TextFrame2
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeShapeToFitText
        .MarginLeft = 3: .MarginRight = 3: .MarginTop = 1: .MarginBottom = 1
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(64, 64, 64)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    lbl.Fill.Visible = msoTrue
    lbl.Fill.ForeColor.RGB = RGB(255, 255, 255)
    lbl.Line.Visible = msoTrue
    lbl.Line.ForeColor.RGB = RGB(166, 166, 166)
    lbl.Line.Weight = 0.5
End Sub

' Move uma linha para ir de (x1,y1) a (x2,y2), ajustando os flips para manter a ponta da seta no fim.
Private Sub Ann_MoveLine(ln As Shape, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    ln.Left = IIf(x1 < x2, x1, x2)
    ln.Top = IIf(y1 < y2, y1, y2)
    ln.Width = Abs(x2 - x1)
    ln.Height = Abs(y2 - y1)
    If (ln.HorizontalFlip = msoTrue) <> (x2 < x1) Then ln.Flip msoFlipHorizontal
    If (ln.VerticalFlip = msoTrue) <> (y2 < y1) Then ln.Flip msoFlipVertical
End Sub

Private Function Ann_SignedFormat(fmt As String) As String
    Dim sec As String
    sec = Split(fmt & ";", ";")(0)
    Ann_SignedFormat = "+" & sec & ";-" & sec & ";0"
End Function

Private Sub Ann_LayoutCagr(cht As Chart, ln As Shape, fromIdx As Long, toIdx As Long, values() As Double, tops() As Double, fmt As String)
    Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double, lbl As Shape, rate As Double, txt As String
    Dim body As Range, periods As Double
    x1 = Geom_ChartCategoryCenterX(cht, fromIdx): x2 = Geom_ChartCategoryCenterX(cht, toIdx)
    y1 = Geom_ChartValueToY(cht, tops(fromIdx)) - 10
    y2 = Geom_ChartValueToY(cht, tops(toIdx)) - 10
    Ann_MoveLine ln, x1, y1, x2, y2
    Set body = Block_Body(Registry_Block(cht))
    periods = Geom_Periods(body.Cells(fromIdx + 1, 1).Value, body.Cells(toIdx + 1, 1).Value, fromIdx, toIdx)
    If Ann_TryCagr(values(fromIdx), values(toIdx), periods, rate) Then
        txt = "CAGR " & Format$(rate, "+0.0%;-0.0%;0.0%")
    Else
        txt = "CAGR n/d"
    End If
    Set lbl = Ann_Shape(cht, ln.Name & "_lbl")
    If lbl Is Nothing Then Exit Sub
    lbl.TextFrame2.TextRange.Text = txt
    lbl.Left = (x1 + x2) / 2 - lbl.Width / 2
    lbl.Top = IIf(y1 < y2, y1, y2) - lbl.Height - 4
    Ann_ClampShape cht, lbl
End Sub

Private Sub Ann_LayoutDiff(cht As Chart, ln As Shape, fromIdx As Long, toIdx As Long, values() As Double, tops() As Double, fmt As String)
    Dim xFrom As Double, xTo As Double, yFrom As Double, yTo As Double, xArrow As Double, half As Double
    Dim lvl As Shape, lbl As Shape, diff As Double, txt As String
    half = Geom_ChartBarHalfWidth(cht)
    xFrom = Geom_ChartCategoryCenterX(cht, fromIdx): xTo = Geom_ChartCategoryCenterX(cht, toIdx)
    yFrom = Geom_ChartValueToY(cht, tops(fromIdx)): yTo = Geom_ChartValueToY(cht, tops(toIdx))
    xArrow = xTo + half + 10
    If xArrow > cht.ChartArea.Width - 2 Then xArrow = cht.ChartArea.Width - 2
    Set lvl = Ann_Shape(cht, ln.Name & "_lvl")
    If Not lvl Is Nothing Then Ann_MoveLine lvl, xFrom + half, yFrom, xArrow + 4, yFrom
    Ann_MoveLine ln, xArrow, yFrom, xArrow, yTo
    diff = values(toIdx) - values(fromIdx)
    txt = Format$(diff, Ann_SignedFormat(fmt))
    If values(fromIdx) <> 0 Then txt = txt & " (" & Format$(diff / Abs(values(fromIdx)), "+0%;-0%;0%") & ")"
    Set lbl = Ann_Shape(cht, ln.Name & "_lbl")
    If lbl Is Nothing Then Exit Sub
    lbl.TextFrame2.TextRange.Text = txt
    lbl.Left = xArrow + 6
    lbl.Top = (yFrom + yTo) / 2 - lbl.Height / 2
    Ann_ClampShape cht, lbl
End Sub

Private Sub Ann_LayoutConn(cht As Chart, shp As Shape, i As Long, j As Long, tops() As Double)
    Dim half As Double, y As Double
    half = Geom_ChartBarHalfWidth(cht)
    y = Geom_ChartValueToY(cht, tops(i))
    Ann_MoveLine shp, Geom_ChartCategoryCenterX(cht, i) + half, y, Geom_ChartCategoryCenterX(cht, j) - half, y
End Sub

Private Sub Ann_LayoutOne(cht As Chart, shp As Shape, kind As String, fromIdx As Long, toIdx As Long, values() As Double, tops() As Double, fmt As String)
    Select Case kind
        Case "cagr": Ann_LayoutCagr cht, shp, fromIdx, toIdx, values, tops, fmt
        Case "diff": Ann_LayoutDiff cht, shp, fromIdx, toIdx, values, tops, fmt
        Case "conn": Ann_LayoutConn cht, shp, fromIdx, toIdx, tops
    End Select
End Sub

' Mostra/oculta a forma e seus companheiros (_lbl, _lvl).
Private Sub Ann_SetShapeVisible(cht As Chart, nm As String, vis As Boolean)
    Dim sfx As Variant, s As Shape, st As Long
    st = IIf(vis, msoTrue, msoFalse)
    On Error Resume Next
    For Each sfx In Array("", "_lbl", "_lvl")
        Set s = Nothing
        Set s = cht.Shapes(nm & sfx)
        If Not s Is Nothing Then
            If s.Visible <> st Then s.Visible = st
        End If
    Next
    On Error GoTo 0
End Sub

Private Function Ann_CountConnectors(cht As Chart) As Long
    Dim i As Long, p As String
    p = ANN_PREFIX & "conn_"
    For i = 1 To cht.Shapes.Count
        If Left$(cht.Shapes(i).Name, Len(p)) = p Then Ann_CountConnectors = Ann_CountConnectors + 1
    Next
End Function

' Mantém a forma dentro da área do gráfico (esquerda/topo têm prioridade quando não cabe).
Private Sub Ann_ClampShape(cht As Chart, shp As Shape)
    Dim w As Double, h As Double
    w = cht.ChartArea.Width: h = cht.ChartArea.Height
    If shp.Left + shp.Width > w Then shp.Left = w - shp.Width
    If shp.Top + shp.Height > h Then shp.Top = h - shp.Height
    If shp.Left < 0 Then shp.Left = 0
    If shp.Top < 0 Then shp.Top = 0
End Sub
