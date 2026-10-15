Attribute VB_Name = "modDelta"
Option Explicit
' Delta entre duas barras escolhidas com o mouse (Parte B, spec 2026-10-14 secao 5). Referencia "k:c": k = ordinal da serie de DADOS (as
' auxiliares tqq_* nao contam) e c = categoria; nas cascatas k = 0 (a barra da categoria); na cascata comparativa k = lado (1 = A, 2 = B) e
' c = etapa. Valor (texto) e extensao (lo..hi, onde a seta encosta) por familia: agrupadas e linha = o ponto; empilhadas e combinacao = o
' segmento (positivos sobem e negativos descem na ordem das series); cascata = End e Lo/Hi do bloco; comparativa = Val e Lo/Hi do lado.
Public Const DELTA_GAP As Double = 8            ' pt: da borda mais externa das barras do trecho ate o rotulo
Public Const DELTA_LABEL_ROOM As Double = 18    ' pt a mais com Rotulos ou Totais ligados (como ANN_TOTALS_EXTRA)
Public Const DELTA_BAD_PICK As String = "Escolha duas barras diferentes do gráfico."

' Texto do delta de a para b: "both" = "+5 (+50%)" (como a Seta diferenca), "abs" = "+5", "pct" = "+50%". Percentual (b - a) / |a|; a = 0:
' "both" mostra so o valor e "pct" mostra "n/d". fmt = codigo do formato do grafico (Style_NumberFormat; a 1a secao). Pura.
Public Function Delta_Text(a As Double, b As Double, mode As String, fmt As String) As String
    Dim d As Double, absTxt As String, pctTxt As String
    d = b - a
    absTxt = Delta_FormatSigned(d, fmt)
    If a <> 0 Then pctTxt = Format$(d / Abs(a), "+0%;-0%;0%") Else pctTxt = "n/d"
    Select Case mode
        Case "abs": Delta_Text = absTxt
        Case "pct": Delta_Text = pctTxt
        Case Else
            Delta_Text = absTxt
            If a <> 0 Then Delta_Text = absTxt & " (" & pctTxt & ")"
    End Select
End Function

' Valor com sinal ("+5", "-5", "0") na 1a secao do formato fmt (codigo de Style_NumberFormat). Virgulas finais = escala do Excel (mil:
' "#,##0,"; milhoes: "#,##0.0,,"): o Format$ do VBA nao as entende, entao cada uma divide por 1000 e sai do formato. Usada pelo delta e pela
' Seta diferenca (Ann_LayoutDiff). Pura.
Public Function Delta_FormatSigned(d As Double, fmt As String) As String
    Dim sec As String, shown As Double
    shown = d
    sec = Split(fmt & ";", ";")(0)
    Do While Len(sec) > 1 And Right$(sec, 1) = ","
        sec = Left$(sec, Len(sec) - 1)
        shown = shown / 1000
    Loop
    Delta_FormatSigned = Format$(shown, "+" & sec & ";-" & sec & ";0")
End Function

' Texto do rotulo de valor (sem sinal forcado) de v no formato fmt, com a mesma escala das virgulas finais: estimativa da largura do rotulo
' de fora nas barras horizontais (Ann_DeltaPrepare). Pura.
Public Function Delta_FormatPlain(v As Double, fmt As String) As String
    Dim sec As String, shown As Double
    shown = v
    sec = Split(fmt & ";", ";")(0)
    Do While Len(sec) > 1 And Right$(sec, 1) = ","
        sec = Left$(sec, Len(sec) - 1)
        shown = shown / 1000
    Loop
    Delta_FormatPlain = Format$(shown, sec & ";-" & sec & ";0")
End Function

' Ciclo do botao "Delta: valor / % / ambos": ambos -> valor -> % -> ambos. Pura.
Public Function Delta_NextFormat(fmt As String) As String
    Select Case fmt
        Case "both": Delta_NextFormat = "abs"
        Case "abs": Delta_NextFormat = "pct"
        Case "pct": Delta_NextFormat = "both"
        Case Else: Delta_NextFormat = "abs"
    End Select
End Function

' "k:c" com k >= 0 e c >= 1, so algarismos (o texto alternativo e editavel: sinais, "1e2", "&H1", separadores e numeros grandes demais
' para Long sao recusados sem erro). Pura.
Public Function Delta_ParseRef(ref As String, ByRef k As Long, ByRef c As Long) As Boolean
    Dim p() As String
    p = Split(ref, ":")
    If UBound(p) <> 1 Then Exit Function
    If Not Delta_IsDigits(p(0)) Or Not Delta_IsDigits(p(1)) Then Exit Function
    k = CLng(p(0)): c = CLng(p(1))
    Delta_ParseRef = (k >= 0 And c >= 1)
End Function

Private Function Delta_IsDigits(t As String) As Boolean
    Dim i As Long
    If Len(t) = 0 Or Len(t) > 9 Then Exit Function
    For i = 1 To Len(t)
        If Mid$(t, i, 1) < "0" Or Mid$(t, i, 1) > "9" Then Exit Function
    Next
    Delta_IsDigits = True
End Function

' Series de colunas/barras (as de linha e XY nao sao barras).
Public Function Delta_IsColumnSeries(s As Series) As Boolean
    Select Case s.ChartType
        Case xlColumnClustered, xlColumnStacked, xlColumnStacked100, xlBarClustered, xlBarStacked, xlBarStacked100
            Delta_IsColumnSeries = True
    End Select
End Function

Public Function Delta_DataSeriesCount(cht As Chart) As Long
    Dim i As Long
    For i = 1 To cht.SeriesCollection.Count
        If Not Style_IsHelperSeries(cht.SeriesCollection(i)) Then Delta_DataSeriesCount = Delta_DataSeriesCount + 1
    Next
End Function

' Indice em SeriesCollection da k-esima serie de dados (0 se nao existe).
Public Function Delta_SeriesIndex(cht As Chart, k As Long) As Long
    Dim i As Long, n As Long
    For i = 1 To cht.SeriesCollection.Count
        If Not Style_IsHelperSeries(cht.SeriesCollection(i)) Then
            n = n + 1
            If n = k Then Delta_SeriesIndex = i: Exit Function
        End If
    Next
End Function

' Ordinal de dados da serie de indice s (0 = auxiliar ou inexistente).
Public Function Delta_DataOrdinal(cht As Chart, s As Long) As Long
    Dim i As Long, n As Long
    If s < 1 Or s > cht.SeriesCollection.Count Then Exit Function
    If Style_IsHelperSeries(cht.SeriesCollection(s)) Then Exit Function
    For i = 1 To s
        If Not Style_IsHelperSeries(cht.SeriesCollection(i)) Then n = n + 1
    Next
    Delta_DataOrdinal = n
End Function

' Indice da serie pelo nome (0 se nao existe). Series.Name devolve o texto.
Public Function Delta_FindSeriesByName(cht As Chart, nm As String) As Long
    Dim i As Long
    For i = 1 To cht.SeriesCollection.Count
        If cht.SeriesCollection(i).Name = nm Then Delta_FindSeriesByName = i: Exit Function
    Next
End Function

Private Function Delta_PointValue(s As Series, c As Long) As Double
    Dim vals As Variant
    vals = s.Values
    If Not IsArray(vals) Then Exit Function
    If LBound(vals) + c - 1 > UBound(vals) Then Exit Function
    Delta_PointValue = Util_ToDouble(vals(LBound(vals) + c - 1))
End Function

Private Function Delta_IsStackedKind(kind As String) As Boolean
    Delta_IsStackedKind = (kind = "stacked" Or kind = "barstacked" Or kind = "combo")
End Function

' Clique do Excel (Arg1 = indice da serie, Arg2 = ponto) -> referencia. False: tipo sem o recurso, auxiliar que nao e barra, vao da
' comparativa, ponto fora das categorias.
Public Function Delta_ResolvePick(cht As Chart, s As Long, p As Long, ByRef ref As String) As Boolean
    Dim kind As String, k As Long
    kind = Registry_ChartType(cht)
    If Not Kinds_Supports(kind, "delta") Then Exit Function
    If s < 1 Or s > cht.SeriesCollection.Count Or p < 1 Then Exit Function
    Select Case Kinds_Family(kind)
        Case "waterfall"
            If Not Delta_IsColumnSeries(cht.SeriesCollection(s)) Then Exit Function
            If p > Geom_ChartCategoryCount(cht) Then Exit Function
            ref = "0:" & p
        Case "waterfallpair"
            If Not Delta_IsColumnSeries(cht.SeriesCollection(s)) Then Exit Function      ' setas, conectores, rotulos e nomes (XY/linhas)
            If p Mod 3 = 0 Then Exit Function                                            ' vao entre as etapas
            If (p + 2) \ 3 > Wfp_StepCount(cht) Then Exit Function
            If p Mod 3 = 1 Then ref = "1:" & ((p + 2) \ 3) Else ref = "2:" & ((p + 2) \ 3)
        Case Else
            k = Delta_DataOrdinal(cht, s)
            If k = 0 Then Exit Function
            If p > Geom_ChartCategoryCount(cht) Then Exit Function
            ref = k & ":" & p
    End Select
    Delta_ResolvePick = True
End Function

' Valor (texto) e extensao lo..hi (no eixo de valores) da referencia. False se ela nao existe mais.
Public Function Delta_Extent(cht As Chart, ref As String, ByRef v As Double, ByRef lo As Double, ByRef hi As Double) As Boolean
    Dim k As Long, c As Long, kind As String, bb As Range, s As Long, i As Long, x As Double, sumPos As Double, sumNeg As Double
    If Not Delta_ParseRef(ref, k, c) Then Exit Function
    kind = Registry_ChartType(cht)
    Select Case Kinds_Family(kind)
        Case "waterfall"
            If k <> 0 Then Exit Function
            Set bb = Block_Body(Registry_Block(cht))
            If c > bb.Rows.Count - 1 Then Exit Function
            v = Util_ToDouble(bb.Cells(c + 1, WF_COL_END).Value)
            lo = Util_ToDouble(bb.Cells(c + 1, WF_COL_LO).Value)
            hi = Util_ToDouble(bb.Cells(c + 1, WF_COL_HI).Value)
        Case "waterfallpair"
            If k < 1 Or k > 2 Or c > Wfp_StepCount(cht) Then Exit Function
            Set bb = Block_Body(Registry_Block(cht))
            v = Util_ToDouble(bb.Cells(c + 1, Wfp_SideCol(k, "Val")).Value)
            lo = Util_ToDouble(bb.Cells(c + 1, Wfp_SideCol(k, "Lo")).Value)
            hi = Util_ToDouble(bb.Cells(c + 1, Wfp_SideCol(k, "Hi")).Value)
        Case Else
            If k < 1 Then Exit Function
            s = Delta_SeriesIndex(cht, k)
            If s = 0 Or c > Geom_ChartCategoryCount(cht) Then Exit Function
            v = Delta_PointValue(cht.SeriesCollection(s), c)
            If Not Delta_IsColumnSeries(cht.SeriesCollection(s)) Then
                lo = v: hi = v                                            ' linha: o proprio ponto
            ElseIf Delta_IsStackedKind(kind) Then
                For i = 1 To s
                    If Not Style_IsHelperSeries(cht.SeriesCollection(i)) And Delta_IsColumnSeries(cht.SeriesCollection(i)) Then
                        x = Delta_PointValue(cht.SeriesCollection(i), c)
                        If x > 0 Then sumPos = sumPos + x Else sumNeg = sumNeg + x
                    End If
                Next
                If v >= 0 Then
                    lo = sumPos - v: hi = sumPos
                Else
                    lo = sumNeg: hi = sumNeg - v
                End If
            Else
                If v >= 0 Then lo = 0: hi = v Else lo = v: hi = 0
            End If
    End Select
    Delta_Extent = True
End Function

' True quando o segmento da referencia e o de cima de uma pilha positiva (o rotulo do total, por fora, fica logo depois dele): empilhadas e
' colunas da combinacao. False para segmento negativo, de baixo, linha ou referencia inexistente.
Public Function Delta_IsStackTop(cht As Chart, ref As String) As Boolean
    Dim k As Long, c As Long, v As Double, lo As Double, hi As Double, i As Long, x As Double, sumPos As Double
    If Not Delta_ParseRef(ref, k, c) Then Exit Function
    If Not Delta_Extent(cht, ref, v, lo, hi) Then Exit Function
    If v <= 0 Then Exit Function
    If Not Delta_IsColumnSeries(cht.SeriesCollection(Delta_SeriesIndex(cht, k))) Then Exit Function
    For i = 1 To cht.SeriesCollection.Count
        If Not Style_IsHelperSeries(cht.SeriesCollection(i)) And Delta_IsColumnSeries(cht.SeriesCollection(i)) Then
            x = Delta_PointValue(cht.SeriesCollection(i), c)
            If x > 0 Then sumPos = sumPos + x
        End If
    Next
    Delta_IsStackTop = (hi >= sumPos - 0.0000001)
End Function

' Indice no eixo de categorias (na comparativa, a fatia da barra; nos demais, a categoria).
Public Function Delta_AxisIndex(cht As Chart, ref As String) As Long
    Dim k As Long, c As Long
    If Not Delta_ParseRef(ref, k, c) Then Exit Function
    If Kinds_Family(Registry_ChartType(cht)) = "waterfallpair" Then Delta_AxisIndex = Wfp_SlotOf(c, k) Else Delta_AxisIndex = c
End Function

' Borda mais externa (maior hi acima; menor lo abaixo; nunca aquem do zero) de TODAS as barras do trecho ia..ib do eixo de categorias.
Public Function Delta_SpanEdge(cht As Chart, ia As Long, ib As Long, above As Boolean) As Double
    Dim a As Long, k As Long, fam As String, nData As Long, edge As Double
    fam = Kinds_Family(Registry_ChartType(cht))
    nData = Delta_DataSeriesCount(cht)
    For a = Application.WorksheetFunction.Min(ia, ib) To Application.WorksheetFunction.Max(ia, ib)
        If fam = "waterfall" Then
            Delta_EdgeAdd cht, "0:" & a, above, edge
        ElseIf fam = "waterfallpair" Then
            If a Mod 3 = 1 Then Delta_EdgeAdd cht, "1:" & ((a + 2) \ 3), above, edge
            If a Mod 3 = 2 Then Delta_EdgeAdd cht, "2:" & ((a + 2) \ 3), above, edge
        Else
            For k = 1 To nData
                Delta_EdgeAdd cht, k & ":" & a, above, edge
            Next
        End If
    Next
    Delta_SpanEdge = edge
End Function

Private Sub Delta_EdgeAdd(cht As Chart, ref As String, above As Boolean, ByRef edge As Double)
    Dim v As Double, lo As Double, hi As Double
    If Not Delta_Extent(cht, ref, v, lo, hi) Then Exit Sub
    If above Then
        If hi > edge Then edge = hi
    Else
        If lo < edge Then edge = lo
    End If
End Sub

' Grupo de eixo da serie da referencia (xlPrimary ou xlSecondary): so a linha da combinacao com o eixo secundario fica no secundario; as
' cascatas (barra do bloco) e referencias inexistentes sao primarias.
Public Function Delta_AxisGroup(cht As Chart, ref As String) As Long
    Dim k As Long, c As Long, s As Long
    Delta_AxisGroup = xlPrimary
    If Not Delta_ParseRef(ref, k, c) Then Exit Function
    Select Case Kinds_Family(Registry_ChartType(cht))
        Case "waterfall", "waterfallpair": Exit Function
    End Select
    s = Delta_SeriesIndex(cht, k)
    If s = 0 Then Exit Function
    If cht.SeriesCollection(s).AxisGroup = xlSecondary Then Delta_AxisGroup = xlSecondary
End Function

' Coordenada (pt, sentido dos valores) do valor v da serie da referencia, pela regra de tres do eixo DELA (a linha da combinacao no eixo
' secundario tem outra escala), presa a area interna (valor alem da escala fixa ou do corte de eixo: a barra termina na borda).
Public Function Delta_ValCoord(cht As Chart, ref As String, v As Double) As Double
    Dim ax As Axis, p As Double, lo As Double, hi As Double
    If Delta_AxisGroup(cht, ref) = xlSecondary Then
        Set ax = cht.Axes(xlValue, xlSecondary)
        With cht.PlotArea
            If Geom_ChartIsHorizontal(cht) Then
                p = Geom_ValueToX(.InsideLeft, .InsideWidth, ax.MinimumScale, ax.MaximumScale, v)
            Else
                p = Geom_ValueToY(.InsideTop, .InsideHeight, ax.MinimumScale, ax.MaximumScale, v)
            End If
        End With
    Else
        p = Geom_ChartValCoord(cht, v)
    End If
    With cht.PlotArea
        If Geom_ChartIsHorizontal(cht) Then lo = .InsideLeft: hi = .InsideLeft + .InsideWidth Else lo = .InsideTop: hi = .InsideTop + .InsideHeight
    End With
    If p < lo Then p = lo
    If p > hi Then p = hi
    Delta_ValCoord = p
End Function

' Borda mais externa (pt) de TODAS as barras do trecho ia..ib, no lado da seta: dirn = sentido "para fora" em pt (-1 = para cima nas colunas,
' +1 = para a direita nas barras; invertido abaixo). Cada extensao vai pela regra de tres do eixo da sua serie (Delta_ValCoord); comeca no
' zero do eixo primario (a borda nunca fica aquem dele).
Public Function Delta_SpanEdgeCoord(cht As Chart, ia As Long, ib As Long, above As Boolean, dirn As Double) As Double
    Dim a As Long, k As Long, fam As String, nData As Long, edge As Double
    fam = Kinds_Family(Registry_ChartType(cht))
    nData = Delta_DataSeriesCount(cht)
    edge = Delta_ValCoord(cht, "0:1", 0)                ' "0:1" e sempre primaria (k = 0 nao e serie de dados)
    For a = Application.WorksheetFunction.Min(ia, ib) To Application.WorksheetFunction.Max(ia, ib)
        If fam = "waterfall" Then
            Delta_EdgeAddCoord cht, "0:" & a, above, dirn, edge
        ElseIf fam = "waterfallpair" Then
            If a Mod 3 = 1 Then Delta_EdgeAddCoord cht, "1:" & ((a + 2) \ 3), above, dirn, edge
            If a Mod 3 = 2 Then Delta_EdgeAddCoord cht, "2:" & ((a + 2) \ 3), above, dirn, edge
        Else
            For k = 1 To nData
                Delta_EdgeAddCoord cht, k & ":" & a, above, dirn, edge
            Next
        End If
    Next
    Delta_SpanEdgeCoord = edge
End Function

Private Sub Delta_EdgeAddCoord(cht As Chart, ref As String, above As Boolean, dirn As Double, ByRef edge As Double)
    Dim v As Double, lo As Double, hi As Double, p As Double
    If Not Delta_Extent(cht, ref, v, lo, hi) Then Exit Sub
    If above Then p = Delta_ValCoord(cht, ref, hi) Else p = Delta_ValCoord(cht, ref, lo)
    If dirn * p > dirn * edge Then edge = p
End Sub

' Cascata comparativa com as setas de variacao do par ligadas (auxiliar tqq_pdroom: um ponto por etapa, 12 pt acima da linha horizontal da
' seta, a sala da elipse): a coordenada (pt) do ponto de sala mais alto entre as etapas que tem fatia no trecho ia..ib. False sem setas, abaixo,
' fora da comparativa ou sem etapa no trecho.
Public Function Delta_PairArrowEdgeCoord(cht As Chart, ia As Long, ib As Long, above As Boolean, ByRef edge As Double) As Boolean
    Dim s As Long, vals As Variant, i As Long, lo As Long, hi As Long, p As Double
    If Not above Then Exit Function
    If Kinds_Family(Registry_ChartType(cht)) <> "waterfallpair" Then Exit Function
    s = Delta_FindSeriesByName(cht, "tqq_pdroom")
    If s = 0 Then Exit Function
    vals = cht.SeriesCollection(s).Values
    If Not IsArray(vals) Then Exit Function
    lo = Application.WorksheetFunction.Min(ia, ib): hi = Application.WorksheetFunction.Max(ia, ib)
    For i = 1 To UBound(vals) - LBound(vals) + 1
        If (3 * i - 2 >= lo And 3 * i - 2 <= hi) Or (3 * i - 1 >= lo And 3 * i - 1 <= hi) Then
            If IsNumeric(vals(LBound(vals) + i - 1)) And Not IsEmpty(vals(LBound(vals) + i - 1)) Then
                p = Geom_ChartValCoord(cht, CDbl(vals(LBound(vals) + i - 1)))
                If Not Delta_PairArrowEdgeCoord Or p < edge Then edge = p
                Delta_PairArrowEdgeCoord = True
            End If
        End If
    Next
End Function

' Centro da barra no sentido das categorias (pt). Agrupadas: o centro do PONTO da serie (cada serie tem a sua barra; geometria do Excel,
' valida depois do layout); demais: o centro da categoria (Geom_ChartCatCoord; na comparativa a fatia).
Public Function Delta_CatCenter(cht As Chart, ref As String) As Double
    Dim k As Long, c As Long, kind As String, pt As Point, s As Long
    If Not Delta_ParseRef(ref, k, c) Then Exit Function
    kind = Registry_ChartType(cht)
    If kind = "clustered" Or kind = "barclustered" Then
        s = Delta_SeriesIndex(cht, k)
        If s = 0 Then Exit Function
        Set pt = cht.SeriesCollection(s).Points(c)
        If Geom_ChartIsHorizontal(cht) Then Delta_CatCenter = pt.Top + pt.Height / 2 Else Delta_CatCenter = pt.Left + pt.Width / 2
    Else
        Delta_CatCenter = Geom_ChartCatCoord(cht, Delta_AxisIndex(cht, ref))
    End If
End Function

' Meia espessura da barra no sentido das categorias (moldura da origem no modo de dois cliques).
Public Function Delta_BarHalf(cht As Chart, ref As String) As Double
    Dim k As Long, c As Long, kind As String, pt As Point, s As Long
    If Not Delta_ParseRef(ref, k, c) Then Exit Function
    kind = Registry_ChartType(cht)
    If kind = "clustered" Or kind = "barclustered" Then
        s = Delta_SeriesIndex(cht, k)
        If s = 0 Then Exit Function
        Set pt = cht.SeriesCollection(s).Points(c)
        If Geom_ChartIsHorizontal(cht) Then Delta_BarHalf = pt.Height / 2 Else Delta_BarHalf = pt.Width / 2
    ElseIf kind = "line" Then
        Delta_BarHalf = 5
    Else
        s = Delta_SeriesIndex(cht, k)
        If s > 0 Then
            If Not Delta_IsColumnSeries(cht.SeriesCollection(s)) Then Delta_BarHalf = 5: Exit Function
        End If
        Delta_BarHalf = Geom_ChartBarHalfWidth(cht)
    End If
End Function
