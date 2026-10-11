Attribute VB_Name = "modAnnotations"
Option Explicit
' Anotações = formas dentro do gráfico (movem com ele). Parâmetros no AlternativeText.
Private mRefreshing As Boolean
Private mLastError As String
Private Const ANN_ARROW_GAP As Double = 10      ' do fim da barra/topo da coluna ate a ponta da seta de CAGR
Private Const ANN_TOTALS_EXTRA As Double = 18   ' afastamento extra (minimo) da seta de CAGR com os totais ligados; nas colunas, calibrado para a ALTURA do rotulo do total
Private Const ANN_IND100 As String = "tqq_ann_ind100"     ' nome da linha do indicador 100% (o rotulo e ANN_IND100 & "_lbl")
Private Const ANN_AUTO_SLACK As Double = 18     ' folga que o eixo com maximo automatico deixa depois da maior barra (~5% da escala, medida em graficos de ~460 pt)
Private Const ANN_BRACKET As String = "tqq_ann_bracket_"  ' prefixo dos colchetes do Gantt (linha nm; tiques nm_t1/nm_t2; rotulo nm_lbl)
Private Const ANN_BRACKET_Y0 As Double = 25     ' linha do 1o colchete: 25 pt acima da area interna; os tiques de 5 pt param 20 pt acima, livres
                                                ' dos rotulos da escala (medido: de InsideTop - 18,3 a InsideTop - 5,5)
Private Const ANN_BRACKET_STEP As Double = 20   ' cada colchete a mais sobe 20 pt: os tiques do de cima (5 pt) ficam acima do rotulo do de baixo (13,1 pt)
Private Const ANN_BRACKET_ROOM As Double = 42   ' folga (pt) acima da area interna com um colchete: 25 + rotulo de 13,1 + respiro; +20 por colchete
Public Const ANN_BRACKET_REFUSAL As String = "O colchete só está disponível em gráficos Gantt."

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
    Dim kind As String, n As Long, i As Long, k As Long, s As Series, vals As Variant, v As Double, ends() As Double
    Dim nSeries As Long, lastIdx As Long, cum() As Double, algebraic As Boolean
    kind = Registry_ChartType(cht)
    n = Geom_ChartCategoryCount(cht)
    If n = 0 Then
        ReDim values(0 To 0): ReDim tops(0 To 0)
        Exit Sub
    End If
    ReDim values(1 To n): ReDim tops(1 To n)
    Select Case kind
        Case "waterfall", "barwaterfall"
            ends = Waterfall_EndValues(cht)
            For i = 1 To n
                If i <= UBound(ends) Then values(i) = ends(i): tops(i) = ends(i)
            Next
        Case "clustered", "barclustered", "line"
            For i = 1 To n
                v = Ann_SeriesValue(cht.SeriesCollection(1), i)
                values(i) = v: tops(i) = v
            Next
        Case Else
            ' soma das series (stacked, stacked100, barstacked, barstacked100, area, area100, mekko, mekkounits); na combinacao a
            ' ultima serie nao auxiliar e a linha e fica de fora (so as colunas empilham). No Mekko n e o numero de PONTOS (5 por
            ' categoria) e o topo de cada ponto e o maior acumulado, como na area (e uma area empilhada): e o que Deco_ScaleClipped le
            ' nas unidades (Mesma escala); o Mekko nao tem setas (recusa 7001)
            If kind = "combo" Then
                For k = 1 To cht.SeriesCollection.Count
                    If Not Style_IsHelperSeries(cht.SeriesCollection(k)) Then nSeries = nSeries + 1: lastIdx = k
                Next
                If nSeries < 2 Then lastIdx = 0
            End If
            ReDim cum(1 To n)
            algebraic = (kind = "area" Or Kinds_Family(kind) = "mekko")
            For k = 1 To cht.SeriesCollection.Count
                Set s = cht.SeriesCollection(k)
                If Not Style_IsHelperSeries(s) And k <> lastIdx Then
                    vals = s.Values
                    For i = 1 To n
                        v = Ann_ArrayValue(vals, i)
                        values(i) = values(i) + v
                        If algebraic Then
                            ' area: o Excel empilha algebricamente (cada faixa entre os acumulados) e o topo VISIVEL e o maior acumulado
                            ' (ver o comentario de Area_BuildBlock); colunas e barras separam positivos e negativos
                            cum(i) = cum(i) + v
                            If cum(i) > tops(i) Then tops(i) = cum(i)
                        ElseIf v > 0 Then
                            tops(i) = tops(i) + v
                        End If
                    Next
                End If
            Next
            If Kinds_Is100(kind) Then
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
    If Not Kinds_Supports(Registry_ChartType(cht), "cagr") Then Err.Raise vbObjectError + 7001, "Ann_AddCagr", "A seta CAGR não está disponível para este tipo de gráfico."
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
    If Geom_ChartIsHorizontal(cht) Then Ann_Refresh cht     ' escreve o texto: a folga a direita usa a largura real do rotulo
    Ann_EnsureRoom cht
    Ann_Refresh cht
    Set Ann_AddCagr = ln
End Function

' Abre espaço na área de plotagem para as setas e seus rótulos (só aumenta a folga).
' Colunas: com os totais ligados, 52 pt acima da área de plotagem para a seta de CAGR e seu rótulo passarem acima
' dos rótulos dos totais (mesmo quando a coluna mais alta quase toca o topo do eixo).
' Barras: havendo seta de CAGR, a folga a direita da área interna (Ann_RightRoomNeeded) e, havendo seta de diferença,
' 34 pt abaixo dela (10 até a seta + 6 + rótulo de ~16 + 2). A folga conta até a borda livre: a da área do gráfico ou,
' com legenda à direita/embaixo, a borda da legenda (o gráfico não desenha setas por cima dela).
Private Sub Ann_EnsureRoom(cht As Chart)
    Dim need As Double, delta As Double, k As Long, nb As Long
    On Error Resume Next
    ' colchetes (Gantt): acima da area interna, sobre a escala de datas; o k-esimo colchete fica 20 pt acima do anterior
    nb = Ann_BracketMaxOrdinal(cht)
    If nb > 0 Then
        need = ANN_BRACKET_ROOM + ANN_BRACKET_STEP * (nb - 1)
        For k = 1 To 3
            delta = need - cht.PlotArea.InsideTop
            If delta <= 0.01 Then Exit For
            cht.PlotArea.Height = cht.PlotArea.Height - delta
            cht.PlotArea.Top = cht.PlotArea.Top + delta
        Next
    End If
    If Geom_ChartIsHorizontal(cht) Then
        If Ann_CountPrefix(cht, ANN_PREFIX & "cagr_") > 0 Then
            need = Ann_RightRoomNeeded(cht)
            For k = 1 To 3     ' a largura interna nem sempre acompanha 1:1 a externa: confere e repete
                delta = need - (Geom_ChartRightEdge(cht) - (cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth))
                If delta <= 0.01 Then Exit For
                cht.PlotArea.Width = cht.PlotArea.Width - delta
            Next
        End If
        If Ann_CountPrefix(cht, ANN_PREFIX & "diff_") > 0 Then
            For k = 1 To 3
                delta = 34 - (Ann_BottomEdge(cht) - (cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight))
                If delta <= 0.01 Then Exit For
                cht.PlotArea.Height = cht.PlotArea.Height - delta
            Next
        End If
    Else
        If Registry_GetOption(cht, "totals", "0") <> "1" Then Exit Sub
        If Ann_CountPrefix(cht, ANN_PREFIX & "cagr_") = 0 Then Exit Sub        ' só a seta de CAGR precisa da folga acima
        need = 52
        delta = need - cht.PlotArea.InsideTop
        If delta > 0 Then
            cht.PlotArea.Height = cht.PlotArea.Height - delta
            cht.PlotArea.Top = cht.PlotArea.Top + delta
        End If
    End If
    If Err.Number <> 0 Then Errors_Log "Ann_EnsureRoom" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Reexecuta a folga das setas depois de uma mudança posterior à criação (legenda, formato numérico dos rótulos, tamanho do
' gráfico): Ann_EnsureRoom só roda ao inserir a seta, e a legenda ligada depois, um rótulo de total mais largo ou um gráfico
' redimensionado podem deixar a seta e o rótulo sob a legenda ou fora da folga. Só reduz a área de plotagem. Sem setas não faz nada.
Public Sub Ann_EnsureRoomAfterChange(cht As Chart)
    If Ann_CountPrefix(cht, ANN_PREFIX & "cagr_") + Ann_CountPrefix(cht, ANN_PREFIX & "diff_") + Ann_CountPrefix(cht, ANN_BRACKET) = 0 Then Exit Sub
    Ann_EnsureRoom cht
End Sub

' Afastamento extra da seta de CAGR quando os totais estão ligados (definição única: layout e folga usam a mesma).
' Colunas: 18 pt fixos (calibrados para a ALTURA do rótulo do total, que fica acima da coluna). Barras: o rótulo do
' total fica ~6 pt depois do fim da barra e ocupa a sua LARGURA (4 dígitos: ~27 pt; 7 dígitos: ~43 pt), e a ponta da
' seta fica a 10 + extra do fim da barra; para sobrar 4 pt de respiro depois do rótulo: 6 + largura + 4 = 10 + extra,
' ou seja, extra = max(18, largura estimada do maior rótulo de total: Labels_MaxTotalWidth). Com 2 e 3 dígitos vale o
' mínimo de 18.
Private Function Ann_TotalsExtra(cht As Chart) As Double
    Dim w As Double, ex As Double
    If Registry_GetOption(cht, "totals", "0") <> "1" Then Exit Function
    ex = ANN_TOTALS_EXTRA
    If Geom_ChartIsHorizontal(cht) Then
        w = Labels_MaxTotalWidth(cht)
        If w > ex Then ex = w
    End If
    Ann_TotalsExtra = ex
End Function

' Folga (pt) entre a área interna e a borda livre para a seta de CAGR em barras e o rótulo: ponta a 10 pt (+ extra dos
' totais) do fim da barra, 4 pt de respiro, o rótulo (largura medida; 62 se ainda não desenhado) e 2 pt de margem.
' Com o máximo do eixo automático o Excel já deixa ~5% da escala (ANN_AUTO_SLACK) depois da maior barra: desconta isso (mínimo 60).
' Em barras 100% ou com o máximo fixo a maior barra pode terminar no fim do eixo: folga cheia (mínimo 80).
Private Function Ann_RightRoomNeeded(cht As Chart) As Double
    Dim w As Double, base As Double, need As Double
    w = Ann_MaxCagrLabelWidth(cht)
    If w < 40 Then w = 62
    base = ANN_ARROW_GAP + 4 + w + 2
    If Ann_BarMayReachAxisEnd(cht) Then
        need = 80
        If base > need Then need = base
    Else
        need = 60
        If base - ANN_AUTO_SLACK > need Then need = base - ANN_AUTO_SLACK
    End If
    Ann_RightRoomNeeded = need + Ann_TotalsExtra(cht)
End Function

' True quando a barra mais longa pode terminar no fim do eixo: barras 100% ou máximo do eixo de valores fixo
' (conservador se a leitura do eixo falhar).
Private Function Ann_BarMayReachAxisEnd(cht As Chart) As Boolean
    Dim isAuto As Boolean
    If Kinds_Is100(Registry_ChartType(cht)) Then Ann_BarMayReachAxisEnd = True: Exit Function
    On Error Resume Next
    isAuto = cht.Axes(xlValue, xlPrimary).MaximumScaleIsAuto
    On Error GoTo 0
    Ann_BarMayReachAxisEnd = Not isAuto
End Function

' Maior largura entre os rótulos das setas de CAGR existentes (0 se não houver nenhum).
Private Function Ann_MaxCagrLabelWidth(cht As Chart) As Double
    Dim i As Long, nm As String, p As String, best As Double
    p = ANN_PREFIX & "cagr_"
    For i = 1 To cht.Shapes.Count
        nm = cht.Shapes(i).Name
        If Left$(nm, Len(p)) = p And Right$(nm, 4) = "_lbl" Then
            If cht.Shapes(i).Width > best Then best = cht.Shapes(i).Width
        End If
    Next
    Ann_MaxCagrLabelWidth = best
End Function

' Borda inferior livre: a da área do gráfico ou, com legenda embaixo, o topo dela.
Private Function Ann_BottomEdge(cht As Chart) As Double
    Dim edge As Double, tp As Double
    edge = cht.ChartArea.Height
    On Error Resume Next
    If cht.HasLegend Then
        If cht.Legend.Position = xlLegendPositionBottom Then
            tp = cht.Legend.Top
            If tp > edge / 2 And tp < edge Then edge = tp
        End If
    End If
    On Error GoTo 0
    Ann_BottomEdge = edge
End Function

Public Function Ann_AddDifference(cht As Chart, fromIdx As Long, toIdx As Long) As Shape
    Dim nm As String, ln As Shape, lvl As Shape, kind As String
    kind = Registry_ChartType(cht)
    If Not Kinds_Supports(kind, "diff") Then
        If Kinds_Is100(kind) And Kinds_Family(kind) <> "mekko" Then     ' o Mekko recebe a mensagem genérica
            Err.Raise vbObjectError + 7001, "Ann_AddDifference", "A seta de diferença não está disponível para gráficos 100%."
        Else
            Err.Raise vbObjectError + 7001, "Ann_AddDifference", "A seta de diferença não está disponível para este tipo de gráfico."
        End If
    End If
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
    If Geom_ChartIsHorizontal(cht) Then Ann_EnsureRoom cht     ' colunas: como na Fase 1, a diferenca nao mexe na area de plotagem
    Ann_Refresh cht
    Set Ann_AddDifference = ln
End Function

' Colchete do Gantt sobre as atividades fromIdx..toIdx (fromIdx = toIdx vale: uma atividade so): linha horizontal nm (kind=bracket;from=i;to=j)
' na area da escala, acima da area de plotagem, com dois tiques de 5 pt para baixo nas pontas (nm_t1, nm_t2: kind=tick;owner=nm) e o rotulo
' nm_lbl com o periodo (dd/mm - dd/mm) acima da linha. Ann_Refresh posiciona (Ann_LayoutBracket); Ann_EnsureRoom abre a folga de cima (42 pt +
' 20 pt por colchete a mais). Recusa (7001) fora do Gantt.
Public Function Ann_AddBracket(cht As Chart, fromIdx As Long, toIdx As Long) As Shape
    Dim nm As String, ln As Shape, tk As Shape, lbl As Shape, sfx As Variant
    If Not Kinds_Supports(Registry_ChartType(cht), "bracket") Then Err.Raise vbObjectError + 7001, "Ann_AddBracket", ANN_BRACKET_REFUSAL
    nm = Ann_NewName(cht, "bracket")
    Set ln = cht.Shapes.AddLine(0, 0, 10, 0)
    ln.Name = nm
    ln.AlternativeText = "kind=bracket;from=" & fromIdx & ";to=" & toIdx
    ln.Line.ForeColor.RGB = RGB(89, 89, 89): ln.Line.Weight = 1
    For Each sfx In Array("_t1", "_t2")
        Set tk = cht.Shapes.AddLine(0, 0, 0, 5)
        tk.Name = nm & sfx
        tk.AlternativeText = "kind=tick;owner=" & nm
        tk.Line.ForeColor.RGB = RGB(89, 89, 89): tk.Line.Weight = 1
    Next
    Ann_AddLabel cht, nm
    Set lbl = cht.Shapes(nm & "_lbl")
    lbl.TextFrame2.TextRange.Font.Bold = msoFalse
    lbl.Fill.Visible = msoFalse
    lbl.Line.Visible = msoFalse
    Ann_EnsureRoom cht
    Ann_Refresh cht
    Set Ann_AddBracket = ln
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

' Remove as setas (e, com includeConnectors, os conectores da cascata). O indicador 100% (tqq_ann_ind100 e o rotulo) NUNCA e
' removido aqui: e um alternador proprio (Ann_AddInd100 / Ann_RemoveInd100, ligado por Deco_SetInd100).
Public Sub Ann_RemoveAll(cht As Chart, Optional includeConnectors As Boolean = False)
    Dim i As Long, nm As String
    For i = cht.Shapes.Count To 1 Step -1
        nm = cht.Shapes(i).Name
        If Left$(nm, Len(ANN_PREFIX)) = ANN_PREFIX And Left$(nm, Len(ANN_IND100)) <> ANN_IND100 Then
            If includeConnectors Or Left$(nm, Len(ANN_PREFIX & "conn_")) <> ANN_PREFIX & "conn_" Then cht.Shapes(i).Delete
        End If
    Next
End Sub

' Indicador 100%: um tique no nivel 100% do eixo de valores (tqq_ann_ind100, 1 pt, kind=ind100) e o rotulo "100%" (tqq_ann_ind100_lbl,
' kind=label;owner=tqq_ann_ind100) fora da area de plotagem: a esquerda nas colunas e areas, em cima nas barras. Ann_Refresh posiciona
' (Ann_LayoutInd100). Fica aqui porque Ann_NewName/Ann_AddLabel sao privadas deste modulo. Se as formas ja existem nao cria de novo.
' Recusa (7002) fora dos tipos com o recurso (a matriz de tipos: Kinds_Supports "ind100"). Quem chama (Deco_SetInd100) abre a folga
' e chama Ann_Refresh.
Public Function Ann_AddInd100(cht As Chart) As Shape
    Dim ln As Shape, lbl As Shape
    If Not Kinds_Supports(Registry_ChartType(cht), "ind100") Then Err.Raise vbObjectError + 7002, "Ann_AddInd100", "O indicador 100% só existe em gráficos 100% e Mekko (%)."
    Set ln = Ann_Shape(cht, ANN_IND100)
    If ln Is Nothing Then
        Set ln = cht.Shapes.AddLine(0, 0, 10, 0)
        ln.Name = ANN_IND100
        ln.AlternativeText = "kind=ind100"
        ln.Line.ForeColor.RGB = RGB(89, 89, 89)
        ln.Line.Weight = 1
    End If
    If Ann_Shape(cht, ANN_IND100 & "_lbl") Is Nothing Then
        Ann_AddLabel cht, ANN_IND100
        Set lbl = cht.Shapes(ANN_IND100 & "_lbl")
        lbl.TextFrame2.TextRange.Text = "100%"
        lbl.TextFrame2.TextRange.Font.Bold = msoFalse
        lbl.Fill.Visible = msoFalse
        lbl.Line.Visible = msoFalse
    End If
    Set Ann_AddInd100 = ln
End Function

' True se alguma das formas do indicador 100% existe (Deco_SetInd100 desligando: so ai ha folga a devolver).
Public Function Ann_HasInd100(cht As Chart) As Boolean
    Ann_HasInd100 = Not Ann_Shape(cht, ANN_IND100) Is Nothing Or Not Ann_Shape(cht, ANN_IND100 & "_lbl") Is Nothing
End Function

' Apaga as duas formas do indicador 100% (nada a fazer se nao existem).
Public Sub Ann_RemoveInd100(cht As Chart)
    Dim sfx As Variant, s As Shape
    For Each sfx In Array("_lbl", "")
        Set s = Ann_Shape(cht, ANN_IND100 & sfx)
        If Not s Is Nothing Then s.Delete
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
    If Kinds_Family(chartKind) = "waterfall" And Registry_GetOption(cht, "connectors", "0") = "1" Then
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
                If valid And (kind = "cagr" Or kind = "diff") Then valid = Kinds_Supports(chartKind, kind)
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
            ElseIf kind = "bracket" Then
                ' colchete (Gantt): uma atividade so (de = ate) vale; sem data nenhuma nas atividades do intervalo, oculto
                fromIdx = Val(Ann_ParseParam(shp.AlternativeText, "from"))
                toIdx = Val(Ann_ParseParam(shp.AlternativeText, "to"))
                valid = (fromIdx >= 1 And toIdx <= n And fromIdx <= toIdx)
                If valid Then valid = Kinds_Supports(chartKind, "bracket")
                If valid Then
                    valid = False                    ' falha no meio do layout: oculto, nunca na posicao antiga
                    On Error Resume Next
                    valid = Ann_LayoutBracket(cht, shp, fromIdx, toIdx)
                    If Err.Number <> 0 Then
                        valid = False
                        mLastError = "Ann_Refresh " & nm & " " & Err.Number & " " & Err.Description
                        Errors_Log mLastError
                        Err.Clear
                    End If
                    On Error GoTo Done
                End If
                Ann_SetShapeVisible cht, nm, valid
            ElseIf kind = "ind100" Then
                ' indicador 100%: valido so nos tipos com o recurso (senao oculto); a posicao segue o PlotArea e a escala atuais
                valid = Kinds_Supports(chartKind, "ind100")
                Ann_SetShapeVisible cht, nm, valid
                If valid Then
                    On Error Resume Next
                    Ann_LayoutInd100 cht, shp
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
    Loop While Not Ann_Shape(cht, nm) Is Nothing Or Not Ann_Shape(cht, nm & "_lbl") Is Nothing Or Not Ann_Shape(cht, nm & "_lvl") Is Nothing Or Not Ann_Shape(cht, nm & "_t1") Is Nothing
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
    Dim body As Range, periods As Double, extra As Double, xMax As Double
    ' Com os totais ligados, o rótulo do total ocupa o espaço logo depois da barra: a seta se afasta um pouco mais.
    extra = Ann_TotalsExtra(cht)
    If Geom_ChartIsHorizontal(cht) Then
        ' barras: a seta corre ao longo das categorias, à direita do fim das barras (ponta em 'to'); o limite
        ' (24 pt da borda livre) só atua quando Ann_EnsureRoom não conseguiu criar a folga
        y1 = Geom_ChartCatCoord(cht, fromIdx): y2 = Geom_ChartCatCoord(cht, toIdx)
        x1 = Geom_ChartValCoord(cht, tops(fromIdx)) + ANN_ARROW_GAP + extra
        x2 = Geom_ChartValCoord(cht, tops(toIdx)) + ANN_ARROW_GAP + extra
        xMax = Geom_ChartRightEdge(cht) - 24
        If x1 > xMax Then x1 = xMax
        If x2 > xMax Then x2 = xMax
    Else
        x1 = Geom_ChartCatCoord(cht, fromIdx): x2 = Geom_ChartCatCoord(cht, toIdx)
        y1 = Geom_ChartValCoord(cht, tops(fromIdx)) - ANN_ARROW_GAP - extra
        y2 = Geom_ChartValCoord(cht, tops(toIdx)) - ANN_ARROW_GAP - extra
        ' fica dentro do gráfico, com espaço para o rótulo da seta acima da linha
        If y1 < 22 Then y1 = 22
        If y2 < 22 Then y2 = 22
    End If
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
    If Geom_ChartIsHorizontal(cht) Then
        lbl.Left = IIf(x1 > x2, x1, x2) + 4
        If lbl.Left + lbl.Width > Geom_ChartRightEdge(cht) Then lbl.Left = Geom_ChartRightEdge(cht) - lbl.Width     ' nunca sobre a legenda
        lbl.Top = (y1 + y2) / 2 - lbl.Height / 2
    Else
        lbl.Left = (x1 + x2) / 2 - lbl.Width / 2
        lbl.Top = IIf(y1 < y2, y1, y2) - lbl.Height - 4
    End If
    Ann_ClampShape cht, lbl
End Sub

Private Sub Ann_LayoutDiff(cht As Chart, ln As Shape, fromIdx As Long, toIdx As Long, values() As Double, tops() As Double, fmt As String)
    Dim xFrom As Double, xTo As Double, yFrom As Double, yTo As Double, xArrow As Double, yArrow As Double, half As Double
    Dim lvl As Shape, lbl As Shape, diff As Double, txt As String
    half = Geom_ChartBarHalfWidth(cht)
    Set lvl = Ann_Shape(cht, ln.Name & "_lvl")
    If Geom_ChartIsHorizontal(cht) Then
        ' barras: seta dupla horizontal abaixo da última barra; linha de nível vertical a partir do fim da barra inicial
        xFrom = Geom_ChartValCoord(cht, tops(fromIdx)): xTo = Geom_ChartValCoord(cht, tops(toIdx))
        yArrow = Geom_ChartCatCoord(cht, toIdx) + half + 10
        If yArrow > Ann_BottomEdge(cht) - 2 Then yArrow = Ann_BottomEdge(cht) - 2
        If Not lvl Is Nothing Then Ann_MoveLine lvl, xFrom, Geom_ChartCatCoord(cht, fromIdx) + half, xFrom, yArrow + 4
        Ann_MoveLine ln, xFrom, yArrow, xTo, yArrow
    Else
        xFrom = Geom_ChartCatCoord(cht, fromIdx): xTo = Geom_ChartCatCoord(cht, toIdx)
        yFrom = Geom_ChartValCoord(cht, tops(fromIdx)): yTo = Geom_ChartValCoord(cht, tops(toIdx))
        xArrow = xTo + half + 10
        If xArrow > cht.ChartArea.Width - 2 Then xArrow = cht.ChartArea.Width - 2
        If Not lvl Is Nothing Then Ann_MoveLine lvl, xFrom + half, yFrom, xArrow + 4, yFrom
        Ann_MoveLine ln, xArrow, yFrom, xArrow, yTo
    End If
    diff = values(toIdx) - values(fromIdx)
    txt = Format$(diff, Ann_SignedFormat(fmt))
    If values(fromIdx) <> 0 Then txt = txt & " (" & Format$(diff / Abs(values(fromIdx)), "+0%;-0%;0%") & ")"
    Set lbl = Ann_Shape(cht, ln.Name & "_lbl")
    If lbl Is Nothing Then Exit Sub
    lbl.TextFrame2.TextRange.Text = txt
    If Geom_ChartIsHorizontal(cht) Then
        lbl.Left = (xFrom + xTo) / 2 - lbl.Width / 2
        lbl.Top = yArrow + 6
    Else
        lbl.Left = xArrow + 6
        lbl.Top = (yFrom + yTo) / 2 - lbl.Height / 2
    End If
    Ann_ClampShape cht, lbl
End Sub

Private Sub Ann_LayoutConn(cht As Chart, shp As Shape, i As Long, j As Long, tops() As Double)
    Dim half As Double, x As Double, y As Double
    half = Geom_ChartBarHalfWidth(cht)
    If Geom_ChartIsHorizontal(cht) Then
        ' cascata horizontal: linha vertical em x = fim da barra i, da base da barra i ao topo da barra j
        x = Geom_ChartValCoord(cht, tops(i))
        Ann_MoveLine shp, x, Geom_ChartCatCoord(cht, i) + half, x, Geom_ChartCatCoord(cht, j) - half
    Else
        y = Geom_ChartValCoord(cht, tops(i))
        Ann_MoveLine shp, Geom_ChartCatCoord(cht, i) + half, y, Geom_ChartCatCoord(cht, j) - half, y
    End If
End Sub

' Indicador 100%: o tique no nivel 100% do eixo de valores. Colunas e areas: tique de (InsideLeft - 8, y) a (InsideLeft - 1, y), rotulo
' a esquerda dele; barras: x = fim da barra 100%, tique logo acima da area de plotagem (InsideTop - 8 a - 1), rotulo centrado acima.
' Deco_EnsureIndicatorRoom abre a folga para o rotulo caber; o Ann_ClampShape so atua se nao coube.
Private Sub Ann_LayoutInd100(cht As Chart, ln As Shape)
    Dim lbl As Shape, v As Double
    v = Geom_ChartValCoord(cht, 1)
    Set lbl = Ann_Shape(cht, ln.Name & "_lbl")
    With cht.PlotArea
        If Geom_ChartIsHorizontal(cht) Then
            Ann_MoveLine ln, v, .InsideTop - 8, v, .InsideTop - 1
            If Not lbl Is Nothing Then lbl.Left = v - lbl.Width / 2: lbl.Top = .InsideTop - 10 - lbl.Height
        Else
            Ann_MoveLine ln, .InsideLeft - 8, v, .InsideLeft - 1, v
            If Not lbl Is Nothing Then lbl.Left = .InsideLeft - 10 - lbl.Width: lbl.Top = v - lbl.Height / 2
        End If
    End With
    If Not lbl Is Nothing Then Ann_ClampShape cht, lbl
End Sub

' Colchete do Gantt: para cada atividade k em fromIdx..toIdx, ultimo dia le = fim (ou o inicio, num marco) e fim plotado pe = fim + 1 (a barra cobre
' o dia do fim) ou o proprio dia do marco; d1 = menor inicio, d2 = maior le (zeros, de celulas sem data, ignorados). A linha vai de x(d1) a x(maior pe)
' na altura InsideTop - 25 - 20 * (ordinal - 1) (ordinal = sufixo numerico do nome: o segundo colchete fica acima do primeiro), os tiques descem 5 pt
' nas pontas e o rotulo "dd/mm - dd/mm" (d1 - d2; so "dd/mm" quando d1 = d2, um marco) fica centrado acima da linha. As datas sao numeros de serie da
' pasta (sistema 1900 ou 1904): o texto passa por Gantt_SerialToDate. False (nada movido) sem nenhuma data no intervalo.
Private Function Ann_LayoutBracket(cht As Chart, ln As Shape, fromIdx As Long, toIdx As Long) As Boolean
    Dim starts() As Double, ends() As Double, k As Long, d1 As Double, d2 As Double, le As Double, pe As Double, peMax As Double
    Dim x1 As Double, x2 As Double, y As Double, t As Shape, lbl As Shape
    Gantt_Dates cht, starts, ends
    For k = fromIdx To toIdx
        If k >= 1 And k >= LBound(starts) And k <= UBound(starts) Then
            If starts(k) > 0 Then
                If d1 = 0 Or starts(k) < d1 Then d1 = starts(k)
            End If
            If ends(k) > 0 Then
                le = ends(k): pe = ends(k) + 1
            Else
                le = starts(k): pe = starts(k)
            End If
            If le > d2 Then d2 = le
            If pe > peMax Then peMax = pe
        End If
    Next
    If d1 <= 0 Or d2 <= 0 Or peMax <= 0 Then Exit Function
    x1 = Geom_ChartValueToX(cht, d1)
    x2 = Geom_ChartValueToX(cht, peMax)
    y = cht.PlotArea.InsideTop - ANN_BRACKET_Y0 - ANN_BRACKET_STEP * (Ann_BracketOrdinal(ln.Name) - 1)
    Ann_MoveLine ln, x1, y, x2, y
    Set t = Ann_Shape(cht, ln.Name & "_t1")
    If Not t Is Nothing Then Ann_MoveLine t, x1, y, x1, y + 5
    Set t = Ann_Shape(cht, ln.Name & "_t2")
    If Not t Is Nothing Then Ann_MoveLine t, x2, y, x2, y + 5
    Set lbl = Ann_Shape(cht, ln.Name & "_lbl")
    If Not lbl Is Nothing Then
        If d2 = d1 Then         ' um dia so (um marco): a data uma vez
            lbl.TextFrame2.TextRange.Text = Format$(Gantt_SerialToDate(cht, d1), "dd/mm")
        Else
            lbl.TextFrame2.TextRange.Text = Format$(Gantt_SerialToDate(cht, d1), "dd/mm") & " - " & Format$(Gantt_SerialToDate(cht, d2), "dd/mm")
        End If
        lbl.Left = (x1 + x2) / 2 - lbl.Width / 2
        lbl.Top = y - lbl.Height
        Ann_ClampShape cht, lbl
    End If
    Ann_LayoutBracket = True
End Function

' Ordinal do colchete: o sufixo numerico do nome (tqq_ann_bracket_2 -> 2); 1 se nao houver.
Private Function Ann_BracketOrdinal(nm As String) As Long
    Ann_BracketOrdinal = CLng(Val(Mid$(nm, Len(ANN_BRACKET) + 1)))
    If Ann_BracketOrdinal < 1 Then Ann_BracketOrdinal = 1
End Function

' Maior ordinal entre os colchetes do grafico (0 sem colchete): a folga de cima cobre o colchete mais alto.
Private Function Ann_BracketMaxOrdinal(cht As Chart) As Long
    Dim i As Long, nm As String, k As Long
    For i = 1 To cht.Shapes.Count
        nm = cht.Shapes(i).Name
        If Left$(nm, Len(ANN_BRACKET)) = ANN_BRACKET Then
            If Ann_ParseParam(cht.Shapes(i).AlternativeText, "kind") = "bracket" Then
                k = Ann_BracketOrdinal(nm)
                If k > Ann_BracketMaxOrdinal Then Ann_BracketMaxOrdinal = k
            End If
        End If
    Next
End Function

Private Sub Ann_LayoutOne(cht As Chart, shp As Shape, kind As String, fromIdx As Long, toIdx As Long, values() As Double, tops() As Double, fmt As String)
    Select Case kind
        Case "cagr": Ann_LayoutCagr cht, shp, fromIdx, toIdx, values, tops, fmt
        Case "diff": Ann_LayoutDiff cht, shp, fromIdx, toIdx, values, tops, fmt
        Case "conn": Ann_LayoutConn cht, shp, fromIdx, toIdx, tops
    End Select
End Sub

' Mostra/oculta a forma e seus companheiros (_lbl, _lvl e os tiques do colchete _t1/_t2).
Private Sub Ann_SetShapeVisible(cht As Chart, nm As String, vis As Boolean)
    Dim sfx As Variant, s As Shape, st As Long
    st = IIf(vis, msoTrue, msoFalse)
    On Error Resume Next
    For Each sfx In Array("", "_lbl", "_lvl", "_t1", "_t2")
        Set s = Nothing
        Set s = cht.Shapes(nm & sfx)
        If Not s Is Nothing Then
            If s.Visible <> st Then s.Visible = st
        End If
    Next
    On Error GoTo 0
End Sub

Private Function Ann_CountConnectors(cht As Chart) As Long
    Ann_CountConnectors = Ann_CountPrefix(cht, ANN_PREFIX & "conn_")
End Function

Private Function Ann_CountPrefix(cht As Chart, p As String) As Long
    Dim i As Long
    For i = 1 To cht.Shapes.Count
        If Left$(cht.Shapes(i).Name, Len(p)) = p Then Ann_CountPrefix = Ann_CountPrefix + 1
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
