Attribute VB_Name = "modStyleApply"
Option Explicit
' Estilos de grafico (1.12.0; spec 2026-10-23 secao 8.5): aplica um estilo validado por PAPEL. Camada reaplicada pelo core depois de tudo que
' recolore (Insert, toggles, opcoes, Alterar dados; refresh quando o carimbo muda). Opcoes do grafico: style (nome), stylever (8 hex do texto),
' stylestamp (carimbo do ultimo aplicar: versao, series, categorias, series que carregam a 1a e a ultima barra da cascata). Nunca mexe na
' formatacao que o estilo nao define. Hachura e cor de ponto sobrevivem a mudar valores (PX1); recriar o grafico perde: por isso a reaplicacao.
Private Const STYLE_ERR As Long = 7140
Private Const STYLE_MAX_LBL_SERIES As Long = 64     ' teto das auxiliares tqq_lbl_j procuradas (areas e Mekko)

' Preenchimento por papel: hachura (padrao valido) com a cor do padrao na frente e a do papel atras; senao cor solida. Nunca propaga (registrado).
Private Sub Style_Fill(ff As FillFormat, colorHex As String, patternName As String, patternHex As String)
    Dim fg As Long, bg As Long, pt As Long
    On Error GoTo Fail
    bg = StyleDef_ColorOf(colorHex)
    pt = StyleDef_PatternOf(patternName)
    If pt <> 0 Then
        fg = StyleDef_ColorOf(patternHex)
        If fg < 0 Then fg = RGB(0, 0, 0)
        If bg < 0 Then bg = RGB(255, 255, 255)
        ff.Visible = msoTrue
        ff.Patterned pt
        ff.ForeColor.RGB = fg
        ff.BackColor.RGB = bg
    ElseIf bg >= 0 Then
        ff.Visible = msoTrue
        ff.Solid
        ff.ForeColor.RGB = bg
    End If
    Exit Sub
Fail:
    Errors_Log "Style_Fill" & vbTab & Err.Number & vbTab & Err.Description
End Sub

Private Function Style_PointHasValue(s As Series, i As Long) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = s.Values
    If IsArray(v) Then
        If IsNumeric(v(i)) Then Style_PointHasValue = (CDbl(v(i)) <> 0)
    End If
    Err.Clear
End Function

' Ponto i de uma serie de papel: com valor = o papel especial (primeira/ultima: cor, hachura, borda); sem valor = volta a cor da serie e sem
' borda (um zero com borda vira linha na base). So age se o estilo define o papel especial. Nunca propaga.
Private Sub Style_EdgePoint(s As Series, i As Long, d As Object, role As String, seriesHex As String)
    Dim p As Point, b As String, hx As String
    If Len(StyleDef_Get(d, role)) = 0 And Len(StyleDef_Get(d, role & ".padrao")) = 0 And Len(StyleDef_Get(d, role & ".borda")) = 0 Then Exit Sub
    On Error GoTo Fail
    Set p = s.Points(i)
    If Style_PointHasValue(s, i) Then
        hx = StyleDef_Get(d, role)
        If Len(hx) = 0 Then hx = seriesHex
        Style_Fill p.Format.Fill, hx, StyleDef_Get(d, role & ".padrao"), StyleDef_Get(d, role & ".padrao.cor")
        b = StyleDef_Get(d, role & ".borda")
        If b = "nenhuma" Then
            p.Format.Line.Visible = msoFalse
        ElseIf Len(b) > 0 Then
            p.Format.Line.Visible = msoTrue
            p.Format.Line.ForeColor.RGB = StyleDef_ColorOf(b)
            p.Format.Line.Weight = 0.75
        End If
    Else
        If Len(seriesHex) > 0 Then Style_Fill p.Format.Fill, seriesHex, "", ""
        p.Format.Line.Visible = msoFalse
    End If
    Exit Sub
Fail:
    Errors_Log "Style_EdgePoint" & vbTab & role & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Ponto que ERA a ultima barra no carimbo anterior (a tabela ganhou ou perdeu linhas sem recriar o grafico: a formatacao por ponto ficaria
' presa ao indice antigo): volta a cor da serie, sem borda. Nunca propaga.
Private Sub Style_ResetPoint(s As Series, i As Long, seriesHex As String)
    Dim p As Point
    On Error GoTo Fail
    If i < 1 Or i > s.Points.Count Then Exit Sub
    Set p = s.Points(i)
    If Len(seriesHex) > 0 Then Style_Fill p.Format.Fill, seriesHex, "", ""
    p.Format.Line.Visible = msoFalse
    Exit Sub
Fail:
    Errors_Log "Style_ResetPoint" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Numero de categorias gravado no carimbo ("versao.series.categorias.portadoras"); 0 sem carimbo. Pura.
Private Function Style_StampCategories(stamp As String) As Long
    Dim parts() As String
    If Len(stamp) = 0 Then Exit Function
    parts = Split(stamp, ".")
    If UBound(parts) >= 2 Then Style_StampCategories = CLng(Val(parts(2)))
End Function

Private Sub Style_WaterfallRoles(c As Chart, d As Object)
    Dim s As Series, role As String, n As Long, hx As String, oldN As Long
    oldN = Style_StampCategories(Registry_GetOption(c, "stylestamp", ""))
    For Each s In c.SeriesCollection
        role = ""
        Select Case s.Name
            Case "tqq_uppos", "tqq_upneg": role = "aumento"
            Case "tqq_downpos", "tqq_downneg": role = "reducao"
            Case "tqq_total": role = "total"
        End Select
        If Len(role) > 0 Then
            hx = StyleDef_Get(d, role)
            If Len(hx) > 0 Then
                Style_Fill s.Format.Fill, hx, "", ""
                s.Format.Line.Visible = msoFalse
            End If
            n = s.Points.Count
            If oldN >= 2 And oldN < n Then
                If Len(StyleDef_Get(d, "ultima")) > 0 Or Len(StyleDef_Get(d, "ultima.padrao")) > 0 Or Len(StyleDef_Get(d, "ultima.borda")) > 0 Then Style_ResetPoint s, oldN, hx
            End If
            If n >= 1 Then Style_EdgePoint s, 1, d, "primeira", hx
            If n >= 2 Then Style_EdgePoint s, n, d, "ultima", hx
        End If
    Next
End Sub

' Cor (hex) da serie de dados k: ciclo nas cores serieN definidas; "" sem nenhuma. Pura.
Private Function Style_SeriesHex(d As Object, k As Long) As String
    Dim m As Long
    m = StyleDef_SeriesCount(d)
    If m = 0 Or k < 1 Then Exit Function
    Style_SeriesHex = StyleDef_Get(d, "serie" & (((k - 1) Mod m) + 1))
End Function

' Cor de serie k. Linhas: linha e marcador; area: preenchimento (a borda branca que separa as faixas fica); demais: preenchimento sem borda e,
' com rotulos no centro, fonte por contraste (a regra de Style_ApplySeriesColors). Nunca propaga.
Private Sub Style_SeriesColor(s As Series, d As Object, k As Long)
    Dim hx As String, rgbv As Long
    hx = Style_SeriesHex(d, k)
    If Len(hx) = 0 Then Exit Sub
    rgbv = StyleDef_ColorOf(hx)
    On Error GoTo Fail
    Select Case s.ChartType
        Case xlLine, xlLineMarkers, xlLineStacked, xlLineMarkersStacked
            s.Format.Line.ForeColor.RGB = rgbv
            s.Format.Fill.ForeColor.RGB = rgbv
        Case xlArea, xlAreaStacked, xlAreaStacked100
            Style_Fill s.Format.Fill, hx, "", ""
        Case Else
            Style_Fill s.Format.Fill, hx, "", ""
            s.Format.Line.Visible = msoFalse
            If s.HasDataLabels Then
                If s.DataLabels.Position = xlLabelPositionCenter Then s.DataLabels.Font.Color = Style_ContrastFont(rgbv)
            End If
    End Select
    Exit Sub
Fail:
    Errors_Log "Style_SeriesColor" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Series de dados pela ordem (auxiliares tqq_* nunca); cascata de varias series: as partes tqq_dn_j/tqq_upn_j/tqq_dnn_j com a cor da serie j e
' tqq_total com "total" (a regra de Waterfall_ApplyColors).
Private Sub Style_SeriesRoles(c As Chart, d As Object, waterfallMulti As Boolean)
    Dim s As Series, k As Long, j As Long, p As Long, nm As String
    For Each s In c.SeriesCollection
        nm = s.Name
        If waterfallMulti Then
            j = 0
            If nm = "tqq_total" Then
                If Len(StyleDef_Get(d, "total")) > 0 Then Style_Fill s.Format.Fill, StyleDef_Get(d, "total"), "", ""
            ElseIf Left$(nm, 7) = "tqq_dn_" Or Left$(nm, 8) = "tqq_upn_" Or Left$(nm, 8) = "tqq_dnn_" Then
                p = InStrRev(nm, "_")
                j = CLng(Val(Mid$(nm, p + 1)))
            ElseIf Not Style_IsHelperSeries(s) Then
                k = k + 1
                j = k
            End If
            If j > 0 Then Style_SeriesColor s, d, j
        ElseIf Not Style_IsHelperSeries(s) Then
            k = k + 1
            Style_SeriesColor s, d, k
        End If
    Next
End Sub

' Pizza e rosca: cor por fatia (serieN em ciclo) e a fonte dos rotulos pela regra da pizza. Nunca propaga.
Private Sub Style_PointRoles(c As Chart, d As Object)
    Dim s As Series, i As Long
    If StyleDef_SeriesCount(d) = 0 Then Exit Sub
    On Error GoTo Fail
    If c.SeriesCollection.Count = 0 Then Exit Sub
    Set s = c.SeriesCollection(1)
    For i = 1 To s.Points.Count
        Style_Fill s.Points(i).Format.Fill, Style_SeriesHex(d, i), "", ""
    Next
    Pie_ApplyLabelColors c
    Exit Sub
Fail:
    Errors_Log "Style_PointRoles" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Dispersao e bolhas: cor por grupo (marcador; bolha: preenchimento com a borda branca) e as tendencias do grupo, como Scatter_ApplyColors.
' Nunca propaga.
Private Sub Style_ScatterRoles(c As Chart, d As Object)
    Dim k As Long, s As Series, rgbv As Long, i As Long, isBubble As Boolean
    If StyleDef_SeriesCount(d) = 0 Then Exit Sub
    On Error GoTo Fail
    isBubble = (Registry_ChartType(c) = "bubble")
    For k = 1 To Scatter_SeriesCount(c)
        Set s = Scatter_DataSeries(c, k)
        rgbv = StyleDef_ColorOf(Style_SeriesHex(d, k))
        If Not s Is Nothing And rgbv >= 0 Then
            s.Format.Fill.Visible = msoTrue
            s.Format.Fill.ForeColor.RGB = rgbv
            If Not isBubble Then
                s.MarkerBackgroundColor = rgbv
                s.MarkerForegroundColor = rgbv
            End If
            For i = 1 To s.Trendlines.Count
                s.Trendlines(i).Format.Line.ForeColor.RGB = rgbv
            Next
        End If
    Next
    If isBubble Then Scatter_SyncBubbleLabels c
    Exit Sub
Fail:
    Errors_Log "Style_ScatterRoles" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Cascata comparativa: serie1/serie2 nas duas series de dados e nas partes negativas (tqq_aneg/tqq_bneg), como Wfp_ApplyColors. Nunca propaga.
Private Sub Style_PairRoles(c As Chart, d As Object)
    Dim s As Series
    If StyleDef_SeriesCount(d) = 0 Then Exit Sub
    On Error GoTo Fail
    Style_SeriesRoles c, d, False
    For Each s In c.SeriesCollection
        Select Case s.Name
            Case "tqq_aneg": Style_SeriesColor s, d, 1
            Case "tqq_bneg": Style_SeriesColor s, d, 2
        End Select
    Next
    Wfp_RefreshFontColors c
    Exit Sub
Fail:
    Errors_Log "Style_PairRoles" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Rotulos de segmento das areas e do Mekko (auxiliares tqq_lbl_j): fonte por contraste com a cor nova da serie j. Nunca propaga.
Private Sub Style_SegmentLabels(c As Chart)
    Dim j As Long, ls As Series, fc As Long, r As Long
    On Error Resume Next
    For j = 1 To STYLE_MAX_LBL_SERIES
        Set ls = Nothing
        Set ls = Chart_FindSeries(c, "tqq_lbl_" & j)
        If ls Is Nothing Then Exit For
        fc = Style_ContrastFont(Style_DataSeriesColor(c, j))
        For r = 1 To ls.Points.Count
            If ls.Points(r).HasDataLabel Then ls.Points(r).DataLabel.Font.Color = fc
        Next
    Next
    Err.Clear
End Sub

' Borboleta: os rotulos das duas series de dados por contraste com a cor nova (a regra de Bfly_ApplyColors). Nunca propaga.
Private Sub Style_ButterflyLabels(c As Chart)
    Dim k As Long, s As Series, fc As Long, i As Long
    On Error Resume Next
    For k = 1 To 2
        Set s = Nothing
        Set s = Bfly_SideSeries(c, k)
        If Not s Is Nothing Then
            fc = Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB))
            If s.HasDataLabels Then
                s.DataLabels.Font.Color = fc
            Else
                For i = 1 To s.Points.Count
                    If s.Points(i).HasDataLabel Then s.Points(i).DataLabel.Font.Color = fc
                Next
            End If
        End If
    Next
    Err.Clear
End Sub

' rotulo.negrito/tamanho em todo rotulo existente; rotulo.cor = "#RRGGBB" em todos; "auto" so nos rotulos dentro de barras de series de
' DADOS (as auxiliares de rotulo ja receberam o contraste da familia). Nunca propaga.
Private Sub Style_Labels(c As Chart, d As Object)
    Dim s As Series, i As Long, p As Point, col As String, pos As Long, bold As String, sz As String, helper As Boolean
    col = StyleDef_Get(d, "rotulo.cor"): bold = StyleDef_Get(d, "rotulo.negrito"): sz = StyleDef_Get(d, "rotulo.tamanho")
    If Len(col) = 0 And Len(bold) = 0 And Len(sz) = 0 Then Exit Sub
    On Error Resume Next
    For Each s In c.SeriesCollection
        helper = Style_IsHelperSeries(s)
        For i = 1 To s.Points.Count
            Set p = Nothing
            Set p = s.Points(i)
            If Not p Is Nothing Then
                If p.HasDataLabel Then
                    If Len(bold) > 0 Then p.DataLabel.Font.Bold = (bold = "1")
                    If Len(sz) > 0 Then p.DataLabel.Font.Size = CLng(sz)
                    If col = "auto" Then
                        If Not helper Then
                            pos = 0
                            pos = p.DataLabel.Position
                            If pos = xlLabelPositionCenter Or pos = xlLabelPositionInsideEnd Or pos = xlLabelPositionInsideBase Then p.DataLabel.Font.Color = Style_ContrastFont(CLng(p.Format.Fill.ForeColor.RGB))
                        End If
                    ElseIf Len(col) > 0 Then
                        p.DataLabel.Font.Color = StyleDef_ColorOf(col)
                    End If
                End If
            End If
            Err.Clear
        Next
    Next
    Err.Clear
End Sub

Private Sub Style_Line(lf As Object, d As Object, role As String)
    On Error Resume Next
    If Len(StyleDef_Get(d, role & ".cor")) > 0 Then
        lf.Visible = msoTrue
        lf.ForeColor.RGB = StyleDef_ColorOf(StyleDef_Get(d, role & ".cor"))
    End If
    If Len(StyleDef_Get(d, role & ".espessura")) > 0 Then lf.Weight = Val(StyleDef_Get(d, role & ".espessura"))
    Select Case StyleDef_Get(d, role & ".traco")
        Case "pontos": lf.DashStyle = msoLineRoundDot
        Case "tracos": lf.DashStyle = msoLineDash
        Case "solido": lf.DashStyle = msoLineSolid
    End Select
    Err.Clear
End Sub

' Conectores: formas tqq_ann_conn_* da cascata e SeriesLines das empilhadas. Nunca propaga.
Private Sub Style_Connectors(c As Chart, d As Object)
    Dim i As Long, g As ChartGroup
    If Len(StyleDef_Get(d, "conector.cor")) = 0 And Len(StyleDef_Get(d, "conector.traco")) = 0 And Len(StyleDef_Get(d, "conector.espessura")) = 0 Then Exit Sub
    On Error Resume Next
    For i = 1 To c.Shapes.Count
        If Left$(c.Shapes(i).Name, Len(ANN_PREFIX & "conn_")) = ANN_PREFIX & "conn_" Then Style_Line c.Shapes(i).Line, d, "conector"
    Next
    For Each g In c.ChartGroups
        If g.HasSeriesLines Then Style_Line g.SeriesLines.Format.Line, d, "conector"
    Next
    Err.Clear
End Sub

' Setas (spec 8.5: formas tqq_ann_cagr_*, tqq_ann_diff_*, tqq_ann_delta_*): linhas com seta.cor; caixas de texto (rotulo) com fundo e borda do
' rotulo. Conectores, indicador 100%, marcas do corte e colchetes do Gantt nao sao setas. Nunca propaga.
Private Sub Style_Arrows(c As Chart, d As Object)
    Dim i As Long, s As Shape, nm As String, b As String, hasTxt As Boolean
    If Len(StyleDef_Get(d, "seta.cor")) = 0 And Len(StyleDef_Get(d, "seta.rotulo.fundo")) = 0 And Len(StyleDef_Get(d, "seta.rotulo.borda")) = 0 Then Exit Sub
    On Error Resume Next
    For i = 1 To c.Shapes.Count
        Set s = c.Shapes(i)
        nm = s.Name
        If Left$(nm, Len(ANN_PREFIX & "cagr_")) = ANN_PREFIX & "cagr_" Or Left$(nm, Len(ANN_PREFIX & "diff_")) = ANN_PREFIX & "diff_" Or Left$(nm, Len(ANN_PREFIX & "delta_")) = ANN_PREFIX & "delta_" Then
            hasTxt = False
            hasTxt = (s.TextFrame2.HasText = msoTrue)
            If hasTxt Then
                If Len(StyleDef_Get(d, "seta.rotulo.fundo")) > 0 Then
                    s.Fill.Visible = msoTrue
                    s.Fill.Solid
                    s.Fill.ForeColor.RGB = StyleDef_ColorOf(StyleDef_Get(d, "seta.rotulo.fundo"))
                End If
                b = StyleDef_Get(d, "seta.rotulo.borda")
                If b = "nenhuma" Then
                    s.Line.Visible = msoFalse
                ElseIf Len(b) > 0 Then
                    s.Line.Visible = msoTrue
                    s.Line.ForeColor.RGB = StyleDef_ColorOf(b)
                End If
            ElseIf Len(StyleDef_Get(d, "seta.cor")) > 0 Then
                s.Line.ForeColor.RGB = StyleDef_ColorOf(StyleDef_Get(d, "seta.cor"))
            End If
        End If
        Err.Clear
    Next
    Err.Clear
End Sub

' Espaco: a mesma regra de ApplyOption "gap" (todos os grupos de colunas/barras; a comparativa acompanha a meia-largura). Nunca propaga.
Private Sub Style_Gap(c As Chart, gap As Long)
    Dim g As ChartGroup, n As Long
    On Error Resume Next
    For Each g In c.ChartGroups
        Select Case g.SeriesCollection(1).ChartType
            Case xlColumnClustered, xlColumnStacked, xlColumnStacked100, xlBarClustered, xlBarStacked, xlBarStacked100
                g.GapWidth = gap
                n = n + 1
        End Select
    Next
    If n = 0 Then c.ChartGroups(1).GapWidth = gap
    Registry_SetOption c, "gap", CStr(gap)
    If Kinds_Family(Registry_ChartType(c)) = "waterfallpair" Then Wfp_SyncGap c
    Err.Clear
End Sub

' Aplica o estilo d (validado) ao grafico. Chaves sem papel no tipo sao ignoradas. Gantt (1.12.0): so fonte, rotulos e espaco (as cores do Gantt
' seguem a paleta; backlog). Propaga so erro estrutural (quem chama trata).
Public Sub Style_ApplyTo(c As Chart, d As Object)
    Dim fam As String
    fam = Kinds_Family(Registry_ChartType(c))
    On Error Resume Next
    If Len(StyleDef_Get(d, "fonte")) > 0 Then c.ChartArea.Font.Name = StyleDef_Get(d, "fonte")
    If Len(StyleDef_Get(d, "fonte.tamanho")) > 0 Then c.ChartArea.Font.Size = CLng(StyleDef_Get(d, "fonte.tamanho"))
    Err.Clear
    On Error GoTo 0
    Select Case fam
        Case "waterfall"
            If Waterfall_SeriesCount(c) > 1 Then Style_SeriesRoles c, d, True Else Style_WaterfallRoles c, d
        Case "pie"
            Style_PointRoles c, d
        Case "scatter"
            Style_ScatterRoles c, d
        Case "waterfallpair"
            Style_PairRoles c, d
        Case "gantt"
        Case Else
            Style_SeriesRoles c, d, False
            If StyleDef_SeriesCount(d) > 0 Then
                If fam = "area" Or fam = "mekko" Then Style_SegmentLabels c
                If fam = "butterfly" Then Style_ButterflyLabels c
            End If
    End Select
    Style_Labels c, d
    Style_Connectors c, d
    Style_Arrows c, d
    If Len(StyleDef_Get(d, "eixo.categoria.linha")) > 0 And fam <> "pie" Then
        On Error Resume Next
        c.Axes(xlCategory).Format.Line.Visible = IIf(StyleDef_Get(d, "eixo.categoria.linha") = "1", msoTrue, msoFalse)
        Err.Clear
        On Error GoTo 0
    End If
    If Len(StyleDef_Get(d, "espaco")) > 0 Then
        If Kinds_SupportsChart(c, "gap") Then Style_Gap c, CLng(StyleDef_Get(d, "espaco"))
    End If
End Sub

' Series (nomes) que carregam a 1a e a ultima barra da cascata de uma serie ("" nos outros tipos): entra no carimbo (o dado mudando de sinal
' muda a serie da barra). Nunca propaga.
Private Function Style_CarrierKey(c As Chart) As String
    Dim s As Series, n As Long, a As String, b As String
    On Error Resume Next
    If Kinds_Family(Registry_ChartType(c)) <> "waterfall" Then Exit Function
    If Waterfall_SeriesCount(c) > 1 Then Exit Function
    For Each s In c.SeriesCollection
        Select Case s.Name
            Case "tqq_uppos", "tqq_upneg", "tqq_downpos", "tqq_downneg", "tqq_total"
                n = s.Points.Count
                If Style_PointHasValue(s, 1) Then a = a & s.Name
                If n > 1 Then
                    If Style_PointHasValue(s, n) Then b = b & s.Name
                End If
        End Select
    Next
    Style_CarrierKey = a & "/" & b
    Err.Clear
End Function

' Carimbo do estilo aplicado: versao.series.categorias.portadoras. Nunca propaga.
Public Function Style_Stamp(c As Chart) As String
    On Error Resume Next
    Style_Stamp = Registry_GetOption(c, "stylever", "") & "." & c.SeriesCollection.Count & "." & Geom_ChartCategoryCount(c) & "." & Style_CarrierKey(c)
End Function

' Estilo pelo nome aplicado e gravado nas opcoes do grafico. Embutido so-paleta (ou "") = a paleta antiga (ApplyOption) e sem style. A pasta
' passa a levar o estilo. Erro STYLE_ERR (vbObjectError + 7140) com a mensagem quando o nome nao existe ou o texto nao vale.
Public Sub Style_ApplyNamed(c As Chart, name As String, Optional preferUser As Boolean = True)
    Dim txt As String, d As Object, ign As Long, pal As String, nm As String
    pal = StyleDef_BuiltinPalette(name)
    If Len(pal) > 0 Or Len(name) = 0 Then
        If Len(pal) = 0 Then pal = "padrao"
        Opt_ApplyPalette c, pal                                     ' grava palette e tira o style
        Exit Sub
    End If
    txt = StyleStore_Get(name, Registry_Workbook(c), preferUser)
    If Len(txt) = 0 Then Err.Raise vbObjectError + STYLE_ERR, "Style", "Estilo não encontrado: " & name & "."
    Set d = StyleDef_Parse(txt, ign)
    If d Is Nothing Then Err.Raise vbObjectError + STYLE_ERR, "Style", "Estilo inválido: " & name & "."
    nm = StyleDef_SafeName(name)
    StyleStore_EnsureInBook Registry_Workbook(c), nm, txt
    If Registry_GetOption(c, "style", "") <> nm Then Registry_SetOption c, "style", nm
    If Registry_GetOption(c, "stylever", "") <> StyleDef_Version(txt) Then Registry_SetOption c, "stylever", StyleDef_Version(txt)
    Style_ApplyTo c, d
    Registry_SetOption c, "stylestamp", Style_Stamp(c)
End Sub

' Tira o estilo do grafico depois de quem chamou reaplicar a paleta (as cores das series ficam as da paleta). O que a paleta nao refaz volta ao
' padrao do ThinQQ: a 1a/ultima barra da cascata (hachura, borda, cor do ponto) toma a cor atual da sua serie e os conectores da cascata voltam
' a linha cinza continua de 0,75 pt. Fonte e rotulos ficam como estao. Nunca propaga.
Public Sub Style_Clear(c As Chart)
    Dim s As Series, n As Long, i As Long, hx As String
    On Error Resume Next
    If Len(Registry_GetOption(c, "style", "")) = 0 Then Exit Sub
    If Kinds_Family(Registry_ChartType(c)) = "waterfall" And Waterfall_SeriesCount(c) = 1 Then
        For Each s In c.SeriesCollection
            Select Case s.Name
                Case "tqq_uppos", "tqq_upneg", "tqq_downpos", "tqq_downneg", "tqq_total"
                    hx = StyleDef_HexOf(CLng(s.Format.Fill.ForeColor.RGB))
                    n = s.Points.Count
                    Style_ResetPoint s, 1, hx
                    If n >= 2 Then Style_ResetPoint s, n, hx
            End Select
            Err.Clear
        Next
    End If
    For i = 1 To c.Shapes.Count
        If Left$(c.Shapes(i).Name, Len(ANN_PREFIX & "conn_")) = ANN_PREFIX & "conn_" Then
            With c.Shapes(i).Line
                .Visible = msoTrue
                .DashStyle = msoLineSolid
                .ForeColor.RGB = RGB(127, 127, 127)
                .Weight = 0.75
            End With
        End If
        Err.Clear
    Next
    Registry_SetOption c, "style", ""
    Registry_SetOption c, "stylever", ""
    Registry_SetOption c, "stylestamp", ""
    Err.Clear
End Sub

' Reaplica o estilo gravado no grafico, pela copia da PASTA (quem recebe a pasta ve o mesmo visual). Nunca propaga erro (registrado).
Public Sub Style_Reapply(c As Chart)
    Dim nm As String
    On Error GoTo Fail
    nm = Registry_GetOption(c, "style", "")
    If Len(nm) = 0 Then Exit Sub
    Style_ApplyNamed c, nm, False
    Exit Sub
Fail:
    Errors_Log "Style_Reapply" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Core_SafeRefresh: so reaplica quando o carimbo mudou (linha nova, serie nova, sinal da 1a/ultima barra). Nunca propaga erro.
Public Sub Style_ReapplyIfNeeded(c As Chart)
    On Error GoTo Fail
    If Len(Registry_GetOption(c, "style", "")) = 0 Then Exit Sub
    If Registry_GetOption(c, "stylestamp", "") = Style_Stamp(c) Then Exit Sub
    Style_Reapply c
    Exit Sub
Fail:
    Errors_Log "Style_ReapplyIfNeeded" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Le o visual atual do grafico (inclusive cores pintadas a mao) como estilo: papeis da cascata ou cores das series de dados, hachura/borda da 1a
' barra, fonte, conectores. Devolve o texto canonico. Propaga erro estrutural.
Public Function Style_FromChart(c As Chart, name As String) As String
    Dim d As Object, s As Series, k As Long, fam As String, i As Long, p As Point, nm As String, gotFirst As Boolean
    Set d = CreateObject("Scripting.Dictionary")
    d("nome") = StyleDef_SafeName(name)
    fam = Kinds_Family(Registry_ChartType(c))
    On Error Resume Next
    d("fonte") = StyleDef_Validate("fonte", CStr(c.ChartArea.Font.Name))
    d("fonte.tamanho") = StyleDef_Validate("fonte.tamanho", CStr(CLng(c.ChartArea.Font.Size)))
    Err.Clear
    If fam = "waterfall" And Waterfall_SeriesCount(c) = 1 Then
        d("aumento") = StyleDef_HexOf(c.SeriesCollection("tqq_uppos").Format.Fill.ForeColor.RGB)
        d("reducao") = StyleDef_HexOf(c.SeriesCollection("tqq_downpos").Format.Fill.ForeColor.RGB)
        d("total") = StyleDef_HexOf(c.SeriesCollection("tqq_total").Format.Fill.ForeColor.RGB)
        Err.Clear
        For Each s In c.SeriesCollection
            nm = s.Name
            If Not gotFirst And (nm = "tqq_uppos" Or nm = "tqq_upneg" Or nm = "tqq_downpos" Or nm = "tqq_downneg" Or nm = "tqq_total") Then
                If Style_PointHasValue(s, 1) Then
                    Set p = Nothing
                    Set p = s.Points(1)
                    If Not p Is Nothing Then
                        gotFirst = True
                        If p.Format.Fill.Type = msoFillPatterned Then
                            d("primeira") = StyleDef_HexOf(p.Format.Fill.BackColor.RGB)
                            d("primeira.padrao") = StyleDef_PatternName(CLng(p.Format.Fill.Pattern))
                            d("primeira.padrao.cor") = StyleDef_HexOf(p.Format.Fill.ForeColor.RGB)
                        Else
                            d("primeira") = StyleDef_HexOf(p.Format.Fill.ForeColor.RGB)
                        End If
                        If p.Format.Line.Visible = msoTrue Then d("primeira.borda") = StyleDef_HexOf(p.Format.Line.ForeColor.RGB)
                    End If
                End If
            End If
            Err.Clear
        Next
    Else
        For Each s In c.SeriesCollection
            If Not Style_IsHelperSeries(s) And k < 8 Then
                k = k + 1
                Select Case s.ChartType
                    Case xlLine, xlLineMarkers, xlLineStacked, xlLineMarkersStacked
                        d("serie" & k) = StyleDef_HexOf(s.Format.Line.ForeColor.RGB)
                    Case Else
                        d("serie" & k) = StyleDef_HexOf(s.Format.Fill.ForeColor.RGB)
                End Select
            End If
            Err.Clear
        Next
    End If
    For i = 1 To c.Shapes.Count
        If Left$(c.Shapes(i).Name, Len(ANN_PREFIX & "conn_")) = ANN_PREFIX & "conn_" Then
            d("conector.cor") = StyleDef_HexOf(c.Shapes(i).Line.ForeColor.RGB)
            Exit For
        End If
    Next
    Err.Clear
    On Error GoTo 0
    Style_FromChart = StyleDef_Serialize(d)
End Function
