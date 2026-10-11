Attribute VB_Name = "modChartStyle"
Option Explicit
' Aparência padrão "limpa": sem grade, eixo de valores invisível (mas presente, para a geometria),
' sem bordas, fonte 9 pt. Paletas de 6 cores.

Public Function Style_Palette(name As String) As Variant
    Select Case LCase$(name)
        Case "cinza"
            Style_Palette = Array(RGB(64, 64, 64), RGB(127, 127, 127), RGB(166, 166, 166), RGB(191, 191, 191), RGB(217, 217, 217), RGB(89, 89, 89))
        Case "tema"
            Style_Palette = Array(-1, -2, -3, -4, -5, -6)    ' accent1..accent6 do tema
        Case Else
            Style_Palette = Array(RGB(31, 56, 100), RGB(68, 114, 196), RGB(157, 195, 230), RGB(165, 165, 165), RGB(0, 150, 136), RGB(237, 162, 36))
    End Select
End Function

Public Sub Style_SetColor(cf As ColorFormat, colorValue As Variant)
    If colorValue < 0 Then
        cf.ObjectThemeColor = msoThemeColorAccent1 + (-colorValue - 1)
    Else
        cf.RGB = CLng(colorValue)
    End If
End Sub

Public Function Style_IsHelperSeries(s As Series) As Boolean
    Style_IsHelperSeries = (Left$(s.Name, Len(HELPER_PREFIX)) = HELPER_PREFIX)
End Function

' Branco para fundos escuros, cinza-escuro para claros (luminância relativa simples).
Public Function Style_ContrastFont(rgbValue As Long) As Long
    Dim r As Long, g As Long, b As Long, lum As Double
    r = rgbValue And &HFF: g = (rgbValue \ &H100) And &HFF: b = (rgbValue \ &H10000) And &HFF
    lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    If lum < 0.55 Then Style_ContrastFont = RGB(255, 255, 255) Else Style_ContrastFont = RGB(64, 64, 64)
End Function

' Eixo de valores invisivel mas presente (a geometria das anotacoes le a escala dele): sem grade, marcas, rotulos e linha.
' Usado pelo eixo primario (Style_ApplyBase) e pelo eixo secundario da combinacao (Combo_SetSecondary).
Public Sub Style_HideValueAxis(ax As Axis)
    With ax
        .HasMajorGridlines = False
        .HasMinorGridlines = False
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabelPosition = xlTickLabelPositionNone
        .Format.Line.Visible = msoFalse
    End With
End Sub

' Gráficos de pizza/rosca não têm eixos: o bloco de eixos é pulado. Em barras (empilhadas, agrupadas, 100%) o eixo de categorias
' fica invertido para a primeira categoria aparecer no topo (como no think-cell).
Public Sub Style_ApplyBase(cht As Chart)
    Dim ct As Long, isPie As Boolean, isBar As Boolean
    On Error Resume Next
    ct = cht.ChartType
    If Err.Number <> 0 Then ct = 0: Err.Clear
    On Error GoTo 0
    isPie = (ct = xlPie Or ct = xlDoughnut Or ct = xlPieExploded Or ct = xlDoughnutExploded)
    isBar = (ct = xlBarStacked Or ct = xlBarClustered Or ct = xlBarStacked100)
    With cht
        .HasTitle = True    ' liga e desliga: marca o título automático (gráfico de 1 série) como excluído
        .HasTitle = False
        .ChartArea.Format.Line.Visible = msoFalse
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Font.Size = 9
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse
        If isPie Then Exit Sub
        .HasAxis(xlValue) = True
        Style_HideValueAxis .Axes(xlValue)
        ' Categorias de data: o Excel escolheria eixo de tempo (posições proporcionais às datas); a geometria das
        ' anotações assume uma fatia por categoria, então força eixo de texto.
        On Error Resume Next
        .Axes(xlCategory).CategoryType = xlCategoryScale
        On Error GoTo 0
        With .Axes(xlCategory)
            .MajorTickMark = xlTickMarkNone
            .MinorTickMark = xlTickMarkNone
            .TickLabelPosition = xlTickLabelPositionLow   ' rótulos de categoria abaixo do gráfico mesmo com valores negativos
            If isBar Then .ReversePlotOrder = True
            .Format.Line.ForeColor.RGB = RGB(166, 166, 166)
            .TickLabels.Font.Size = 9
            .TickLabels.Font.Color = RGB(64, 64, 64)
        End With
    End With
End Sub

' Cores das series de dados (as auxiliares tqq_* nao mudam). Series de area: preenchimento da paleta e borda branca que
' separa as faixas; a espessura da borda (pt) vem de borderWeight (> 0) ou, sem ele, da opcao "border" do grafico (padrao
' 1 pt). Ficar na OPCAO e nao so no parametro e o que faz uma troca de paleta (ApplyOption "palette") manter a borda de
' um grafico que usa outra espessura (o Mekko, na Tarefa 9, grava border=1.5); o parametro, quando informado, nao persiste.
Public Sub Style_ApplySeriesColors(cht As Chart, paletteName As String, Optional borderWeight As Double = 0)
    Dim pal As Variant, s As Series, k As Long, c As Variant, bw As Double
    pal = Style_Palette(paletteName)
    bw = borderWeight
    If bw <= 0 Then bw = Val(Replace(Registry_GetOption(cht, "border", "1"), ",", "."))      ' Val so entende o ponto decimal
    If bw <= 0 Then bw = 1
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            c = pal(k Mod 6)
            If s.ChartType = xlLine Or s.ChartType = xlLineMarkers Then
                Style_SetColor s.Format.Line.ForeColor, c
                s.Format.Line.Weight = 2
                s.MarkerStyle = xlMarkerStyleCircle
                s.MarkerSize = 5
                Style_SetColor s.Format.Fill.ForeColor, c
            ElseIf Style_IsAreaType(s.ChartType) Then
                Style_SetColor s.Format.Fill.ForeColor, c
                s.Format.Line.Visible = msoTrue
                s.Format.Line.ForeColor.RGB = RGB(255, 255, 255)
                s.Format.Line.Weight = bw
            Else
                Style_SetColor s.Format.Fill.ForeColor, c
                s.Format.Line.Visible = msoFalse
                If s.HasDataLabels Then If s.DataLabels.Position = xlLabelPositionCenter Then s.DataLabels.Font.Color = Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB))
            End If
            k = k + 1
        End If
    Next
End Sub

' Cores por PONTO da pizza e da rosca (uma serie so): VaryByCategories e a fatia k recebe pal((k - 1) Mod 6) (a 7a volta ao inicio da
' paleta), com borda branca de 1 pt que separa as fatias. Nao mexe nos rotulos: quem chama (Pie_ApplyColors) refaz a cor da fonte.
Public Sub Style_ApplyPointColors(cht As Chart, paletteName As String)
    Dim pal As Variant, s As Series, k As Long
    If cht.SeriesCollection.Count = 0 Then Exit Sub
    pal = Style_Palette(paletteName)
    Set s = cht.SeriesCollection(1)
    cht.ChartGroups(1).VaryByCategories = True
    For k = 1 To s.Points.Count
        Style_SetColor s.Points(k).Format.Fill.ForeColor, pal((k - 1) Mod 6)
    Next
    s.Format.Line.Visible = msoTrue
    s.Format.Line.ForeColor.RGB = RGB(255, 255, 255)
    s.Format.Line.Weight = 1
End Sub

' Grava a espessura da borda das areas (pt) na opcao "border" e a aplica as areas existentes. O valor e GRAVADO com ponto
' decimal em qualquer localidade (Str$ nunca usa a virgula), que e o que Style_ApplySeriesColors le. E o ponto de entrada para
' quem quer outra espessura (o Mekko, 1,5 pt): a troca de paleta passa a preserva-la.
Public Sub Style_SetAreaBorder(cht As Chart, weightPt As Double)
    Dim s As Series, txt As String
    If weightPt <= 0 Then Err.Raise 5, "Style_SetAreaBorder", "A espessura da borda deve ser maior que zero."
    txt = Trim$(Str$(weightPt))
    If Left$(txt, 1) = "." Then txt = "0" & txt                  ' Str$ omite o zero a esquerda (".75")
    Registry_SetOption cht, "border", txt
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            If Style_IsAreaType(s.ChartType) Then s.Format.Line.Weight = weightPt
        End If
    Next
End Sub

Public Function Style_IsAreaType(chartType As Long) As Boolean
    Style_IsAreaType = (chartType = xlArea Or chartType = xlAreaStacked Or chartType = xlAreaStacked100)
End Function

' Cor (RGB) de preenchimento atual da j-esima serie de dados (a j-esima nao auxiliar); 0 se nao existe.
Public Function Style_DataSeriesColor(cht As Chart, j As Long) As Long
    Dim s As Series, k As Long
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            k = k + 1
            If k = j Then Style_DataSeriesColor = CLng(s.Format.Fill.ForeColor.RGB): Exit Function
        End If
    Next
End Function

Public Function Style_NumberFormat(key As String) As String
    Select Case key
        Case "dec1": Style_NumberFormat = "#,##0.0;-#,##0.0;"
        Case "dec2": Style_NumberFormat = "#,##0.00;-#,##0.00;"
        Case "pct0": Style_NumberFormat = "0%;-0%;"
        Case "pct1": Style_NumberFormat = "0.0%;-0.0%;"
        Case "mil":  Style_NumberFormat = "#,##0,;-#,##0,;"
        Case "mi":   Style_NumberFormat = "#,##0.0,,;-#,##0.0,,;"
        Case Else:   Style_NumberFormat = "#,##0;-#,##0;"
    End Select
End Function

' fmt = CODIGO de formato (Style_NumberFormat). Pizza e rosca guardam a CHAVE (int, pct0, ...) e mostram participacao ou valor: sai por
' Pie_ApplyNumberFormat com a chave do registro (ApplyOption grava numfmt antes de despachar, entao a chave lida ja e a nova).
Public Sub Style_ApplyNumberFormat(cht As Chart, fmt As String)
    Dim s As Series, i As Long, segFmt As String, totFmt As String, useFmt As String
    If Kinds_Family(Registry_ChartType(cht)) = "pie" Then
        Pie_ApplyNumberFormat cht, Registry_GetOption(cht, "numfmt", "int")
        Exit Sub
    End If
    segFmt = fmt: totFmt = fmt
    If Kinds_Is100(Registry_ChartType(cht)) Then
        ' 100%: segmentos mostram participação (pct0/pct1); numfmt vale para os totais absolutos
        segFmt = Labels_SegmentFormat(cht): totFmt = Labels_TotalFormat(cht)
    End If
    Labels_ApplyCellFormat cht, totFmt
    For Each s In cht.SeriesCollection
        If s.Name = "tqq_vline" Then
            ' linha de valor: o rotulo so existe no ultimo ponto e mostra o texto da celula VLine (formato proprio: Deco_ApplyVLineFormat)
        ElseIf s.HasDataLabels And Not Labels_IsSourceFormatSeries(cht, s) Then      ' a linha da combinacao segue o formato da celula
            If s.Name = "tqq_total" Then useFmt = totFmt Else useFmt = segFmt
            Style_AssignNumberFormat s.DataLabels, useFmt
            ' rótulos ligados a células (por ponto) guardam formato próprio
            For i = 1 To s.Points.Count
                If s.Points(i).HasDataLabel Then Style_AssignNumberFormat s.Points(i).DataLabel, useFmt
            Next
        End If
    Next
End Sub

' Tabela de formatos: os códigos em inglês (Style_NumberFormat) são a fonte neutra. Formatos de rótulos
' de gráfico são interpretados na notação do idioma do Excel neste ambiente, então os rótulos são sempre
' gravados por NumberFormatLocal com o código derivado por Style_LocalizeCode (nunca por NumberFormat).
' (A regra "códigos em inglês" do plano vale para formatos de célula/intervalo.)
Public Function Style_LocalizeCode(fmt As String) As String
    Dim d As String, t As String
    d = Application.International(xlDecimalSeparator)
    t = Application.International(xlThousandsSeparator)
    If Len(d) <> 1 Then d = "."
    If Len(t) <> 1 Then t = ","
    If t = d Then t = IIf(d = ",", ".", ",")
    ' NBSP (separador de milhar de alguns idiomas) é usado literalmente no código
    Style_LocalizeCode = Replace(Replace(Replace(fmt, ",", ChrW(1)), ".", d), ChrW(1), t)
End Function

' Atribui o formato (código em inglês de Style_NumberFormat) a um DataLabels/DataLabel, de forma determinística.
Public Sub Style_AssignNumberFormat(target As Object, fmt As String)
    target.NumberFormatLocal = Style_LocalizeCode(fmt)
End Sub
