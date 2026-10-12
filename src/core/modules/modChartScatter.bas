Attribute VB_Name = "modChartScatter"
Option Explicit
' Dispersao (kind scatter) e bolhas (kind bubble), familia scatter: dois eixos de valor (xlXYScatter / xlBubble), UMA SERIE POR GRUPO
' (coluna de grupo opcional, a ultima da tabela), rotulo de cada ponto ligado ao nome (DataLabel.Formula), titulos dos eixos ligados aos
' cabecalhos (AxisTitle.Formula), linha de tendencia nativa (trend) e quadrantes nas medias (quad: Axes(xlCategory).CrossesAt = media de X
' desenha a linha vertical; Axes(xlValue).CrossesAt = media de Y, a horizontal; medido). Bolhas nao combinam com outros tipos de serie
' (o Excel converte o grafico inteiro): por isso nao ha auxiliares.
' Tabela: Nome | X | Y [| Grupo] (dispersao) e Nome | X | Y | Tamanho [| Grupo] (bolhas).
' Bloco (corpo): linha 1 = cabecalhos; n linhas: 1 Nome | 2 X | 3 Y | 4 Tam | 5 Grupo | 6..5+G Xg_k | 6+G..5+2G Yg_k | 6+2G..5+3G Sg_k,
' com Nome = IF(ISBLANK(src),"",src) (no formato numerico da celula de origem), X/Y = IF(ISNUMBER(src),src,NA()),
' Tam = IF(ISNUMBER(src),IF(src>0,src,NA()),NA()) (dispersao: 1), Grupo = =src&"" (o valor como texto), e
' Xg_k = IF(AND(Grupo="<nome k>",ISNUMBER(X),ISNUMBER(Y),ISNUMBER(Tam)),X,NA()) (idem Yg, Sg). DisplayBlanksAs = xlNotPlotted: o ponto NA
' nao e desenhado e o seu rotulo tambem nao (medido). Meta: E = AvgX, F = AvgY (AGGREGATE ignora os NA()).
' Um grupo que so aparece DEPOIS de inserir nao tem serie: "Alterar dados" recria (documentado no tutorial).
Public Const SC_COL_NAME As Long = 1
Public Const SC_COL_X As Long = 2
Public Const SC_COL_Y As Long = 3
Public Const SC_COL_SIZE As Long = 4
Public Const SC_COL_GROUP As Long = 5
Private Const SC_FIXED As Long = 5
Private Const SC_META_AVGX As Long = 5
Private Const SC_META_AVGY As Long = 6
Private Const SC_MAX_GROUPS As Long = 6
Private Const SC_ERR_COLS As String = "Dispersão: a tabela precisa de nome, X e Y (e, nas bolhas, tamanho)."
Private Const SC_ERR_GROUPS As String = "No máximo 6 grupos (coluna de grupo)."
Private Const SC_NO_GROUP As String = "(sem grupo)"
Private Const SC_AXIS_COLOR As Long = 10921638   ' RGB(166, 166, 166)
Private Const SC_TEXT_COLOR As Long = 5855577    ' RGB(89, 89, 89)
Private Const SC_LABEL_COLOR As Long = 4210752   ' RGB(64, 64, 64): nomes ao lado do ponto

Public Function Scatter_GroupCount(blk As Range) As Long
    Scatter_GroupCount = (blk.Columns.Count - SC_FIXED) \ 3
End Function

Public Function Scatter_XgCol(blk As Range, k As Long) As Long
    Scatter_XgCol = SC_FIXED + k
End Function

Public Function Scatter_YgCol(blk As Range, k As Long) As Long
    Scatter_YgCol = SC_FIXED + Scatter_GroupCount(blk) + k
End Function

Public Function Scatter_SgCol(blk As Range, k As Long) As Long
    Scatter_SgCol = SC_FIXED + 2 * Scatter_GroupCount(blk) + k
End Function

' Numero minimo de colunas de valores: 2 (X, Y) na dispersao, 3 (X, Y, tamanho) nas bolhas. Erro 6012 (aviso de Core_Insert).
Private Sub Scatter_CheckData(data As clsChartData, kind As String)
    If data.SeriesCount < IIf(kind = "bubble", 3, 2) Then Err.Raise vbObjectError + 6012, "Scatter_Insert", SC_ERR_COLS
End Sub

' Coluna da tabela (1 = categoria) que traz o grupo; 0 se a tabela nao tem coluna de grupo.
Private Function GroupColumn(data As clsChartData, kind As String) As Long
    Dim need As Long
    need = IIf(kind = "bubble", 4, 3)                ' colunas de valores obrigatorias + 1
    If data.SeriesCount >= need Then GroupColumn = need + 1
End Function

' Nomes distintos dos grupos na ordem em que aparecem (VALOR da celula como texto, igual ao =src&"" do bloco: o texto exibido de uma data
' ou de um numero formatado nao bateria com a formula; vazio = "(sem grupo)"); sem coluna de grupo, um unico grupo com
' o nome do cabecalho de Y. Mais de SC_MAX_GROUPS => 6013.
' A comparacao "=" de texto do Excel ignora maiusculas (medido: o grupo VERDADEIRO casou com o literal "Verdadeiro"): "a" e "A" sao o MESMO
' grupo (o primeiro que aparece da o nome), senao o ponto entraria nas duas series. Logico: o texto do Excel (=VERDADEIRO&"" da
' "VERDADEIRO"; CStr da "Verdadeiro").
Private Function GroupNames(data As clsChartData, kind As String) As Variant
    Dim gc As Long, i As Long, nm As String, names() As String, g As Long, k As Long, found As Boolean, v As Variant
    gc = GroupColumn(data, kind)
    ReDim names(1 To SC_MAX_GROUPS)
    If gc = 0 Then
        names(1) = data.SeriesName(2): g = 1
    Else
        For i = 1 To data.CategoryCount
            v = data.Source.Cells(i + 1, gc).Value2
            If IsError(v) Then nm = "" Else nm = CStr(v)
            If VarType(v) = vbBoolean Then nm = UCase$(nm)
            If Len(nm) = 0 Then nm = SC_NO_GROUP
            found = False
            For k = 1 To g
                If StrComp(names(k), nm, vbTextCompare) = 0 Then found = True: Exit For
            Next
            If Not found Then
                g = g + 1
                If g > SC_MAX_GROUPS Then Err.Raise vbObjectError + 6013, "Scatter_Insert", SC_ERR_GROUPS
                names(g) = nm
            End If
        Next
        If g = 0 Then names(1) = SC_NO_GROUP: g = 1
    End If
    ReDim Preserve names(1 To g)
    GroupNames = names
End Function

Public Function Scatter_BuildBlock(wb As Workbook, id As String, data As clsChartData, kind As String) As Range
    Dim blk As Range, body As Range, n As Long, i As Long, r As Long, names As Variant, g As Long, k As Long, gc As Long
    Dim sx As String, sy As String, ss As String, sg As String, x As String, y As String, t As String, grp As String, lit As String, cond As String
    Dim sn As String
    Scatter_CheckData data, kind
    names = GroupNames(data, kind)
    g = UBound(names)
    gc = GroupColumn(data, kind)
    n = data.CategoryCount
    Set blk = CalcSheet_AllocateBlock(wb, id, kind, Mid$(ChartData_ExternalAddress(data.Source), 2), n + 1, SC_FIXED + 3 * g)
    Set body = Block_Body(blk)
    body.Cells(1, SC_COL_NAME).Value = "Nome": body.Cells(1, SC_COL_X).Value = "X": body.Cells(1, SC_COL_Y).Value = "Y"
    body.Cells(1, SC_COL_SIZE).Value = "Tam": body.Cells(1, SC_COL_GROUP).Value = "Grupo"
    For k = 1 To g
        body.Cells(1, Scatter_XgCol(blk, k)).Value = "Xg" & k
        body.Cells(1, Scatter_YgCol(blk, k)).Value = "Yg" & k
        body.Cells(1, Scatter_SgCol(blk, k)).Value = "Sg" & k
    Next
    For i = 1 To n
        r = i + 1
        sx = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 2)), 2)
        sy = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 3)), 2)
        sn = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 1)), 2)
        body.Cells(r, SC_COL_NAME).Formula = "=IF(ISBLANK(" & sn & "),""""," & sn & ")"            ' nome vazio: sem "0" ao lado do ponto
        ' o rotulo ligado mostra o TEXTO da celula: o formato do nome de origem (data, ano) vem junto (como na borboleta e no Mekko)
        If Not IsNull(data.Source.Cells(r, 1).NumberFormat) Then body.Cells(r, SC_COL_NAME).NumberFormat = data.Source.Cells(r, 1).NumberFormat
        body.Cells(r, SC_COL_X).Formula = "=IF(ISNUMBER(" & sx & ")," & sx & ",NA())"
        body.Cells(r, SC_COL_Y).Formula = "=IF(ISNUMBER(" & sy & ")," & sy & ",NA())"
        If kind = "bubble" Then
            ss = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 4)), 2)
            body.Cells(r, SC_COL_SIZE).Formula = "=IF(ISNUMBER(" & ss & "),IF(" & ss & ">0," & ss & ",NA()),NA())"      ' IF aninhado: AND propagaria o erro
        Else
            body.Cells(r, SC_COL_SIZE).Formula = "=1"
        End If
        If gc > 0 Then
            sg = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, gc)), 2)
            body.Cells(r, SC_COL_GROUP).Formula = "=" & sg & "&"""""
        Else
            body.Cells(r, SC_COL_GROUP).Formula = "=""" & Replace(CStr(names(1)), """", """""") & """"
        End If
        x = body.Cells(r, SC_COL_X).Address(False, False): y = body.Cells(r, SC_COL_Y).Address(False, False)
        t = body.Cells(r, SC_COL_SIZE).Address(False, False): grp = body.Cells(r, SC_COL_GROUP).Address(False, False)
        For k = 1 To g
            lit = Replace(CStr(names(k)), """", """""")
            If CStr(names(k)) = SC_NO_GROUP And gc > 0 Then lit = ""
            cond = "AND(" & grp & "=""" & lit & """,ISNUMBER(" & x & "),ISNUMBER(" & y & "),ISNUMBER(" & t & "))"
            body.Cells(r, Scatter_XgCol(blk, k)).Formula = "=IF(" & cond & "," & x & ",NA())"
            body.Cells(r, Scatter_YgCol(blk, k)).Formula = "=IF(" & cond & "," & y & ",NA())"
            body.Cells(r, Scatter_SgCol(blk, k)).Formula = "=IF(" & cond & "," & t & ",NA())"
        Next
    Next
    blk.Cells(1, SC_META_AVGX).Formula = "=IFERROR(AGGREGATE(1,6," & body.Cells(2, SC_COL_X).Resize(n, 1).Address(True, True) & "),0)"
    blk.Cells(1, SC_META_AVGY).Formula = "=IFERROR(AGGREGATE(1,6," & body.Cells(2, SC_COL_Y).Resize(n, 1).Address(True, True) & "),0)"
    Set Scatter_BuildBlock = blk
End Function

Public Function Scatter_Insert(ws As Worksheet, data As clsChartData, kind As String) As ChartObject
    Dim id As String, co As ChartObject, cht As Chart, anchor As Range, blk As Range, body As Range, n As Long, g As Long, k As Long, s As Series
    Dim names As Variant, isBubble As Boolean, en As Long, es As String, ed As String
    If Kinds_Family(kind) <> "scatter" Then Err.Raise 5, "Scatter_Insert", "Tipo desconhecido: " & kind
    Scatter_CheckData data, kind                     ' recusas (6012/6013) antes de alocar: nada a desfazer
    names = GroupNames(data, kind)
    isBubble = (kind = "bubble")
    On Error GoTo Fail
    id = Registry_NewId()
    Set blk = Scatter_BuildBlock(ws.Parent, id, data, kind)
    Util_TestHookAfterBlock
    Set body = Block_Body(blk)
    n = data.CategoryCount: g = Scatter_GroupCount(blk)
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, Chart_FreeTop(ws, anchor.Left, anchor.Top, 460, 280), 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    For k = 1 To g
        Set s = cht.SeriesCollection.NewSeries
        s.Name = CStr(names(k))
        s.Values = Block_Ref(body.Cells(2, Scatter_YgCol(blk, k)).Resize(n, 1))
    Next
    cht.ChartType = IIf(isBubble, xlBubble, xlXYScatter)
    For k = 1 To g                                   ' depois do tipo: uma serie XY precisa existir como XY para aceitar X numerico
        Set s = cht.SeriesCollection(k)
        s.XValues = Block_Ref(body.Cells(2, Scatter_XgCol(blk, k)).Resize(n, 1))
        ' bolhas: o Excel 2019 aceita a referencia A1 (='_ThinQQ'!$J$5:$J$8) ao gravar BubbleSizes, apesar de a documentacao pedir R1C1
        ' (medido; lida de volta como _ThinQQ!$J$5:$J$8 e no 4o argumento de =SERIES)
        If isBubble Then s.BubbleSizes = Block_Ref(body.Cells(2, Scatter_SgCol(blk, k)).Resize(n, 1))
    Next
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlNotPlotted               ' ponto NA() (outro grupo, dado invalido) sem marcador e sem rotulo
    Style_ApplyBase cht
    Util_TestHookAfterChart
    Scatter_StyleAxes cht, data
    If isBubble Then
        cht.ChartGroups(1).BubbleScale = 60
        cht.ChartGroups(1).SizeRepresents = xlSizeIsArea
        cht.ChartGroups(1).ShowNegativeBubbles = False
    End If
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionRight
    cht.Legend.Font.Size = 9
    Registry_SetOption cht, "palette", "padrao"
    Registry_SetOption cht, "n", CStr(n)
    Registry_SetOption cht, "legend", "right"
    Registry_SetOption cht, "trend", "0"
    Registry_SetOption cht, "quad", "0"
    Scatter_ApplyColors cht, "padrao"
    Scatter_SetLabels cht, True
    Set Scatter_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' Eixos visiveis (os numeros nao estao em rotulos): linha cinza, rotulos de 9 pt na BORDA (Low: com os quadrantes a linha cruza no meio
' e os rotulos ficariam sobre os pontos), sem grade, titulos ligados aos cabecalhos de X e Y.
Private Sub Scatter_StyleAxes(cht As Chart, data As clsChartData)
    Dim ax As Axis, k As Long
    cht.HasAxis(xlCategory, xlPrimary) = True
    cht.HasAxis(xlValue, xlPrimary) = True
    For k = 1 To 2
        Set ax = cht.Axes(IIf(k = 1, xlCategory, xlValue), xlPrimary)
        With ax
            .HasMajorGridlines = False
            .HasMinorGridlines = False
            .MajorTickMark = xlTickMarkOutside
            .MinorTickMark = xlTickMarkNone
            .TickLabelPosition = xlTickLabelPositionLow
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = SC_AXIS_COLOR
            .Format.Line.DashStyle = msoLineSolid
            .TickLabels.Font.Size = 9
            .TickLabels.Font.Color = SC_TEXT_COLOR
            .HasTitle = True
            .AxisTitle.Formula = ChartData_ExternalAddress(data.SeriesNameCell(k))
            .AxisTitle.Font.Size = 9
            .AxisTitle.Font.Bold = False
            .AxisTitle.Font.Color = SC_TEXT_COLOR
        End With
    Next
End Sub

Public Function Scatter_SeriesCount(cht As Chart) As Long
    Dim s As Series
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then Scatter_SeriesCount = Scatter_SeriesCount + 1
    Next
End Function

Public Function Scatter_DataSeries(cht As Chart, k As Long) As Series
    Dim s As Series, i As Long
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            i = i + 1
            If i = k Then Set Scatter_DataSeries = s: Exit Function
        End If
    Next
End Function

' Rotulos: o nome de cada ponto, ligado a coluna Nome do bloco (ponto a ponto; nunca Labels_LinkPoints, que gravaria formato numerico nas
' celulas de texto). Dispersao: a direita do marcador, cinza; bolhas: no centro, fonte por contraste com a cor da serie, se o nome couber
' na bolha (senao a direita, cinza: Scatter_SyncBubbleLabels).
Public Sub Scatter_SetLabels(cht As Chart, turnOn As Boolean)
    Dim body As Range, n As Long, k As Long, i As Long, s As Series, isBubble As Boolean, c As Long
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    isBubble = (Registry_ChartType(cht) = "bubble")
    For k = 1 To Scatter_SeriesCount(cht)
        Set s = Scatter_DataSeries(cht, k)
        If Not turnOn Then
            s.HasDataLabels = False
        Else
            c = IIf(isBubble, Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB)), SC_LABEL_COLOR)
            For i = 1 To n
                If i <= s.Points.Count Then
                    ' isolado por ponto (como Labels_LinkPoints): so a dispersao foi medida com rotulo em ponto NA(); um ponto NA() de bolha
                    ' que recuse o rotulo nao pode desfazer o Insert inteiro
                    On Error Resume Next
                    With s.Points(i)
                        .HasDataLabel = True
                        .DataLabel.Formula = Block_Ref(body.Cells(i + 1, SC_COL_NAME))
                        .DataLabel.Position = IIf(isBubble, xlLabelPositionCenter, xlLabelPositionRight)
                        .DataLabel.Font.Size = 9
                        .DataLabel.Font.Bold = False
                        .DataLabel.Font.Color = c
                        ' (a quebra automatica do Excel fica: um nome longo vai em 2 linhas centradas na altura do ponto; numa linha so ele
                        ' invadia a legenda quando o ponto fica perto da borda direita, medido no PNG)
                    End With
                    If Err.Number <> 0 Then Errors_Log "Scatter_SetLabels" & vbTab & s.Name & vbTab & i & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
                    On Error GoTo 0
                End If
            Next
        End If
    Next
    Registry_SetOption cht, "labels", IIf(turnOn, "1", "0")
    If turnOn And isBubble Then Scatter_SyncBubbleLabels cht       ' nome que nao cabe na bolha: a direita, cinza-escuro
End Sub

' Cores por serie (grupo): marcadores XY so aceitam a cor por MarkerBackground/ForegroundColor (medido na Fase 2); as bolhas so pelo
' preenchimento (MarkerBackgroundColor numa serie de bolhas: erro 1004 "propriedade nao valida para o tipo de grafico", medido); a
' tendencia (se existe) e os rotulos das bolhas acompanham. Rotulos das bolhas ponto a ponto: Series.HasDataLabels le False quando o
' usuario apagou um rotulo e escrever em Series.DataLabels recriaria os apagados.
Public Sub Scatter_ApplyColors(cht As Chart, paletteName As String)
    Dim pal As Variant, k As Long, s As Series, c As Long, isBubble As Boolean, i As Long
    pal = Style_Palette(paletteName)
    isBubble = (Registry_ChartType(cht) = "bubble")
    For k = 1 To Scatter_SeriesCount(cht)
        Set s = Scatter_DataSeries(cht, k)
        c = Style_PaletteRGB(cht, pal((k - 1) Mod 6))
        s.Format.Fill.Visible = msoTrue
        s.Format.Fill.ForeColor.RGB = c
        If isBubble Then
            s.Format.Line.Visible = msoTrue
            s.Format.Line.ForeColor.RGB = RGB(255, 255, 255)
            s.Format.Line.Weight = 1
        Else
            s.MarkerBackgroundColor = c
            s.MarkerForegroundColor = c
            s.MarkerStyle = xlMarkerStyleCircle
            s.MarkerSize = 7
            s.Format.Line.Visible = msoFalse         ' sem linha ligando os pontos
        End If
        For i = 1 To s.Trendlines.Count
            s.Trendlines(i).Format.Line.ForeColor.RGB = c
        Next
    Next
    If isBubble Then Scatter_SyncBubbleLabels cht     ' a cor dos rotulos depende da cor nova (contraste) e de onde o rotulo esta
End Sub

' Bolhas: o nome fica DENTRO da bolha (centro, fonte por contraste com a cor do grupo) so quando cabe no diametro; senao vai para a direita,
' em cinza-escuro, como na dispersao (branco por contraste sobre o fundo branco sumiria). Diametro = Point.Width (medido: devolve o tamanho
' atual sem o grafico ser desenhado, inclusive logo depois de mudar os dados ou o tamanho do grafico); texto pela largura estimada
' (Labels_EstimateWidth). Chamado ao ligar os rotulos, ao trocar a paleta e em todo refresh (Scatter_RefreshAfterDataChange). So mexe em
' rotulos que existem (o apagado pelo usuario continua apagado) e so escreve o que mudou; cada ponto e isolado: uma leitura que falha deixa
' o ponto como esta (nunca decide com o valor de outro ponto). Ponto nao plotado (diametro 0) fica como esta. Nunca propaga erro.
Public Sub Scatter_SyncBubbleLabels(cht As Chart)
    Dim k As Long, i As Long, s As Series, fc As Long, has As Boolean, d As Double, txt As String, wantPos As Long, wantColor As Long, ok As Boolean
    On Error Resume Next
    If Registry_ChartType(cht) <> "bubble" Then Exit Sub
    For k = 1 To Scatter_SeriesCount(cht)
        Set s = Nothing
        Set s = Scatter_DataSeries(cht, k)
        fc = Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB))
        ok = (Err.Number = 0)
        If Not ok Then Errors_Log "Scatter_SyncBubbleLabels" & vbTab & k & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
        If ok Then
            For i = 1 To s.Points.Count
                has = False: d = 0: txt = vbNullString
                has = s.Points(i).HasDataLabel
                If Err.Number = 0 And has Then d = s.Points(i).Width
                If Err.Number = 0 And has And d > 0 Then txt = s.Points(i).DataLabel.Text
                If Err.Number = 0 And has And d > 0 Then
                    If Labels_EstimateWidth(txt) <= d Then
                        wantPos = xlLabelPositionCenter: wantColor = fc
                    Else
                        wantPos = xlLabelPositionRight: wantColor = SC_LABEL_COLOR
                    End If
                    With s.Points(i).DataLabel
                        If .Position <> wantPos Then .Position = wantPos
                        If .Font.Color <> wantColor Then .Font.Color = wantColor
                    End With
                End If
                If Err.Number <> 0 Then Errors_Log "Scatter_SyncBubbleLabels" & vbTab & s.Name & vbTab & i & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
            Next
        End If
    Next
End Sub

' Gancho de teste (Util_FailThinSyncForTest = "scatter-<ponto>:<chave>", como o bfly- da borboleta): falha simulada num ponto do codigo para
' testar o caminho de erro. Inerte fora do build Test.
Private Sub ScatterTestHook(site As String, key As String)
    If BUILD_CONFIG = "Test" Then
        If Len(Util_FailThinSyncForTest) > 0 Then
            If Util_FailThinSyncForTest = "scatter-" & site & ":" & key Then Err.Raise vbObjectError + 9996, "ThinQQ", "Falha simulada (teste) " & site & " " & key
        End If
    End If
End Sub

' Linha de tendencia linear (regressao nativa do Excel) por serie, tracejada, na cor da serie; as entradas de legenda das tendencias sao
' apagadas (Scatter_HideTrendEntries). Desligar apaga as tendencias.
' Transacional: uma falha no meio de ligar desfaz as tendencias ja criadas (nada pela metade, opcao "0"); em qualquer falha as entradas de
' legenda das tendencias que sobraram sao apagadas antes de o erro seguir para o chamador.
Public Sub Scatter_SetTrend(cht As Chart, turnOn As Boolean)
    Dim k As Long, s As Series, t As Trendline, pal As Variant, en As Long, es As String, ed As String
    On Error GoTo Fail
    pal = Style_Palette(Registry_GetOption(cht, "palette", "padrao"))
    For k = 1 To Scatter_SeriesCount(cht)
        Set s = Scatter_DataSeries(cht, k)
        Do While s.Trendlines.Count > 0
            s.Trendlines(1).Delete
        Loop
        If turnOn Then
            Set t = s.Trendlines.Add(xlLinear)
            t.Format.Line.ForeColor.RGB = Style_PaletteRGB(cht, pal((k - 1) Mod 6))
            t.Format.Line.Weight = 1
            t.Format.Line.DashStyle = msoLineDash
            ScatterTestHook "trend", CStr(k)
        End If
    Next
    Registry_SetOption cht, "trend", IIf(turnOn, "1", "0")
    Scatter_HideTrendEntries cht
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error GoTo -1                                 ' encerra o tratamento do erro: sem isso o Resume Next abaixo nao protegeria a limpeza
    On Error Resume Next
    If turnOn Then
        For k = 1 To Scatter_SeriesCount(cht)
            Set s = Scatter_DataSeries(cht, k)
            Do While s.Trendlines.Count > 0
                s.Trendlines(1).Delete
                If Err.Number <> 0 Then Exit Do
            Loop
            Err.Clear
        Next
        Registry_SetOption cht, "trend", "0"
    End If
    Scatter_HideTrendEntries cht
    On Error GoTo 0
    Err.Raise en, es, ed
End Sub

' As entradas de legenda das tendencias vem DEPOIS das entradas das series (medido): apaga as que excedem o numero de series. Idempotente;
' religar a legenda (Deco_SetLegend) as traz de volta, e o refresh as apaga de novo. Nunca propaga erro.
Public Sub Scatter_HideTrendEntries(cht As Chart)
    Dim ns As Long
    On Error Resume Next
    If Not cht.HasLegend Then Exit Sub
    ns = Scatter_SeriesCount(cht)
    Do While cht.Legend.LegendEntries.Count > ns
        cht.Legend.LegendEntries(cht.Legend.LegendEntries.Count).Delete
        If Err.Number <> 0 Then Exit Do
    Loop
End Sub

' Quadrantes: os eixos cruzam nas medias (Axes(xlCategory).CrossesAt = media de X: linha vertical; Axes(xlValue).CrossesAt = media de Y:
' horizontal), tracejados; os rotulos continuam na borda (TickLabelPosition Low). Desligar volta ao cruzamento automatico e a linha solida.
Public Sub Scatter_SetQuad(cht As Chart, turnOn As Boolean)
    Registry_SetOption cht, "quad", IIf(turnOn, "1", "0")
    If turnOn Then
        Scatter_ApplyQuadrants cht
    Else
        cht.Axes(xlCategory, xlPrimary).Crosses = xlAxisCrossesAutomatic
        cht.Axes(xlValue, xlPrimary).Crosses = xlAxisCrossesAutomatic
        cht.Axes(xlCategory, xlPrimary).Format.Line.DashStyle = msoLineSolid
        cht.Axes(xlValue, xlPrimary).Format.Line.DashStyle = msoLineSolid
    End If
End Sub

' Reaplica o cruzamento nas medias (meta E/F). So escreve quando mudou (chamada a cada refresh). Sem o ThinQQ as medias ficam no ultimo
' valor aplicado (como o intervalo do eixo do Gantt).
Public Sub Scatter_ApplyQuadrants(cht As Chart)
    Dim blk As Range, ax As Axis, v As Variant, k As Long
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    For k = 1 To 2
        v = blk.Cells(1, IIf(k = 1, SC_META_AVGX, SC_META_AVGY)).Value2
        If IsError(v) Or IsEmpty(v) Then Exit Sub
        Set ax = cht.Axes(IIf(k = 1, xlCategory, xlValue), xlPrimary)
        If ax.Crosses <> xlAxisCrossesCustom Or Abs(CDbl(ax.CrossesAt) - CDbl(v)) > 0.000000001 * (1 + Abs(CDbl(v))) Then ax.CrossesAt = CDbl(v)
        If ax.TickLabelPosition <> xlTickLabelPositionLow Then ax.TickLabelPosition = xlTickLabelPositionLow
        If ax.Format.Line.DashStyle <> msoLineDash Then ax.Format.Line.DashStyle = msoLineDash
    Next
End Sub

' O que na dispersao depende dos DADOS e muda depois de inserir (Core_SafeRefresh): as medias dos quadrantes e as entradas de legenda das
' tendencias. Nunca propaga erro.
Public Sub Scatter_RefreshAfterDataChange(cht As Chart)
    On Error Resume Next
    If Kinds_Family(Registry_ChartType(cht)) <> "scatter" Then Exit Sub
    If Registry_GetOption(cht, "quad", "0") = "1" Then Scatter_ApplyQuadrants cht
    If Err.Number <> 0 Then Errors_Log "Scatter_RefreshAfterDataChange/quad" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Scatter_HideTrendEntries cht
    Scatter_SyncBubbleLabels cht                     ' bolhas: o tamanho mudou com os dados ou com o grafico; o nome entra ou sai da bolha
End Sub

' Numero de linhas da tabela de origem entre o primeiro e o ultimo ponto, lido nas formulas da coluna X do bloco
' (=IF(ISNUMBER('Planilha'!$B$7),...)). As series apontam para o BLOCO (sempre n linhas), nao para a tabela: o tamanho dela so aparece nessas
' referencias (como no Mekko). Linhas inseridas no meio da tabela afastam as duas pontas; uma linha apagada no meio as aproxima; a primeira ou a
' ultima apagada vira #REF!. -1 se nao for possivel ler (Core_ChartSourceMismatch trata como tabela mudada). Nunca propaga erro.
Public Function Scatter_SourceRowsSpanned(cht As Chart) As Long
    Dim blk As Range, body As Range, n As Long, r1 As Long, r2 As Long
    On Error GoTo Fail
    Scatter_SourceRowsSpanned = -1
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    If Kinds_Family(CStr(blk.Cells(1, 2).Value)) <> "scatter" Then Exit Function
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    If n < 1 Then Exit Function
    r1 = SourceRefRow(CStr(body.Cells(2, SC_COL_X).Formula))
    r2 = SourceRefRow(CStr(body.Cells(n + 1, SC_COL_X).Formula))
    If r1 < 1 Or r2 < r1 Then Exit Function
    Scatter_SourceRowsSpanned = r2 - r1 + 1
    Exit Function
Fail:
    Scatter_SourceRowsSpanned = -1
End Function

' Linha da referencia a tabela numa formula do bloco: os digitos depois do ultimo "$" do endereco que segue o ULTIMO "!" (o nome da planilha
' pode ter "!" ou ","; o endereco termina no primeiro caractere que nao e letra, digito ou "$"). 0 sem referencia legivel (ex.: #REF!).
Private Function SourceRefRow(f As String) As Long
    Dim p As Long, i As Long, ch As String, addr As String
    If InStr(f, "#REF!") > 0 Then Exit Function
    p = InStrRev(f, "!")
    If p = 0 Then Exit Function
    For i = p + 1 To Len(f)
        ch = Mid$(f, i, 1)
        If Not (ch Like "[A-Za-z0-9$]") Then Exit For
        addr = addr & ch
    Next
    p = InStrRev(addr, "$")
    If p = 0 Then Exit Function
    SourceRefRow = CLng(Val(Mid$(addr, p + 1)))
End Function
