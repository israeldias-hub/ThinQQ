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
' Tam = IF(ISNUMBER(src),IF(src>0,src,NA()),NA()) (dispersao: 1), Grupo = LEFT(IFERROR(src&"",""),255) (o valor como texto; erro = vazio), e
' Xg_k = IF(AND(Grupo="<nome k>",ISNUMBER(X),ISNUMBER(Y),ISNUMBER(Tam)),X,NA()) (idem Yg, Sg). DisplayBlanksAs = xlNotPlotted: o ponto NA
' nao e desenhado e o seu rotulo tambem nao (medido). Meta: E = AvgX, F = AvgY sobre Xg e Yg (so os pontos plotados; AGGREGATE ignora os NA()).
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
Private Const SC_MAX_NAME As Long = 255          ' literal de texto numa formula e Series.Name: no maximo 255 caracteres (medido)
Private Const SC_AXIS_COLOR As Long = 10921638   ' RGB(166, 166, 166)
Private Const SC_TEXT_COLOR As Long = 5855577    ' RGB(89, 89, 89)
Private Const SC_LABEL_COLOR As Long = 4210752   ' RGB(64, 64, 64): nomes ao lado do ponto
Private mBubbleSyncBusy As Boolean                ' Scatter_SyncBubbleLabels em andamento: a reentrada (Calculate de uma escrita) sai
Private mBubbleSyncRuns As Long                   ' gancho de teste: passadas de Scatter_SyncBubbleLabels que fizeram o trabalho
Private Const SC_TIGHT_FRAC As Double = 0.6667  ' escala apertada quando a faixa dos dados e menor que 2/3 do maior |valor| (Scatter_NiceRange)
Private Const SC_MAX_STEPS As Long = 8           ' no maximo 8 intervalos na escala apertada
Private Const SC_AXIS_PAD As Double = 0.05       ' folga de cada lado da escala apertada (fracao da faixa dos dados)

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
        names(1) = Left$(data.SeriesName(2), SC_MAX_NAME): g = 1
    Else
        For i = 1 To data.CategoryCount
            v = data.Source.Cells(i + 1, gc).Value2
            If IsError(v) Then nm = "" Else nm = CStr(v)
            If VarType(v) = vbBoolean Then nm = UCase$(nm)
            nm = Left$(nm, SC_MAX_NAME)                 ' o bloco compara os mesmos 255 primeiros caracteres (LEFT)
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

' Nome da serie (legenda) de um grupo: o Excel le um Series.Name que comeca com "=" como formula (erro 1004, medido) e o ThinQQ tomaria um
' que comeca com "tqq_" por serie auxiliar (Style_IsHelperSeries); os dois ganham um espaco na frente. No maximo SC_MAX_NAME caracteres.
' So o nome exibido muda: o bloco compara o texto original do grupo.
Private Function SeriesDisplayName(nm As String) As String
    Dim s As String
    s = nm
    If Left$(s, 1) = "=" Or Left$(s, Len(HELPER_PREFIX)) = HELPER_PREFIX Then s = " " & s
    SeriesDisplayName = Left$(s, SC_MAX_NAME)
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
            ' erro na celula de grupo = vazio = "(sem grupo)", como GroupNames (sem o IFERROR o erro chegava as colunas por grupo como #DIV/0!,
            ' que a serie le como 0: medido, Series.Values devolveu 0 no ponto, e nao vazio como o #N/D)
            body.Cells(r, SC_COL_GROUP).Formula = "=LEFT(IFERROR(" & sg & "&"""",""""),255)"
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
    ' medias so dos pontos PLOTADOS: as colunas por grupo (Xg, Yg) ja sao NA() fora do grupo e com X/Y invalido, tamanho <= 0 ou grupo
    ' com erro; cada linha tem no maximo um valor nelas (um grupo so)
    blk.Cells(1, SC_META_AVGX).Formula = "=IFERROR(AGGREGATE(1,6," & body.Cells(2, Scatter_XgCol(blk, 1)).Resize(n, g).Address(True, True) & "),0)"
    blk.Cells(1, SC_META_AVGY).Formula = "=IFERROR(AGGREGATE(1,6," & body.Cells(2, Scatter_YgCol(blk, 1)).Resize(n, g).Address(True, True) & "),0)"
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
        s.Name = SeriesDisplayName(CStr(names(k)))
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
    Scatter_SyncAxisFormats cht
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
    Scatter_SyncAxes cht                             ' depois do tamanho das bolhas e da legenda: a folga das bolhas depende deles
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

' Celula de origem de uma formula X/Y do bloco (=IF(ISNUMBER(<ref>),<ref>,NA())): a referencia aparece duas vezes, entao o tamanho dela sai
' do tamanho da formula (o Excel pode tirar as aspas do nome da planilha, mas o faz nas duas ocorrencias). Nothing se a formula nao tem essa
' forma ou a referencia nao resolve (#REF!).
Private Function SourceCellOf(f As String, wb As Workbook) As Range
    Const PFX As String = "=IF(ISNUMBER("
    Const SFX As String = ",NA())"
    Dim l As Long, ref As String, bang As Long, sh As String
    On Error GoTo Fail
    If Left$(f, Len(PFX)) <> PFX Or Right$(f, Len(SFX)) <> SFX Then Exit Function
    l = Len(f) - Len(PFX) - Len(SFX) - 2
    If l < 4 Or (l Mod 2) <> 0 Then Exit Function
    l = l \ 2
    ref = Mid$(f, Len(PFX) + 1, l)
    If Mid$(f, Len(PFX) + l + 1, 2) <> ")," Or Mid$(f, Len(PFX) + l + 3, l) <> ref Then Exit Function
    bang = InStrRev(ref, "!")
    If bang < 2 Then Exit Function
    sh = Left$(ref, bang - 1)
    If Left$(sh, 1) = "'" And Right$(sh, 1) = "'" Then sh = Replace(Mid$(sh, 2, Len(sh) - 2), "''", "'")
    Set SourceCellOf = wb.Worksheets(sh).Range(Mid$(ref, bang + 1))
    Exit Function
Fail:
    Set SourceCellOf = Nothing
End Function

' Primeira celula NUMERICA da coluna de origem de X (col = SC_COL_X) ou Y (SC_COL_Y), pelas formulas do bloco (acompanham a tabela movida ou
' renomeada). Nothing se nenhuma for numerica ou legivel. Para no primeiro numero (no refresh, quase sempre a primeira linha).
Private Function SourceNumberCell(cht As Chart, col As Long) As Range
    Dim blk As Range, body As Range, i As Long, n As Long, c As Range, v As Variant
    On Error GoTo Fail
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    For i = 1 To n
        Set c = SourceCellOf(CStr(body.Cells(i + 1, col).Formula), Registry_Workbook(cht))
        If Not c Is Nothing Then
            v = c.Value2
            If Not IsError(v) And Not IsEmpty(v) And VarType(v) <> vbString And VarType(v) <> vbBoolean Then
                If IsNumeric(v) Then Set SourceNumberCell = c: Exit Function
            End If
        End If
    Next
    Exit Function
Fail:
    Set SourceNumberCell = Nothing
End Function

' Formato dos numeros dos eixos X e Y = o da primeira celula numerica da coluna de origem (sem isso a margem em % aparecia 0,3 / 0,325).
' Usa o NumberFormatLocal da CELULA: ja na notacao do Excel desta maquina, inclusive datas e textos entre aspas, que Style_LocalizeCode nao
' traduz. Geral (ou nenhuma celula numerica) = o padrao do eixo novo: ligado a origem (o bloco, em Geral). Nao grava o codigo "General" por
' NumberFormat: medido, isso (com um grafico movido para folha de grafico na mesma sessao) fazia o DataLabels.NumberFormat de OUTROS graficos
' devolver o codigo local. O formato aplicado fica gravado (xsrc/ysrc e o devolvido pelo eixo em xfmt/yfmt, codificados) e o eixo so e
' reescrito enquanto e do ThinQQ: ligado a origem ou mostrando o formato gravado. Um formato diferente posto pelo usuario fica (como a escala:
' xscale/yscale); devolvido a "ligado a origem", o ThinQQ volta a seguir a coluna. Na insercao e em todo refresh; so escreve o que mudou.
' Nunca propaga erro.
Public Sub Scatter_SyncAxisFormats(cht As Chart)
    SyncAxisFormat cht, 1
    SyncAxisFormat cht, 2
End Sub

' Um eixo (1 = X, 2 = Y) de Scatter_SyncAxisFormats. Duas opcoes por eixo: xsrc/ysrc = o formato da coluna que o ThinQQ aplicou e xfmt/yfmt =
' o formato que o EIXO devolveu depois de gravado (medido: o eixo nem sempre devolve o que recebeu: #.##0,00 "R$" volta #.##0,00 R$ e
' [$R$-416] volta [$R$-pt-BR]). O eixo e do ThinQQ quando esta ligado a origem ou mostra exatamente o que devolveu; e reescrito quando a
' coluna muda (formato diferente do aplicado). Transacional: as opcoes so sao gravadas depois de o eixo aceitar o formato; uma falha devolve o
' eixo ao que era (ligado ou o formato anterior) e as opcoes ficam. Nunca propaga erro.
Private Sub SyncAxisFormat(cht As Chart, k As Long)
    Dim c As Range, ax As Axis, want As String, key As String, srcKey As String, stored As String, src As String, linked As Boolean
    Dim cur As String, started As Boolean, back As String
    On Error GoTo Fail
    key = IIf(k = 1, "xfmt", "yfmt"): srcKey = IIf(k = 1, "xsrc", "ysrc")
    Set c = SourceNumberCell(cht, IIf(k = 1, SC_COL_X, SC_COL_Y))
    Set ax = cht.Axes(IIf(k = 1, xlCategory, xlValue), xlPrimary)
    stored = OptDecode(Registry_GetOption(cht, key, "")): src = OptDecode(Registry_GetOption(cht, srcKey, ""))
    linked = ax.TickLabels.NumberFormatLinked
    cur = CStr(ax.TickLabels.NumberFormatLocal)
    If Not (linked Or (Len(stored) > 0 And cur = stored)) Then Exit Sub         ' formato do usuario: fica
    If Not c Is Nothing Then
        If CStr(c.NumberFormat) <> "General" Then want = CStr(c.NumberFormatLocal)
    End If
    If Len(want) = 0 Then
        If Not linked Then
            started = True
            ax.TickLabels.NumberFormatLinked = True                             ' padrao do eixo novo (Geral, ligado ao bloco)
            ScatterTestHook "fmt", CStr(k)
        End If
        If Len(stored) > 0 Then Registry_SetOption cht, key, ""
        If Len(src) > 0 Then Registry_SetOption cht, srcKey, ""
    ElseIf linked Or want <> src Then
        started = True
        ax.TickLabels.NumberFormatLocal = want
        ScatterTestHook "fmt", CStr(k)
        back = CStr(ax.TickLabels.NumberFormatLocal)
        Registry_SetOption cht, key, OptEncode(back)
        Registry_SetOption cht, srcKey, OptEncode(want)
    End If
    Exit Sub
Fail:
    Errors_Log "Scatter_SyncAxisFormats" & vbTab & k & vbTab & Err.Number & vbTab & Err.Description
    On Error GoTo -1
    On Error Resume Next
    If started Then                                                             ' o eixo e as opcoes voltam ao que eram
        If linked Then ax.TickLabels.NumberFormatLinked = True Else ax.TickLabels.NumberFormatLocal = cur
        Registry_SetOption cht, key, OptEncode(stored)
        Registry_SetOption cht, srcKey, OptEncode(src)
    End If
End Sub

' Texto livre numa opcao do registro (que nao aceita ";" nem "="): %, ; e = viram %25, %3B e %3D.
Private Function OptEncode(s As String) As String
    OptEncode = Replace(Replace(Replace(s, "%", "%25"), ";", "%3B"), "=", "%3D")
End Function

Private Function OptDecode(s As String) As String
    OptDecode = Replace(Replace(Replace(s, "%3D", "="), "%3B", ";"), "%25", "%")
End Function

' ---------- escala dos eixos ----------

' Faixa "bonita" de um eixo de dispersao com os dados em [lo, hi]. A escala automatica do Excel comeca em zero quando a faixa passa de 1/6 do
' maximo e amontoa pontos como 20..50 num canto (visto na demonstracao). True (com outMin/outMax/outStep) quando os dados nao incluem o zero
' e a faixa e menor que SC_TIGHT_FRAC do maior valor absoluto: [lo - folga, hi + folga] (folga = padFrac da faixa; um valor so: padFrac do
' valor) arredondada para fora a um passo 1/2/2,5/5 x 10^k, o menor com no maximo SC_MAX_STEPS intervalos e >= minUnit (resolucao do
' formato: com 0% um passo de 0,5% repetiria "30%"). Nunca cruza o zero pelo arredondamento. False = escala do Excel. Pura.
Public Function Scatter_NiceRange(lo As Double, hi As Double, minUnit As Double, padFrac As Double, ByRef outMin As Double, ByRef outMax As Double, ByRef outStep As Double) As Boolean
    Dim big As Double, pad As Double, a As Double, b As Double, e As Double, s As Double, k As Long, lo0 As Double, hi0 As Double
    If hi < lo Then Exit Function
    If lo <= 0 And hi >= 0 Then Exit Function
    If Abs(lo) > Abs(hi) Then big = Abs(lo) Else big = Abs(hi)
    If hi - lo >= SC_TIGHT_FRAC * big Then Exit Function
    pad = padFrac * (hi - lo)
    If pad <= 0 Then pad = padFrac * big
    a = lo - pad: b = hi + pad
    e = 10 ^ Int(Log((b - a) / SC_MAX_STEPS) / Log(10))
    For k = 0 To 15
        s = Choose((k Mod 4) + 1, 1, 2, 2.5, 5) * e * 10 ^ (k \ 4)
        lo0 = Int(a / s + 0.000000001): hi0 = -Int(-b / s + 0.000000001)
        If s >= minUnit * (1 - 0.000000001) And hi0 - lo0 <= SC_MAX_STEPS Then Exit For
    Next
    outMin = lo0 * s: outMax = hi0 * s: outStep = s
    If lo > 0 And outMin < 0 Then outMin = 0
    If hi < 0 And outMax > 0 Then outMax = 0
    Scatter_NiceRange = True
End Function

' Resolucao de um codigo de formato (ingles, da celula de origem): 10^-casas (dividida por 100 com %) da primeira secao limpa
' (Style_CleanFormatSection); 0 para Geral, datas e
' horas (sem restricao de passo).
Private Function FormatUnit(code As String) As Double
    Dim i As Long, ch As String, clean As String, p As Long, d As Long
    If Len(code) = 0 Or code = "General" Then Exit Function
    clean = Style_CleanFormatSection(code)                      ' sem textos, colchetes ([Red] tem um d) e literais
    If clean Like "*[dDmMyYhHsS]*" Then Exit Function
    p = InStr(clean, ".")
    If p > 0 Then
        For i = p + 1 To Len(clean)
            ch = Mid$(clean, i, 1)
            If ch = "0" Or ch = "#" Or ch = "?" Then d = d + 1 Else Exit For
        Next
    End If
    FormatUnit = 10 ^ (-d)
    If InStr(clean, "%") > 0 Then FormatUnit = FormatUnit / 100
End Function

' Menor e maior valor PLOTADO no eixo k (1 = X, 2 = Y): os valores numericos das series de dados (os pontos fora do grupo ou invalidos sao
' NA() no bloco e chegam vazios). False sem nenhum.
Private Function DataRange(cht As Chart, k As Long, ByRef lo As Double, ByRef hi As Double) As Boolean
    Dim s As Series, v As Variant, i As Long, x As Double, j As Long
    For j = 1 To Scatter_SeriesCount(cht)
        Set s = Scatter_DataSeries(cht, j)
        If k = 1 Then v = s.XValues Else v = s.Values
        If IsArray(v) Then
            For i = LBound(v) To UBound(v)
                If Not IsError(v(i)) And Not IsEmpty(v(i)) And VarType(v(i)) <> vbString And VarType(v(i)) <> vbBoolean Then
                    x = CDbl(v(i))
                    If Not DataRange Then
                        lo = x: hi = x: DataRange = True
                    Else
                        If x < lo Then lo = x
                        If x > hi Then hi = x
                    End If
                End If
            Next
        End If
    Next
End Function

' Folga de cada lado (fracao da faixa dos dados) do eixo k nas bolhas: 5% mais o raio da MAIOR bolha (Point.Width: o diametro atual, lido sem
' desenhar o grafico; independe da escala dos eixos), para nenhuma bolha da borda sair da area de plotagem: com fr = raio / lado interno,
' folga p >= fr x (faixa + 2p) + 5% da faixa => p = (fr + 5%) / (1 - 2 fr). Dispersao: 5%. -1 se as bolhas nao cabem (fr >= 0,4).
Private Function PadFraction(cht As Chart, k As Long) As Double
    Dim s As Series, j As Long, i As Long, w As Double, d As Double, side As Double, fr As Double
    PadFraction = SC_AXIS_PAD
    If Registry_ChartType(cht) <> "bubble" Then Exit Function
    For j = 1 To Scatter_SeriesCount(cht)
        Set s = Scatter_DataSeries(cht, j)
        For i = 1 To s.Points.Count
            w = 0
            On Error Resume Next
            w = s.Points(i).Width
            On Error GoTo 0
            If w > d Then d = w
        Next
    Next
    If k = 1 Then side = cht.PlotArea.InsideWidth Else side = cht.PlotArea.InsideHeight
    If side <= 0 Then PadFraction = -1: Exit Function
    fr = (d / 2) / side
    If fr >= 0.4 Then PadFraction = -1: Exit Function
    PadFraction = (fr + SC_AXIS_PAD) / (1 - 2 * fr)
End Function

Private Function ScaleText(mn As Double, mx As Double, st As Double) As String
    ScaleText = Trim$(Str$(mn)) & "|" & Trim$(Str$(mx)) & "|" & Trim$(Str$(st))
End Function

Private Function SameNumber(a As Double, b As Double) As Boolean
    SameNumber = (Abs(a - b) <= 0.000000001 * (1 + Abs(a) + Abs(b)))
End Function

' O eixo esta exatamente na escala gravada pelo ThinQQ (minimo, maximo e passo fixos e iguais aos gravados)?
Private Function AxisIsOurs(ax As Axis, stored As String) As Boolean
    Dim p As Variant
    If Len(stored) = 0 Then Exit Function
    p = Split(stored, "|")
    If UBound(p) < 2 Then Exit Function
    If ax.MinimumScaleIsAuto Or ax.MaximumScaleIsAuto Or ax.MajorUnitIsAuto Then Exit Function
    AxisIsOurs = SameNumber(ax.MinimumScale, Val(p(0))) And SameNumber(ax.MaximumScale, Val(p(1))) And SameNumber(ax.MajorUnit, Val(p(2)))
End Function

' Escala do eixo k (1 = X: Axes(xlCategory); 2 = Y: Axes(xlValue)) pela faixa dos dados (Scatter_NiceRange), gravada na opcao xscale/yscale
' ("min|max|passo"; o registro nao aceita ";") para reconhecer a propria escala depois (como o "secmin" da combinacao): um eixo fixo que nao e o gravado foi ajustado
' pelo usuario e fica como esta (inclusive so o passo). Sem aperto (zero na faixa, faixa larga, sem dados) a escala gravada volta ao
' automatico do Excel. So escreve o que mudou. Transacional: uma falha no meio devolve o eixo e a opcao ao que eram. Nunca propaga erro.
Private Sub SyncAxisScale(cht As Chart, k As Long)
    Dim ax As Axis, key As String, stored As String, ours As Boolean, lo As Double, hi As Double, mn As Double, mx As Double, st As Double
    Dim want As Boolean, pf As Double, c As Range, unit As Double, started As Boolean, newText As String
    Dim pAutoMin As Boolean, pAutoMax As Boolean, pAutoUnit As Boolean, pMin As Double, pMax As Double, pUnit As Double
    On Error GoTo Fail
    key = IIf(k = 1, "xscale", "yscale")
    Set ax = cht.Axes(IIf(k = 1, xlCategory, xlValue), xlPrimary)
    stored = Registry_GetOption(cht, key, "")
    ours = AxisIsOurs(ax, stored)
    If Not ours Then
        If Not (ax.MinimumScaleIsAuto And ax.MaximumScaleIsAuto) Then Exit Sub        ' escala do usuario: fica
    End If
    If DataRange(cht, k, lo, hi) Then
        pf = PadFraction(cht, k)
        If pf >= 0 Then
            Set c = SourceNumberCell(cht, IIf(k = 1, SC_COL_X, SC_COL_Y))
            If Not c Is Nothing Then unit = FormatUnit(CStr(c.NumberFormat))
            want = Scatter_NiceRange(lo, hi, unit, pf, mn, mx, st)
        End If
    End If
    pAutoMin = ax.MinimumScaleIsAuto: pAutoMax = ax.MaximumScaleIsAuto: pAutoUnit = ax.MajorUnitIsAuto
    pMin = ax.MinimumScale: pMax = ax.MaximumScale: pUnit = ax.MajorUnit
    If want Then
        newText = ScaleText(mn, mx, st)
        If ours And stored = newText Then Exit Sub
        started = True
        If mn < ax.MaximumScale Then
            ax.MinimumScale = mn: ax.MaximumScale = mx
        Else
            ax.MaximumScale = mx: ax.MinimumScale = mn                      ' subindo: o minimo novo passaria do maximo atual
        End If
        ax.MajorUnit = st
        ScatterTestHook "scale", CStr(k)
        Registry_SetOption cht, key, newText
    ElseIf ours Then
        started = True
        ax.MinimumScaleIsAuto = True: ax.MaximumScaleIsAuto = True: ax.MajorUnitIsAuto = True
        Registry_SetOption cht, key, ""
    ElseIf Len(stored) > 0 Then
        Registry_SetOption cht, key, ""
    End If
    Exit Sub
Fail:
    Errors_Log "Scatter_SyncAxisScale" & vbTab & k & vbTab & Err.Number & vbTab & Err.Description
    On Error GoTo -1
    On Error Resume Next
    If started Then
        If pAutoMin Then ax.MinimumScaleIsAuto = True Else ax.MinimumScale = pMin
        If pAutoMax Then ax.MaximumScaleIsAuto = True Else ax.MaximumScale = pMax
        If pAutoMin Then ax.MinimumScaleIsAuto = True Else ax.MinimumScale = pMin      ' de novo: a ordem pode ter sido recusada
        If pAutoUnit Then ax.MajorUnitIsAuto = True Else ax.MajorUnit = pUnit
        Registry_SetOption cht, key, stored
    End If
End Sub

' Escala dos dois eixos (SyncAxisScale). Na insercao e em todo refresh. Nunca propaga erro.
Public Sub Scatter_SyncAxes(cht As Chart)
    On Error Resume Next
    SyncAxisScale cht, 1
    SyncAxisScale cht, 2
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
' As escritas rodam com os eventos desligados e um sinalizador de ocupado (cada escrita dispararia Calculate e o refresh reentraria aqui);
' os dois voltam ao estado anterior na unica saida depois de ligados.
Public Sub Scatter_SyncBubbleLabels(cht As Chart)
    Dim k As Long, i As Long, s As Series, fc As Long, has As Boolean, d As Double, txt As String, wantPos As Long, wantColor As Long, ok As Boolean
    Dim prevEv As Boolean
    On Error Resume Next
    Err.Clear                                        ' um erro pendente do chamador nao vale como falha de leitura aqui
    If mBubbleSyncBusy Then Exit Sub
    If Registry_ChartType(cht) <> "bubble" Then Exit Sub
    Err.Clear
    mBubbleSyncBusy = True
    mBubbleSyncRuns = mBubbleSyncRuns + 1
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
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
    Application.EnableEvents = prevEv
    mBubbleSyncBusy = False
End Sub

' Gancho de teste: passadas de Scatter_SyncBubbleLabels que fizeram o trabalho (a reentrada barrada nao conta); reset zera.
Public Function Scatter_BubbleSyncRunsForTest(Optional reset As Boolean = False) As Long
    Scatter_BubbleSyncRunsForTest = mBubbleSyncRuns
    If reset Then mBubbleSyncRuns = 0
End Function

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
' religar a legenda (Deco_SetLegend) as traz de volta, e o refresh as apaga de novo. Nunca propaga erro. Laco limitado ao numero de entradas
' lido no inicio (uma exclusao que nao reduz a contagem nao prende o refresh).
Public Sub Scatter_HideTrendEntries(cht As Chart)
    Dim ns As Long, cnt As Long, i As Long
    On Error Resume Next
    If Not cht.HasLegend Then Exit Sub
    ns = Scatter_SeriesCount(cht)
    cnt = cht.Legend.LegendEntries.Count
    If Err.Number <> 0 Then Exit Sub
    For i = cnt To ns + 1 Step -1
        cht.Legend.LegendEntries(i).Delete
        If Err.Number <> 0 Then Exit For
    Next
End Sub

' Quadrantes: os eixos cruzam nas medias (Axes(xlCategory).CrossesAt = media de X: linha vertical; Axes(xlValue).CrossesAt = media de Y:
' horizontal), tracejados; os rotulos continuam na borda (TickLabelPosition Low). Desligar volta ao cruzamento automatico e a linha solida.
' Transacional (como Scatter_SetTrend): aplica primeiro e SO ENTAO grava a opcao; uma falha no meio de ligar desfaz o eixo ja cruzado (os dois
' voltam ao automatico) e a opcao fica como estava; o erro segue para o chamador.
Public Sub Scatter_SetQuad(cht As Chart, turnOn As Boolean)
    Dim en As Long, es As String, ed As String
    On Error GoTo Fail
    If turnOn Then
        Scatter_ApplyQuadrants cht
    Else
        QuadAxisOff cht, 1
        QuadAxisOff cht, 2
    End If
    Registry_SetOption cht, "quad", IIf(turnOn, "1", "0")
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error GoTo -1                                 ' encerra o tratamento do erro: sem isso o Resume Next abaixo nao protegeria a limpeza
    On Error Resume Next
    If turnOn Then
        QuadAxisOff cht, 1                           ' cada eixo isolado: a falha de um nao impede desfazer o outro
        QuadAxisOff cht, 2
    End If
    On Error GoTo 0
    Err.Raise en, es, ed
End Sub

' Eixo k (1 = categorias/X, linha vertical; 2 = valores/Y, linha horizontal) de volta ao cruzamento automatico, linha solida.
Private Sub QuadAxisOff(cht As Chart, k As Long)
    With cht.Axes(IIf(k = 1, xlCategory, xlValue), xlPrimary)
        .Crosses = xlAxisCrossesAutomatic
        .Format.Line.DashStyle = msoLineSolid
    End With
End Sub

' Reaplica o cruzamento nas medias (meta E/F). So escreve quando mudou (chamada a cada refresh). Sem o ThinQQ as medias ficam no ultimo
' valor aplicado (como o intervalo do eixo do Gantt).
Public Sub Scatter_ApplyQuadrants(cht As Chart)
    Dim blk As Range, ax As Axis, v As Variant, k As Long
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    For k = 1 To 2
        ScatterTestHook "quad", CStr(k)
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
    Scatter_SyncAxisFormats cht                      ' o formato da coluna de origem pode ter mudado
    Scatter_SyncAxes cht                             ' a faixa dos dados (e o tamanho das bolhas) pode ter mudado
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
