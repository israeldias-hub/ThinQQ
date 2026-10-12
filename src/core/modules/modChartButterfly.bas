Attribute VB_Name = "modChartButterfly"
Option Explicit
' Borboleta (kind butterfly, familia butterfly, horizontal): barras EMPILHADAS com dois lados espelhados. Series na ordem tqq_gapL (-Gap,
' invisivel), <lado esquerdo> (-A), tqq_gapR (+Gap, invisivel), <lado direito> (+B): os negativos empilham para a esquerda e os positivos
' para a direita na ordem das series (medido), o que deixa um vao central de 2*Gap onde ficam os nomes das categorias (auxiliar XY
' tqq_cat em x = 0, rotulos ligados a coluna 1, posicao central). Os totais (A + B) sao a auxiliar XY tqq_total no fim da barra direita
' (x = Gap + B). A escala do eixo de valores e AUTOMATICA e simetrica: a auxiliar XY tqq_sym tem dois pontos em x = -(Gap + max) e
' +(Gap + max) (linha de meta do bloco) e os pontos XY do grupo secundario entram na escala automatica do eixo primario (medido na sonda):
' sem o ThinQQ a simetria continua viva. O lado esquerdo e desenhado com valores negativos, mas os rotulos usam o formato "absoluto"
' (secao negativa igual a positiva: Bfly_AbsFormat).
' Bloco (corpo): linha 1 = cabecalhos; n linhas: 1 Categoria | 2 Total (A+B) | 3 TotalPos (Gap+B) | 4 A | 5 B | 6 Gap | 7 NegGap | 8 NegA |
' 9 CatPos (0) | 10 Pos (n - i + 0,5). Meta (linha 1 do bloco): E = GapK (numero: fracao do maximo que vira o vao, gravado por Bfly_RefreshGap),
' F = SymL, G = SymR (formulas), H:I = n/2 (Y dos pontos de simetria). Gap = GapK * MAX(A, B) em todas as linhas.
Public Const BF_COL_CAT As Long = 1
Public Const BF_COL_TOTAL As Long = 2
Public Const BF_COL_TOTALPOS As Long = 3
Public Const BF_COL_A As Long = 4
Public Const BF_COL_B As Long = 5
Public Const BF_COL_GAP As Long = 6
Public Const BF_COL_NEGGAP As Long = 7
Public Const BF_COL_NEGA As Long = 8
Public Const BF_COL_CATPOS As Long = 9
Public Const BF_COL_POS As Long = 10
Private Const BF_COLS As Long = 10
Private Const BF_META_GAPK As Long = 5
Private Const BF_META_SYML As Long = 6
Private Const BF_META_SYMR As Long = 7
Private Const BF_META_SYMY As Long = 8
Private Const BF_GAPK_MIN As Double = 0.15       ' vao minimo: 15% do maior valor de cada lado
Private Const BF_GAPK_MAX As Double = 1          ' teto de GapK: com o eixo automatico o vao fica <= metade da area interna (Bfly_RefreshGap)
Private Const BF_GAP_MAXFRAC As Double = 0.5     ' teto do vao em pontos: metade da largura interna da area de plotagem
Private Const BF_LABEL_MARGIN As Double = 8      ' respiro (pt) alem da largura estimada do maior nome
Private Const BF_THIN_OPT As String = "thinlbl"  ' pares lado.ponto (1 = esquerdo, 2 = direito) cujo rotulo o ThinQQ ocultou (Bfly_SyncThinLabels)
Private Const BF_ERR_COLS As String = "A borboleta precisa de exatamente 2 colunas de valores (lado esquerdo e lado direito)."
Private mGapBusy As Boolean                      ' Bfly_RefreshGap grava uma celula: o evento Calculate do grafico reentra em Core_SafeRefresh

Private Sub Bfly_CheckData(data As clsChartData)
    If data.SeriesCount <> 2 Then Err.Raise vbObjectError + 6011, "Bfly_Insert", BF_ERR_COLS
End Sub

' Calculo manual: recalcula o bloco (Range.Calculate segue as dependencias dentro da faixa). Em automatico nao faz nada (o Excel ja
' recalculou ao gravar).
Private Sub Bfly_CalcIfManual(blk As Range)
    If Application.Calculation <> xlCalculationAutomatic Then blk.Calculate
End Sub

Public Function Bfly_BuildBlock(wb As Workbook, id As String, data As clsChartData) As Range
    Dim blk As Range, body As Range, n As Long, i As Long, r As Long, sA As String, sB As String, a As String, b As String, g As String
    Dim allAB As String, gapK As String, hdr As Variant, k As Long, sC As String
    Bfly_CheckData data
    n = data.CategoryCount
    Set blk = CalcSheet_AllocateBlock(wb, id, "butterfly", Mid$(ChartData_ExternalAddress(data.Source), 2), n + 1, BF_COLS)
    Set body = Block_Body(blk)
    hdr = Array("Categoria", "Total", "TotalPos", "A", "B", "Gap", "NegGap", "NegA", "CatPos", "Pos")
    For k = 0 To BF_COLS - 1
        body.Cells(1, k + 1).Value = hdr(k)
    Next
    allAB = body.Cells(2, BF_COL_A).Resize(n, 2).Address(True, True)       ' mesma planilha: sem nome de planilha
    gapK = blk.Cells(1, BF_META_GAPK).Address(True, True)
    For i = 1 To n
        r = i + 1
        sA = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 2)), 2)
        sB = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 3)), 2)
        a = body.Cells(r, BF_COL_A).Address(False, False): b = body.Cells(r, BF_COL_B).Address(False, False): g = body.Cells(r, BF_COL_GAP).Address(False, False)
        sC = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 1)), 2)
        body.Cells(r, BF_COL_CAT).Formula = "=IF(ISBLANK(" & sC & "),""""," & sC & ")"     ' categoria vazia: sem "0" no centro
        ' o rotulo ligado mostra o TEXTO da celula: o formato da categoria de origem (data, ano) vem junto (como no Mekko)
        If Not IsNull(data.Source.Cells(r, 1).NumberFormat) Then body.Cells(r, BF_COL_CAT).NumberFormat = data.Source.Cells(r, 1).NumberFormat
        body.Cells(r, BF_COL_A).Formula = "=IF(ISNUMBER(" & sA & "),MAX(0," & sA & "),0)"      ' negativo, texto e erro valem 0
        body.Cells(r, BF_COL_B).Formula = "=IF(ISNUMBER(" & sB & "),MAX(0," & sB & "),0)"
        body.Cells(r, BF_COL_TOTAL).Formula = "=" & a & "+" & b
        body.Cells(r, BF_COL_GAP).Formula = "=" & gapK & "*MAX(" & allAB & ")"
        body.Cells(r, BF_COL_TOTALPOS).Formula = "=" & g & "+" & b
        body.Cells(r, BF_COL_NEGGAP).Formula = "=-" & g
        body.Cells(r, BF_COL_NEGA).Formula = "=-" & a
        body.Cells(r, BF_COL_CATPOS).Formula = "=0"
        body.Cells(r, BF_COL_POS).Formula = "=" & (n - i) & ".5"           ' independe do separador decimal da localidade
    Next
    blk.Cells(1, BF_META_GAPK).Value = BF_GAPK_MIN
    blk.Cells(1, BF_META_SYML).Formula = "=-(" & body.Cells(2, BF_COL_GAP).Address(True, True) & "+MAX(" & allAB & "))"
    blk.Cells(1, BF_META_SYMR).Formula = "=" & body.Cells(2, BF_COL_GAP).Address(True, True) & "+MAX(" & allAB & ")"
    blk.Cells(1, BF_META_SYMY).Value = n / 2
    blk.Cells(1, BF_META_SYMY + 1).Value = n / 2
    Bfly_CalcIfManual blk                            ' calculo manual: as formulas acima foram gravadas antes das celulas que elas leem
    Set Bfly_BuildBlock = blk
End Function

Private Function AddBar(cht As Chart, name As String, valueCells As Range, cats As String) As Series
    Dim s As Series
    Set s = cht.SeriesCollection.NewSeries
    s.Name = name
    s.Values = Block_Ref(valueCells)
    s.XValues = cats
    Set AddBar = s
End Function

Public Function Bfly_Insert(ws As Worksheet, data As clsChartData) As ChartObject
    Dim id As String, co As ChartObject, cht As Chart, anchor As Range, blk As Range, body As Range, n As Long, cats As String
    Dim en As Long, es As String, ed As String
    Bfly_CheckData data                              ' recusa (6011) antes de alocar: nada a desfazer
    On Error GoTo Fail
    id = Registry_NewId()
    Set blk = Bfly_BuildBlock(ws.Parent, id, data)
    Util_TestHookAfterBlock
    Set body = Block_Body(blk)
    n = data.CategoryCount
    cats = ChartData_ExternalAddress(data.CategoryCells())
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, Chart_FreeTop(ws, anchor.Left, anchor.Top, 460, 280), 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    AddBar cht, "tqq_gapL", body.Cells(2, BF_COL_NEGGAP).Resize(n, 1), cats
    AddBar cht, ChartData_ExternalAddress(data.SeriesNameCell(1)), body.Cells(2, BF_COL_NEGA).Resize(n, 1), cats
    AddBar cht, "tqq_gapR", body.Cells(2, BF_COL_GAP).Resize(n, 1), cats
    AddBar cht, ChartData_ExternalAddress(data.SeriesNameCell(2)), body.Cells(2, BF_COL_B).Resize(n, 1), cats
    cht.ChartType = xlBarStacked
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlZero
    Style_ApplyBase cht                              ' categorias invertidas (1a no topo), eixo de valores oculto
    Util_TestHookAfterChart
    cht.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionNone     ' os nomes ficam no vao central (tqq_cat)
    cht.Axes(xlCategory).Format.Line.Visible = msoFalse                  ' a linha do eixo (em x = 0) cortaria os nomes ao meio
    cht.ChartGroups(1).GapWidth = 50
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionRight
    cht.Legend.Font.Size = 9
    Registry_SetOption cht, "numfmt", "int"
    Registry_SetOption cht, "palette", "padrao"
    Registry_SetOption cht, "gap", "50"
    Registry_SetOption cht, "n", CStr(n)
    Registry_SetOption cht, "legend", "right"
    Bfly_ApplyColors cht, "padrao"
    ' auxiliares: a simetria primeiro (uma linha: Chart_SetupSecondaryAxes a fixaria em 0..1), depois nomes e totais (n linhas)
    Chart_AddHelperSeries cht, "tqq_sym", blk.Cells(1, BF_META_SYML).Resize(1, 2), blk.Cells(1, BF_META_SYMY).Resize(1, 2)
    Bfly_SetCategoryLabels cht
    Bfly_SetLabels cht, True
    Bfly_SetTotals cht, True
    Chart_SetupSecondaryAxes cht, n
    Chart_HideHelperLegendEntries cht
    Bfly_RefreshGap cht
    Bfly_SyncThinLabels cht                          ' barras finas demais para o rotulo (na escala final, depois do vao)
    Set Bfly_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' k-esima serie NAO auxiliar: 1 = lado esquerdo (valores negativos), 2 = lado direito. Nothing se nao existe.
Public Function Bfly_SideSeries(cht As Chart, side As Long) As Series
    Dim s As Series, k As Long
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            k = k + 1
            If k = side Then Set Bfly_SideSeries = s: Exit Function
        End If
    Next
End Function

' Nomes das categorias no vao central: auxiliar XY tqq_cat (x = CatPos = 0, y = Pos) com o rotulo de cada ponto ligado a celula da
' categoria (DataLabel.Formula; nunca Labels_LinkPoints, que gravaria um formato numerico na celula e mostraria um ano como 2.022).
Private Sub Bfly_SetCategoryLabels(cht As Chart)
    Dim s As Series, body As Range, n As Long, i As Long
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    Set s = Chart_FindSeries(cht, "tqq_cat")
    If s Is Nothing Then Set s = Chart_AddHelperSeries(cht, "tqq_cat", body.Cells(2, BF_COL_CATPOS).Resize(n, 1), body.Cells(2, BF_COL_POS).Resize(n, 1))
    For i = 1 To n
        With s.Points(i)
            .HasDataLabel = True
            .DataLabel.Formula = Block_Ref(body.Cells(i + 1, BF_COL_CAT))
            .DataLabel.Position = xlLabelPositionCenter
            .DataLabel.Font.Size = 9
            .DataLabel.Font.Bold = False
            .DataLabel.Font.Color = RGB(64, 64, 64)
            .DataLabel.Format.TextFrame2.WordWrap = msoFalse        ' nome longo numa linha so: o Excel quebraria o rotulo mesmo com o vao folgado
        End With
    Next
End Sub

' Rotulos de valor nativos nas duas series de dados (InsideEnd, fonte por contraste); o lado esquerdo mostra o valor absoluto.
Public Sub Bfly_SetLabels(cht As Chart, turnOn As Boolean)
    Dim k As Long, s As Series
    For k = 1 To 2
        Set s = Bfly_SideSeries(cht, k)
        If Not s Is Nothing Then
            s.HasDataLabels = turnOn
            If turnOn Then
                With s.DataLabels
                    .ShowValue = True: .ShowSeriesName = False: .ShowCategoryName = False
                    .Position = xlLabelPositionInsideEnd
                    .Font.Size = 9
                    .Font.Color = Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB))
                End With
            End If
        End If
    Next
    Registry_SetOption cht, "labels", IIf(turnOn, "1", "0")
    Registry_SetOption cht, BF_THIN_OPT, ""          ' ligar refaz todos os rotulos e desligar apaga todos: nada fica oculto pelo ThinQQ
    If turnOn Then
        Bfly_ApplyNumberFormat cht
        Bfly_SyncThinLabels cht
    End If
End Sub

' Formato "absoluto": a secao negativa igual a positiva (o lado esquerdo e plotado em negativo). Pura.
Public Function Bfly_AbsFormat(fmt As String) As String
    Dim sec As String
    sec = Split(fmt & ";", ";")(0)
    Bfly_AbsFormat = sec & ";" & sec & ";"
End Function

' Formato numerico (opcao numfmt): lado esquerdo absoluto, lado direito normal, celulas do Total (Labels_ApplyCellFormat) e os rotulos
' ligados dos totais, ponto a ponto.
Public Sub Bfly_ApplyNumberFormat(cht As Chart)
    Dim fmt As String, s As Series, i As Long, k As Long
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    For k = 1 To 2
        Set s = Bfly_SideSeries(cht, k)
        If Not s Is Nothing Then
            If s.HasDataLabels Then
                Style_AssignNumberFormat s.DataLabels, SideFormat(cht, k)
            Else
                ' alguns pontos sem rotulo (barra fina oculta ou apagada pelo usuario; HasDataLabels le False): ponto a ponto, porque escrever
                ' em DataLabels da serie recriaria os rotulos que faltam
                For i = 1 To s.Points.Count
                    If s.Points(i).HasDataLabel Then Style_AssignNumberFormat s.Points(i).DataLabel, SideFormat(cht, k)
                Next
            End If
        End If
    Next
    Labels_ApplyCellFormat cht, fmt
    Set s = Chart_FindSeries(cht, "tqq_total")
    If Not s Is Nothing Then
        For i = 1 To s.Points.Count
            If s.Points(i).HasDataLabel Then Style_AssignNumberFormat s.Points(i).DataLabel, fmt
        Next
    End If
End Sub

' Totais: auxiliar XY tqq_total em (TotalPos = Gap + B, Pos) com o rotulo ligado ao Total (A + B), negrito, a direita; a folga a
' direita e a mesma das barras (Stacked_EnsureRoomAfterChange le a coluna 3 = TotalPos).
Public Sub Bfly_SetTotals(cht As Chart, turnOn As Boolean)
    Dim s As Series, body As Range, n As Long
    Set s = Chart_FindSeries(cht, "tqq_total")
    If Not turnOn Then
        If Not s Is Nothing Then s.Delete
        Registry_SetOption cht, "totals", "0"
        Exit Sub
    End If
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    If s Is Nothing Then Set s = Chart_AddHelperSeries(cht, "tqq_total", body.Cells(2, BF_COL_TOTALPOS).Resize(n, 1), body.Cells(2, BF_COL_POS).Resize(n, 1))
    Chart_HideHelperLegendEntries cht
    Labels_LinkPoints s, body.Cells(2, BF_COL_TOTAL).Resize(n, 1), Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int")), , , True
    Labels_RefreshPositions cht
    s.DataLabels.Font.Bold = True
    Registry_SetOption cht, "totals", "1"
    Stacked_EnsureRoomAfterChange cht
End Sub

' Cores: a paleta nas duas series de dados (Style_ApplySeriesColors pula as auxiliares), os vaos sempre invisiveis e a fonte dos rotulos
' por contraste com a cor nova.
Public Sub Bfly_ApplyColors(cht As Chart, paletteName As String)
    Dim s As Series, k As Long, i As Long, fc As Long
    Style_ApplySeriesColors cht, paletteName
    For Each s In cht.SeriesCollection
        If s.Name = "tqq_gapL" Or s.Name = "tqq_gapR" Then
            s.Format.Fill.Visible = msoFalse
            s.Format.Line.Visible = msoFalse
        End If
    Next
    For k = 1 To 2
        Set s = Bfly_SideSeries(cht, k)
        If Not s Is Nothing Then
            fc = Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB))
            If s.HasDataLabels Then
                s.DataLabels.Font.Color = fc
            Else
                For i = 1 To s.Points.Count                  ' alguns pontos sem rotulo: ponto a ponto (a serie recriaria os que faltam)
                    If s.Points(i).HasDataLabel Then s.Points(i).DataLabel.Font.Color = fc
                Next
            End If
        End If
    Next
End Sub

' Vao central: GapK (meta E) x maior valor. O vao em pontos e medido na escala REAL do eixo (Geom_ChartValueToX) e comparado com a largura
' estimada do maior nome (Labels_EstimateWidth, que superestima de proposito) + margem: cresce se nao cabe (ate 3 passadas) e encolhe se
' sobra mais de 60%. So grava quando muda mais de 2% (gravar recalcula e dispara o Calculate do grafico, que reentra aqui: guarda mGapBusy).
' Teto: o vao nunca passa de metade da area interna (BF_GAP_MAXFRAC); um nome mais largo que isso transborda a sua faixa (melhor que as barras
' sumirem). Com o eixo automatico o maximo do eixo e >= Gap + maior valor (a auxiliar tqq_sym), entao o vao em pontos e <= largura x k / (k + 1):
' com k = BF_GAPK_MAX = 1 ele ja e no maximo a metade. Sem o teto o ramo "cresce" dispararia em TODO refresh (o vao nunca alcanca a largura
' interna) e GapK se multiplicaria a cada calculo. Nome que pede o teto com eixo automatico: k vai direto a BF_GAPK_MAX (estavel).
' Sem o ThinQQ, GapK fica fixo e o vao continua proporcional ao maximo. Nunca propaga erro.
Public Sub Bfly_RefreshGap(cht As Chart)
    Dim blk As Range, body As Range, n As Long, i As Long, w As Double, lw As Double, need As Double, k As Double, g As Double, gapPt As Double, pass As Long
    Dim cap As Double, tgt As Double, capped As Boolean, wrote As Boolean
    If mGapBusy Then Exit Sub
    mGapBusy = True
    On Error GoTo Done
    If Kinds_Family(Registry_ChartType(cht)) <> "butterfly" Then GoTo Done
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    For i = 1 To n
        w = Labels_EstimateWidth(Util_CellLabel(body.Cells(i + 1, BF_COL_CAT)))
        If w > lw Then lw = w
    Next
    need = lw + BF_LABEL_MARGIN
    cap = BF_GAP_MAXFRAC * cht.PlotArea.InsideWidth
    If need > cap Then need = cap: capped = True
    k = Util_ToDouble(blk.Cells(1, BF_META_GAPK).Value)
    If k < BF_GAPK_MIN Then k = BF_GAPK_MIN
    For pass = 1 To 3
        g = Util_ToDouble(body.Cells(2, BF_COL_GAP).Value)
        If g <= 0 Then Exit For
        gapPt = Geom_ChartValueToX(cht, g) - Geom_ChartValueToX(cht, -g)
        If gapPt <= 0 Then Exit For
        If gapPt < need - 0.5 Then
            tgt = need * 1.05
            If tgt > cap Then tgt = cap
        ElseIf (gapPt > need * 1.6 Or gapPt > cap + 0.5) And k > BF_GAPK_MIN Then
            tgt = need * 1.2
            If tgt > cap Then tgt = cap
        Else
            Exit For
        End If
        k = k * tgt / gapPt
        If capped And tgt = cap And cht.Axes(xlValue, xlPrimary).MaximumScaleIsAuto Then k = BF_GAPK_MAX
        If k > BF_GAPK_MAX Then k = BF_GAPK_MAX
        If k < BF_GAPK_MIN Then k = BF_GAPK_MIN
        If Abs(k - Util_ToDouble(blk.Cells(1, BF_META_GAPK).Value)) <= 0.02 * k Then Exit For
        blk.Cells(1, BF_META_GAPK).Value = k
        wrote = True
        Bfly_CalcIfManual blk                        ' calculo manual: sem isto a passada seguinte releria o vao antigo e multiplicaria GapK
    Next
    ' o vao novo empurra os totais (x = Gap + B): a folga reservada por Bfly_SetTotals valia para o vao antigo. So reduz a area de
    ' plotagem e nao chama o refresh (sem reentrada; mGapBusy continua ligado).
    If wrote Then Stacked_EnsureRoomAfterChange cht
Done:
    If Err.Number <> 0 Then Errors_Log "Bfly_RefreshGap" & vbTab & Err.Number & vbTab & Err.Description
    mGapBusy = False
End Sub

' O que na borboleta depende dos DADOS ou do tamanho e muda depois de inserir (Core_SafeRefresh): o vao central. Nunca propaga erro.
Public Sub Bfly_RefreshAfterDataChange(cht As Chart)
    On Error Resume Next
    If Kinds_Family(Registry_ChartType(cht)) <> "butterfly" Then Exit Sub
    Bfly_RefreshGap cht
    Bfly_SyncThinLabels cht                          ' depois do vao: a escala nova decide quais barras ficam finas demais
End Sub

' Formato (codigo em ingles) dos rotulos do lado k: o esquerdo e plotado em negativo e mostra o valor absoluto (Bfly_AbsFormat).
Private Function SideFormat(cht As Chart, k As Long) As String
    Dim fmt As String
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    If k = 1 Then SideFormat = Bfly_AbsFormat(fmt) Else SideFormat = fmt
End Function

' Barras finas: o rotulo InsideEnd de uma barra mais estreita que o texto nao cabe nela; o Excel o poe fora da barra (o esquerdo sobre o fundo,
' com a fonte branca de contraste: some) ou, no lado direito, dentro do vao central ao lado do nome (medido). Esse rotulo fica OCULTO enquanto
' nao cabe e volta (com o formato e a fonte de Bfly_SetLabels) quando cabe. Largura da barra pela escala atual (valor x largura interna / faixa do
' eixo); largura do texto estimada (Labels_EstimateWidth do valor no formato do lado). So mexe no que o ThinQQ escondeu: cada par lado.ponto
' oculto aqui fica na opcao "thinlbl" (a mesma da cascata: "Alterar dados" a limpa) e sai dela ao voltar; um rotulo que falta e nao esta na lista
' foi apagado pelo usuario e fica apagado. Valor zero (rotulo vazio) fica como esta. Cada ponto e isolado; nunca propaga erro.
Public Sub Bfly_SyncThinLabels(cht As Chart)
    Dim body As Range, n As Long, k As Long, i As Long, s As Series, ax As Axis, ppu As Double, v As Double, txtW As Double
    Dim list0 As String, list As String, key As String, has As Boolean, mine As Boolean, fmt As String, sec As String, col As Long
    On Error GoTo Done
    If Kinds_Family(Registry_ChartType(cht)) <> "butterfly" Then Exit Sub
    If Registry_GetOption(cht, "labels", "1") <> "1" Then Exit Sub
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    Set ax = cht.Axes(xlValue, xlPrimary)
    If ax.MaximumScale <= ax.MinimumScale Then Exit Sub
    ppu = cht.PlotArea.InsideWidth / (ax.MaximumScale - ax.MinimumScale)
    list0 = Registry_GetOption(cht, BF_THIN_OPT, "")
    list = "," & list0 & ","
    For k = 1 To 2
        Set s = Bfly_SideSeries(cht, k)
        If Not s Is Nothing Then
            fmt = SideFormat(cht, k)
            sec = Split(fmt, ";")(0)
            If k = 1 Then col = BF_COL_A Else col = BF_COL_B
            For i = 1 To n
                key = k & "." & i
                mine = (InStr(list, "," & key & ",") > 0)
                v = Util_ToDouble(body.Cells(i + 1, col).Value)
                On Error Resume Next
                has = False
                has = s.Points(i).HasDataLabel
                If Err.Number = 0 And v > 0 Then
                    txtW = Labels_EstimateWidth(Format$(v, sec))
                    If has And v * ppu < txtW Then
                        s.Points(i).HasDataLabel = False
                        If Err.Number = 0 And Not mine Then list = list & key & ","
                    ElseIf (Not has) And mine And v * ppu >= txtW Then
                        s.Points(i).HasDataLabel = True
                        FormatSideLabel s, i, fmt
                        If Err.Number = 0 Then list = Replace(list, "," & key & ",", ",")
                    ElseIf has And mine Then
                        list = Replace(list, "," & key & ",", ",")      ' o rotulo voltou por outro caminho: deixa de ser "oculto pelo ThinQQ"
                    End If
                End If
                If Err.Number <> 0 Then Errors_Log "Bfly_SyncThinLabels" & vbTab & key & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
                On Error GoTo Done
            Next
        End If
    Next
Done:
    If Err.Number <> 0 Then Errors_Log "Bfly_SyncThinLabels" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    On Error Resume Next
    Do While InStr(list, ",,") > 0
        list = Replace(list, ",,", ",")
    Loop
    If Left$(list, 1) = "," Then list = Mid$(list, 2)
    If Right$(list, 1) = "," Then list = Left$(list, Len(list) - 1)
    If list <> list0 Then Registry_SetOption cht, BF_THIN_OPT, list
End Sub

' Rotulo de um ponto de lado com o formato de Bfly_SetLabels (ao voltar depois de oculto por Bfly_SyncThinLabels). Propaga erros.
Private Sub FormatSideLabel(s As Series, i As Long, fmt As String)
    With s.Points(i).DataLabel
        .ShowValue = True: .ShowSeriesName = False: .ShowCategoryName = False
        .Position = xlLabelPositionInsideEnd
        .Font.Size = 9
        .Font.Color = Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB))
    End With
    Style_AssignNumberFormat s.Points(i).DataLabel, fmt
End Sub
