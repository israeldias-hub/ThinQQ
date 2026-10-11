Attribute VB_Name = "modChartGantt"
Option Explicit
' Gantt (kind gantt; design da Fase 2, secoes 3.3 e 4.6): barras empilhadas HORIZONTAIS com a serie inicial invisivel (tqq_start = BarStart)
' e a duracao visivel (tqq_dur = Dur, fim inclusivo: a barra cobre o dia do fim inteiro e termina em Fim + 1) e auxiliares XY no grupo
' secundario (Chart_AddHelperSeries): marcos (tqq_ms, losango no dia), rotulos a direita das barras (tqq_lbl: a data do fim ou a 4a coluna
' da tabela) e a escala de datas no topo (tqq_scale: um ponto por marca em y = n, rotulo acima e barra de erro tracejada descendo ate 0 =
' separador vertical). O eixo de valores das barras e numerico (datas = numeros de serie) e as auxiliares XY usam o MESMO eixo horizontal
' (medido: o Excel nao cria eixo X secundario): a escala (semanas, meses, trimestres ou anos) e calculada por formula no bloco e
' Gantt_RefreshScale aplica minimo/maximo ao eixo (no Insert e em cada Core_SafeRefresh). Sem o ThinQQ, barras, marcos e rotulos continuam
' ligados as celulas; so o intervalo do eixo fica como estava.
' Bloco (corpo): linha 1 = cabecalhos; MAX(n, 40) linhas; 19 colunas. Atividades (linhas 2..n+1): 1 Atividade | 2 Inicio | 3 Fim | 4 IsMs |
' 5 BarStart | 6 Dur | 7 MsX | 8 Pos | 9 LblX | 10 LblTxt (11 vazia: separa atividades e escala). Escala: 12 Unit | 13 AxisMin | 14 AxisMax
' (so na linha 2) | 15 Tick | 16 TickY | 17 TickLbl (linhas 2..41: GT_SCALE_SLOTS marcas) | 18 DMin | 19 DMax (so na linha 2).
' Rotulos ligados ponto a ponto (DataLabel.Formula = Block_Ref(celula)), nunca por Labels_LinkPoints: ele grava na celula o NumberFormatLocal
' de um formato numerico e o rotulo perderia o formato de data (dd/mm, mmm/yy) da celula.
' Pontos NA() das auxiliares (marcas que sobram das 40, atividade sem data): com DisplayBlanksAs = xlZero (o padrao dos outros tipos) o rotulo de
' um ponto NA() aparece no canto do zero com o texto da celula (medido na Tarefa 4 e aqui); com xlNotPlotted, que o Gantt usa, o Excel nao o
' desenha (medido: posicao indefinida, nada no PNG). Alem disso GanttSyncLabels so liga o rotulo dos pontos com posicao valida (e apaga os demais)
' e as celulas de texto (LblTxt, TickLbl) devolvem "" sem posicao.
Public Const GT_COL_ACT As Long = 1
Public Const GT_COL_START As Long = 2
Public Const GT_COL_END As Long = 3
Public Const GT_COL_ISMS As Long = 4
Public Const GT_COL_BARSTART As Long = 5
Public Const GT_COL_DUR As Long = 6
Public Const GT_COL_MSX As Long = 7
Public Const GT_COL_POS As Long = 8
Public Const GT_COL_LBLX As Long = 9
Public Const GT_COL_LBLTXT As Long = 10
Public Const GT_COL_UNIT As Long = 12
Public Const GT_COL_AXMIN As Long = 13
Public Const GT_COL_AXMAX As Long = 14
Public Const GT_COL_TICK As Long = 15
Public Const GT_COL_TICKY As Long = 16
Public Const GT_COL_TICKLBL As Long = 17
Public Const GT_COL_DMIN As Long = 18
Public Const GT_COL_DMAX As Long = 19
Public Const GT_SCALE_SLOTS As Long = 40
Private Const GT_COLS As Long = 19
Private Const GT_TOP_ROOM As Double = 24        ' folga minima (pt) acima da area interna: rotulos da escala (acima do ponto em y = n)
Private Const GT_LABEL_OFFSET As Double = 6     ' o rotulo "Right" de um ponto XY comeca ~5,5 pt depois dele (medido)
Private Const GT_LABEL_MARGIN As Double = 4     ' respiro entre o fim do rotulo da barra e a borda do grafico
Private Const GT_SCALE_GAP As Double = 3        ' espaco minimo entre dois rotulos vizinhos da escala
Private Const GT_LABEL_COLOR As Long = 5855577  ' RGB(89, 89, 89): rotulos das barras e da escala
Private Const GT_GRID_COLOR As Long = 14277081  ' RGB(217, 217, 217): separadores verticais (barras de erro da escala)

' ---------------------------------------------------------------------------------------------------------------------
' Bloco
' ---------------------------------------------------------------------------------------------------------------------

' Recusa (6010, aviso de Core_Insert) antes de alocar qualquer coisa: menos de 3 colunas, ou nenhuma data (numero) na coluna de inicio.
Private Sub Gantt_CheckData(data As clsChartData)
    If data.SeriesCount < 2 Then
        Err.Raise vbObjectError + 6010, "Gantt_Insert", "Gantt: a tabela precisa de 3 colunas: atividade, início e fim (datas). Deixe o fim vazio para um marco."
    End If
    If Application.WorksheetFunction.Count(data.SeriesValueCells(1)) = 0 Then
        Err.Raise vbObjectError + 6010, "Gantt_Insert", "Gantt: a coluna de início não tem datas válidas (use células de data, não texto)."
    End If
End Sub

Public Function Gantt_BuildBlock(wb As Workbook, id As String, data As clsChartData) As Range
    Dim blk As Range, body As Range, n As Long, i As Long, k As Long, r As Long, f() As Variant, hdr As Variant
    Dim s2 As String, s3 As String, s4 As String, st As String, en As String, ms As String, lx As String
    Dim dmin As String, dmax As String, u As String, am As String, ax As String, tk As String, t As String, span As String
    Gantt_CheckData data
    n = data.CategoryCount
    Set blk = CalcSheet_AllocateBlock(wb, id, "gantt", Mid$(ChartData_ExternalAddress(data.Source), 2), IIf(n > GT_SCALE_SLOTS, n, GT_SCALE_SLOTS) + 1, GT_COLS)
    Set body = Block_Body(blk)
    hdr = Array("Atividade", "Inicio", "Fim", "IsMs", "BarStart", "Dur", "MsX", "Pos", "LblX", "LblTxt", "", "Unit", "AxisMin", "AxisMax", "Tick", "TickY", "TickLbl", "DMin", "DMax")
    For k = 0 To GT_COLS - 1
        body.Cells(1, k + 1).Value = hdr(k)
    Next
    ' atividades (linhas 2..n+1): as formulas vao de uma vez (uma escrita para o retangulo em vez de uma por celula)
    ReDim f(1 To n, 1 To GT_COL_LBLTXT)
    For i = 1 To n
        r = i + 1
        s2 = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 2)), 2)
        s3 = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 3)), 2)
        st = body.Cells(r, GT_COL_START).Address(False, False)
        en = body.Cells(r, GT_COL_END).Address(False, False)
        ms = body.Cells(r, GT_COL_ISMS).Address(False, False)
        lx = body.Cells(r, GT_COL_LBLX).Address(False, False)
        f(i, GT_COL_ACT) = ChartData_ExternalAddress(data.Source.Cells(r, 1))
        f(i, GT_COL_START) = "=IF(ISNUMBER(" & s2 & ")," & s2 & ",NA())"
        f(i, GT_COL_END) = "=IF(ISNUMBER(" & s3 & ")," & s3 & ",NA())"
        f(i, GT_COL_ISMS) = "=AND(ISNUMBER(" & s2 & "),NOT(ISNUMBER(" & s3 & ")))"         ' fim vazio (ou texto) = marco
        f(i, GT_COL_BARSTART) = "=IF(" & ms & ",NA()," & st & ")"
        f(i, GT_COL_DUR) = "=IF(" & ms & ",NA(),MAX(0," & en & "-" & st & "+1))"          ' fim inclusivo
        f(i, GT_COL_MSX) = "=IF(" & ms & "," & st & ",NA())"
        f(i, GT_COL_POS) = "=" & (n - i) & ".5"                                          ' independe do separador decimal da localidade
        ' depois do fim visual da barra (Fim + 1); marco: no dia. Sem inicio valido nao ha barra nem marco: sem rotulo solto (NA())
        f(i, GT_COL_LBLX) = "=IF(" & ms & "," & st & ",IF(ISNUMBER(" & st & ")," & en & "+1,NA()))"
        If data.SeriesCount >= 3 Then
            s4 = Mid$(ChartData_ExternalAddress(data.Source.Cells(r, 4)), 2)
            f(i, GT_COL_LBLTXT) = "=IF(ISNUMBER(" & lx & ")," & s4 & "&"""","""")"
        Else
            f(i, GT_COL_LBLTXT) = "=IF(ISNUMBER(" & lx & "),IF(" & ms & "," & st & "," & en & "),"""")"     ' o ultimo dia da atividade
        End If
    Next
    body.Cells(2, 1).Resize(n, GT_COL_LBLTXT).Formula = f
    If data.SeriesCount < 3 Then body.Cells(2, GT_COL_LBLTXT).Resize(n, 1).NumberFormat = "dd/mm"
    ' escala (linha 2): DMin/DMax e, a partir deles, a unidade e o intervalo do eixo alinhados a unidade
    body.Cells(2, GT_COL_DMIN).Formula = "=AGGREGATE(5,6," & body.Cells(2, GT_COL_START).Resize(n, 1).Address(True, True) & ")"
    body.Cells(2, GT_COL_DMAX).Formula = "=AGGREGATE(4,6," & body.Cells(2, GT_COL_START).Resize(n, 1).Address(True, True) & "," & body.Cells(2, GT_COL_END).Resize(n, 1).Address(True, True) & ")"
    dmin = body.Cells(2, GT_COL_DMIN).Address(True, True)
    dmax = body.Cells(2, GT_COL_DMAX).Address(True, True)
    u = body.Cells(2, GT_COL_UNIT).Address(True, True)
    am = body.Cells(2, GT_COL_AXMIN).Address(True, True)
    ax = body.Cells(2, GT_COL_AXMAX).Address(True, True)
    span = dmax & "-" & dmin
    body.Cells(2, GT_COL_UNIT).Formula = "=IF(" & span & "<=98,""w"",IF(" & span & "<=1100,""m"",IF(" & span & "<=3700,""q"",""y"")))"
    t = "=IF(" & u & "=""w""," & dmin & "-WEEKDAY(" & dmin & ",3),IF(" & u & "=""m"",DATE(YEAR(" & dmin & "),MONTH(" & dmin & "),1),"
    t = t & "IF(" & u & "=""q"",DATE(YEAR(" & dmin & "),3*INT((MONTH(" & dmin & ")-1)/3)+1,1),DATE(YEAR(" & dmin & "),1,1))))"
    body.Cells(2, GT_COL_AXMIN).Formula = t
    t = "=IF(" & u & "=""w""," & dmax & "-WEEKDAY(" & dmax & ",3)+7,IF(" & u & "=""m"",EDATE(DATE(YEAR(" & dmax & "),MONTH(" & dmax & "),1),1),"
    t = t & "IF(" & u & "=""q"",EDATE(DATE(YEAR(" & dmax & "),3*INT((MONTH(" & dmax & ")-1)/3)+1,1),3),DATE(YEAR(" & dmax & ")+1,1,1))))"
    body.Cells(2, GT_COL_AXMAX).Formula = t
    ' marcas k = 0..39 (linha 2 + k): Tick so antes do fim do eixo; TickY = n (topo do eixo Y secundario); TickLbl = a data (formato pela
    ' unidade, Gantt_RefreshScale), "T1/2026" no trimestre, o ano no ano; "" sem marca
    ReDim f(1 To GT_SCALE_SLOTS, 1 To 3)
    For k = 0 To GT_SCALE_SLOTS - 1
        r = k + 2
        tk = body.Cells(r, GT_COL_TICK).Address(False, False)
        t = "IF(" & u & "=""w""," & am & "+7*" & k & ",EDATE(" & am & "," & k & "*IF(" & u & "=""m"",1,IF(" & u & "=""q"",3,12))))"
        f(k + 1, 1) = "=IF(" & t & "<" & ax & "," & t & ",NA())"
        f(k + 1, 2) = "=IF(ISNUMBER(" & tk & ")," & n & ",NA())"
        f(k + 1, 3) = "=IF(ISNUMBER(" & tk & "),IF(" & u & "=""q"",""T""&(INT((MONTH(" & tk & ")-1)/3)+1)&""/""&YEAR(" & tk & "),IF(" & u & "=""y"",YEAR(" & tk & ")," & tk & ")),"""")"
    Next
    body.Cells(2, GT_COL_TICK).Resize(GT_SCALE_SLOTS, 3).Formula = f
    Set Gantt_BuildBlock = blk
End Function

' ---------------------------------------------------------------------------------------------------------------------
' Insercao
' ---------------------------------------------------------------------------------------------------------------------

Public Function Gantt_Insert(ws As Worksheet, data As clsChartData) As ChartObject
    Dim id As String, co As ChartObject, cht As Chart, s As Series, anchor As Range, blk As Range, body As Range, n As Long
    Dim en As Long, es As String, ed As String
    Gantt_CheckData data                             ' recusa (6010) antes de alocar: nada a desfazer
    On Error GoTo Fail
    id = Registry_NewId()
    Set blk = Gantt_BuildBlock(ws.Parent, id, data)
    Util_TestHookAfterBlock
    Set body = Block_Body(blk)
    n = data.CategoryCount
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, Chart_FreeTop(ws, anchor.Left, anchor.Top, 460, 280), 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    Set s = cht.SeriesCollection.NewSeries
    s.Name = "tqq_start"
    s.Values = Block_Ref(body.Cells(2, GT_COL_BARSTART).Resize(n, 1))
    s.XValues = ChartData_ExternalAddress(data.CategoryCells())
    Set s = cht.SeriesCollection.NewSeries
    s.Name = "tqq_dur"
    s.Values = Block_Ref(body.Cells(2, GT_COL_DUR).Resize(n, 1))
    s.XValues = ChartData_ExternalAddress(data.CategoryCells())
    cht.ChartType = xlBarStacked
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlNotPlotted               ' ponto NA() das auxiliares sem rotulo desenhado (ver o cabecalho)
    Style_ApplyBase cht                              ' categorias invertidas (1a atividade no topo), eixo de valores oculto
    Util_TestHookAfterChart
    cht.HasLegend = False
    cht.ChartGroups(1).GapWidth = 50
    With Chart_FindSeries(cht, "tqq_start").Format
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
    Registry_SetOption cht, "palette", "padrao"
    Registry_SetOption cht, "gap", "50"
    Registry_SetOption cht, "n", CStr(n)
    Set s = Chart_AddHelperSeries(cht, "tqq_ms", body.Cells(2, GT_COL_MSX).Resize(n, 1), body.Cells(2, GT_COL_POS).Resize(n, 1))
    s.MarkerStyle = xlMarkerStyleDiamond
    s.MarkerSize = 9
    Gantt_ApplyColors cht, "padrao"
    Gantt_SetLabels cht, True
    Gantt_AddScale cht, body, n
    Gantt_RefreshScale cht                           ' eixo, formato das marcas, folgas e rotulos da escala
    Set Gantt_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' Escala de datas: a auxiliar XY tqq_scale com X = Tick e Y = TickY (n: o topo do eixo Y secundario 0..n) nas GT_SCALE_SLOTS linhas e uma barra
' de erro de -n (do topo ate o zero) cinza-clara tracejada sem extremidade: o separador vertical de cada marca. Os rotulos (acima, ligados a
' TickLbl) sao ligados por Gantt_RefreshScale, depois de as celulas TickLbl receberem o formato da unidade.
Private Sub Gantt_AddScale(cht As Chart, body As Range, n As Long)
    Dim s As Series
    Set s = Chart_AddHelperSeries(cht, "tqq_scale", body.Cells(2, GT_COL_TICK).Resize(GT_SCALE_SLOTS, 1), body.Cells(2, GT_COL_TICKY).Resize(GT_SCALE_SLOTS, 1))
    s.ErrorBar Direction:=xlY, Include:=xlErrorBarIncludeMinusValues, Type:=xlErrorBarTypeFixedValue, Amount:=n
    With s.ErrorBars
        .EndStyle = xlNoCap
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = GT_GRID_COLOR
        .Format.Line.Weight = 0.75
        .Format.Line.DashStyle = msoLineDash
    End With
End Sub

' Liga o rotulo de cada ponto cuja posicao (posCells) nao e erro e cuja ordem cai no passo (ponto 1, 1 + passo, ...) a celula de txtCells
' (posicao, fonte de 'size' pt cinza) e apaga o rotulo dos demais. Sem 'force', so escreve onde o estado mudou (ponto que passou a ter ou deixou
' de ter marca, ou passo novo): e o que Gantt_RefreshScale chama a cada refresh. Cada ponto e isolado: um erro e registrado e os demais continuam.
Private Sub GanttSyncLabels(s As Series, posCells As Range, txtCells As Range, pos As Long, size As Double, force As Boolean, Optional stepK As Long = 1)
    Dim r As Long, m As Long, want As Boolean, has As Boolean
    m = posCells.Rows.Count
    If s.Points.Count < m Then m = s.Points.Count
    If stepK < 1 Then stepK = 1
    For r = 1 To m
        want = Not IsError(posCells.Cells(r, 1).Value)
        If want Then want = ((r - 1) Mod stepK = 0)
        On Error Resume Next
        has = s.Points(r).HasDataLabel
        If want Then
            If force Or Not has Then
                With s.Points(r)
                    .HasDataLabel = True
                    .DataLabel.Formula = Block_Ref(txtCells.Cells(r, 1))
                    .DataLabel.Position = pos
                    .DataLabel.Font.Size = size
                    .DataLabel.Font.Bold = False
                    .DataLabel.Font.Color = GT_LABEL_COLOR
                End With
            End If
        ElseIf has Or force Then
            s.Points(r).HasDataLabel = False
        End If
        If Err.Number <> 0 Then Errors_Log "GanttSyncLabels" & vbTab & s.Name & " ponto " & r & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
        On Error GoTo 0
    Next
End Sub

' Passo dos rotulos da escala: com marcas mais estreitas que o rotulo (muitos meses ou trimestres, grafico pequeno) os rotulos se sobreporiam;
' rotula so uma marca a cada 'passo' (os separadores ficam todos). Fatia de uma marca = largura interna * dias da unidade / dias do eixo;
' largura do rotulo de 8 pt medida no Excel 2019: "02/03" 25,3; "nov/26" 29,2 (o mais largo dos meses medidos); "T1/2026" 33,3; "2026" 22,2.
Private Function ScaleLabelStep(cht As Chart, u As String, mn As Double, mx As Double) As Long
    Dim days As Double, w As Double, slot As Double
    ScaleLabelStep = 1
    Select Case u
        Case "w": days = 7: w = 25.5
        Case "m": days = 30.44: w = 30
        Case "q": days = 91.31: w = 33.5
        Case Else: days = 365.25: w = 22.5
    End Select
    If mx <= mn Then Exit Function
    slot = cht.PlotArea.InsideWidth * days / (mx - mn)
    If slot <= 0 Then Exit Function
    ScaleLabelStep = -Int(-(w + GT_SCALE_GAP) / slot)
    If ScaleLabelStep < 1 Then ScaleLabelStep = 1
End Function

' ---------------------------------------------------------------------------------------------------------------------
' Rotulos, cores, escala e folga
' ---------------------------------------------------------------------------------------------------------------------

' Numero de atividades (opcao n gravada no Insert; sem ela, os pontos da primeira serie).
Private Function GanttCount(cht As Chart) As Long
    GanttCount = CLng(Val(Registry_GetOption(cht, "n", "")))
    If GanttCount <= 0 Then GanttCount = Geom_ChartCategoryCount(cht)
End Function

' Rotulos a direita das barras: a auxiliar tqq_lbl (X = LblX: Fim + 1, ou o dia do marco; Y = Pos) com o rotulo de cada ponto ligado a LblTxt
' (a data do fim/marco em dd/mm, ou o texto da 4a coluna), 9 pt cinza; a folga a direita para o rotulo mais a direita (Gantt_EnsureRoom).
' Desligar apaga a auxiliar.
Public Sub Gantt_SetLabels(cht As Chart, turnOn As Boolean)
    Dim s As Series, body As Range, n As Long
    Set s = Chart_FindSeries(cht, "tqq_lbl")
    If Not turnOn Then
        If Not s Is Nothing Then s.Delete
        Registry_SetOption cht, "labels", "0"
        Exit Sub
    End If
    Set body = Block_Body(Registry_Block(cht))
    n = GanttCount(cht)
    If s Is Nothing Then Set s = Chart_AddHelperSeries(cht, "tqq_lbl", body.Cells(2, GT_COL_LBLX).Resize(n, 1), body.Cells(2, GT_COL_POS).Resize(n, 1))
    GanttSyncLabels s, body.Cells(2, GT_COL_LBLX).Resize(n, 1), body.Cells(2, GT_COL_LBLTXT).Resize(n, 1), xlLabelPositionRight, 9, True
    Registry_SetOption cht, "labels", "1"
    ' com o eixo ainda automatico (no Insert, antes de Gantt_RefreshScale) as posicoes nao valem: a folga fica para o RefreshScale
    If Not cht.Axes(xlValue, xlPrimary).MinimumScaleIsAuto Then Gantt_EnsureRoom cht
End Sub

' Cores do Gantt: barras (tqq_dur) = pal(1), marcos (tqq_ms: preenchimento e borda do losango) = pal(0). ApplyOption "palette" chama esta.
' O marcador da auxiliar XY nao aceita a cor pelo Format.Fill (medido: lido de volta 0, losango preto): vai por MarkerBackground/ForegroundColor,
' com a cor do tema resolvida em RGB na paleta "tema".
Public Sub Gantt_ApplyColors(cht As Chart, paletteName As String)
    Dim pal As Variant, s As Series, c As Long
    pal = Style_Palette(paletteName)
    Set s = Chart_FindSeries(cht, "tqq_dur")
    If Not s Is Nothing Then
        s.Format.Fill.Visible = msoTrue
        Style_SetColor s.Format.Fill.ForeColor, pal(1)
        s.Format.Line.Visible = msoFalse
    End If
    Set s = Chart_FindSeries(cht, "tqq_ms")
    If Not s Is Nothing Then
        c = PaletteRGB(cht, pal(0))
        s.MarkerBackgroundColor = c
        s.MarkerForegroundColor = c
    End If
End Sub

' RGB de uma cor da paleta (Style_Palette): o proprio valor ou, na paleta "tema" (-1..-6 = destaque 1..6), a cor do tema da pasta.
Private Function PaletteRGB(cht As Chart, v As Variant) As Long
    On Error GoTo Fail
    If v >= 0 Then
        PaletteRGB = CLng(v)
    Else
        PaletteRGB = Registry_Workbook(cht).Theme.ThemeColorScheme.Colors(msoThemeAccent1 + (-v - 1)).RGB
    End If
    Exit Function
Fail:
    PaletteRGB = RGB(68, 114, 196)
End Function

' Aplica a escala calculada no bloco (linha 2: Unit, AxisMin, AxisMax) ao eixo de valores primario (o das barras e das auxiliares XY), alinha um
' eixo X secundario se o Excel criou um, fixa o eixo Y secundario em 0..n, troca o formato das celulas TickLbl pela unidade (w: dd/mm; m: mmm/yy;
' q, y: General, o texto ja vem pronto), refaz as folgas (Gantt_EnsureRoom) e religa/apaga os rotulos das marcas (passo: ScaleLabelStep) e das
' barras que passaram a ter (ou deixaram de ter) posicao. So escreve o que mudou (chamada a cada calculo do grafico: barata e sem laco de
' eventos). Nunca propaga erro (registra).
Public Sub Gantt_RefreshScale(cht As Chart)
    Dim blk As Range, body As Range, u As Variant, mn As Variant, mx As Variant, ax As Axis, n As Long, want As String, s As Series, stepK As Long
    On Error GoTo Fail
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    If Kinds_Family(CStr(blk.Cells(1, 2).Value)) <> "gantt" Then Exit Sub
    Set body = Block_Body(blk)
    n = GanttCount(cht)
    u = body.Cells(2, GT_COL_UNIT).Value
    mn = body.Cells(2, GT_COL_AXMIN).Value
    mx = body.Cells(2, GT_COL_AXMAX).Value
    stepK = 1
    If IsError(u) Or IsError(mn) Or IsError(mx) Then
        Errors_Log "Gantt_RefreshScale" & vbTab & "escala sem datas validas (eixo mantido)"
    ElseIf CDbl(mx) > CDbl(mn) Then
        Set ax = cht.Axes(xlValue, xlPrimary)
        If ax.MinimumScaleIsAuto Or ax.MaximumScaleIsAuto Or ax.MinimumScale <> CDbl(mn) Or ax.MaximumScale <> CDbl(mx) Then
            If CDbl(mn) >= ax.MaximumScale Then      ' o novo minimo passa do maximo atual: o maximo primeiro (minimo >= maximo nao vale)
                ax.MaximumScale = CDbl(mx): ax.MinimumScale = CDbl(mn)
            Else
                ax.MinimumScale = CDbl(mn): ax.MaximumScale = CDbl(mx)
            End If
        End If
        Select Case CStr(u)
            Case "w": want = "dd/mm"
            Case "m": want = "mmm/yy"
            Case Else: want = "General"
        End Select
        If CStr(body.Cells(2, GT_COL_TICKLBL).NumberFormat) <> want Then
            With body.Cells(2, GT_COL_TICKLBL).Resize(GT_SCALE_SLOTS, 1)
                ' medido: de General direto para mmm/yy a celula mostra "jan/26" mas o rotulo ligado mostra o mes em INGLES ("Jan/26"); passando
                ' antes por um formato de data o rotulo segue a celula (o nome do mes da localidade)
                If want = "mmm/yy" Then .NumberFormat = "dd/mm"
                .NumberFormat = want
            End With
        End If
    End If
    Chart_SyncSecondaryX cht
    If cht.HasAxis(xlValue, xlSecondary) Then
        With cht.Axes(xlValue, xlSecondary)
            If .MinimumScaleIsAuto Or .MinimumScale <> 0 Then .MinimumScale = 0
            If .MaximumScaleIsAuto Or .MaximumScale <> n Then .MaximumScale = n
        End With
    End If
    Gantt_EnsureRoom cht                             ' antes do passo dos rotulos: a folga a direita muda a largura interna
    If Not IsError(u) And Not IsError(mn) And Not IsError(mx) Then stepK = ScaleLabelStep(cht, CStr(u), CDbl(mn), CDbl(mx))
    Set s = Chart_FindSeries(cht, "tqq_scale")
    If Not s Is Nothing Then GanttSyncLabels s, body.Cells(2, GT_COL_TICK).Resize(GT_SCALE_SLOTS, 1), body.Cells(2, GT_COL_TICKLBL).Resize(GT_SCALE_SLOTS, 1), xlLabelPositionAbove, 8, False, stepK
    Set s = Chart_FindSeries(cht, "tqq_lbl")
    If Not s Is Nothing Then GanttSyncLabels s, body.Cells(2, GT_COL_LBLX).Resize(n, 1), body.Cells(2, GT_COL_LBLTXT).Resize(n, 1), xlLabelPositionRight, 9, False
    Exit Sub
Fail:
    Errors_Log "Gantt_RefreshScale" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Folgas do Gantt (so aumentam; 3 passadas porque a area interna nem sempre acompanha 1:1 a externa): acima da area interna, InsideTop >=
' GT_TOP_ROOM para os rotulos da escala (os colchetes pedem mais, em Ann_EnsureRoom); a direita, com os rotulos ligados, cada rotulo de barra
' (a direita de Fim + 1 ou do marco) cabe antes da borda do grafico: perto do fim do eixo o Excel empurraria o rotulo para cima da barra (medido).
' Largura do rotulo estimada pelo texto (Labels_EstimateWidth, que superestima o rotulo de 9 pt sem negrito), nunca lida do rotulo (ler a
' geometria de um rotulo a congela ate o proximo desenho). Fora do Gantt nao faz nada. Chamada no Insert, a cada Gantt_RefreshScale e por
' Chart_EnsureRoomAfterChange (o layout automatico da area de plotagem a perderia). Nunca propaga erro.
Public Sub Gantt_EnsureRoom(cht As Chart)
    Dim k As Long, delta As Double, body As Range, n As Long, i As Long, x() As Double, w() As Double, v As Variant, d As Double, hasAny As Boolean
    On Error Resume Next
    If Kinds_Family(Registry_ChartType(cht)) <> "gantt" Then Exit Sub
    For k = 1 To 3
        delta = GT_TOP_ROOM - cht.PlotArea.InsideTop
        If delta <= 0.01 Then Exit For
        cht.PlotArea.Height = cht.PlotArea.Height - delta
        cht.PlotArea.Top = cht.PlotArea.Top + delta
    Next
    If Err.Number <> 0 Then Errors_Log "Gantt_EnsureRoom/topo" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    If Registry_GetOption(cht, "labels", "1") <> "1" Then Exit Sub
    Set body = Block_Body(Registry_Block(cht))
    n = GanttCount(cht)
    If body Is Nothing Or n < 1 Then Exit Sub
    ReDim x(1 To n): ReDim w(1 To n)
    For i = 1 To n
        v = body.Cells(i + 1, GT_COL_LBLX).Value2
        If Not IsError(v) And Not IsEmpty(v) Then
            x(i) = Util_ToDouble(v)
            w(i) = Labels_EstimateWidth(Util_CellLabel(body.Cells(i + 1, GT_COL_LBLTXT)))
            hasAny = True
        End If
    Next
    If Not hasAny Then Exit Sub
    For k = 1 To 3
        delta = 0
        For i = 1 To n
            If w(i) > 0 Then
                d = Geom_ChartValueToX(cht, x(i)) + GT_LABEL_OFFSET + w(i) + GT_LABEL_MARGIN - cht.ChartArea.Width
                If d > delta Then delta = d
            End If
        Next
        If delta <= 0.5 Then Exit For
        If Not Deco_ShrinkPlotWidth(cht, delta) Then Exit For
    Next
    If Err.Number <> 0 Then Errors_Log "Gantt_EnsureRoom/direita" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Datas de inicio e fim (1..n) das atividades, lidas das colunas Inicio/Fim do bloco; NA(), vazio ou texto = 0. Usada pelos colchetes
' (Ann_Refresh). Com o grafico sem bloco ou sem atividades devolve arrays 0..0.
Public Sub Gantt_Dates(cht As Chart, ByRef starts() As Double, ByRef ends() As Double)
    Dim blk As Range, body As Range, n As Long, i As Long
    Set blk = Registry_Block(cht)
    If Not blk Is Nothing Then n = GanttCount(cht)
    If n < 1 Then
        ReDim starts(0 To 0): ReDim ends(0 To 0)
        Exit Sub
    End If
    Set body = Block_Body(blk)
    If n > body.Rows.Count - 1 Then n = body.Rows.Count - 1
    ReDim starts(1 To n): ReDim ends(1 To n)
    For i = 1 To n
        starts(i) = CellDate(body.Cells(i + 1, GT_COL_START))
        ends(i) = CellDate(body.Cells(i + 1, GT_COL_END))
    Next
End Sub

Private Function CellDate(c As Range) As Double
    Dim v As Variant
    v = c.Value2
    If IsError(v) Or IsEmpty(v) Then Exit Function
    CellDate = Util_ToDouble(v)
End Function
