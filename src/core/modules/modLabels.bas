Attribute VB_Name = "modLabels"
Option Explicit
Private mRefreshing As Boolean
Private mLinking As Long        ' Labels_LinkPoints em andamento: um refresh reentrante (evento Calculate no meio da ligacao) nao esconde rotulos
Private Const SIDE_LABEL_OFFSET As Double = 6   ' o rotulo "Left"/"Right" de um ponto XY comeca ~5,5 pt depois dele (medido nos totais XY)
Private Const SEGMENT_TEXT_HEIGHT As Double = 12 ' altura (pt) do texto de um rotulo de 9 pt: segmento de cascata mais baixo que isso esconde o rotulo

' Liga o rótulo de cada ponto à célula correspondente (rótulo "vivo"). Posição: acima/abaixo conforme o sinal
' (fixedPos = 0) ou a posição fixa informada (ex.: xlLabelPositionCenter em segmentos empilhados). Com horizontal
' (auxiliares XY de barras), a posição por sinal é à direita (>= 0) ou à esquerda (< 0) do ponto.
' fontColor >= 0 define a cor da fonte (contraste com o preenchimento do segmento).
' Isolamento por ponto: quando o valor PLOTADO do ponto e #N/D (Series.Values devolve Empty para ele, medido no Excel 2019;
' as auxiliares de total da cascata com varias series so plotam as linhas de total) e o Excel recusa o rotulo desse ponto, o
' erro e registrado e os demais pontos continuam ligados. Em ponto plotado o erro propaga (o Insert desfaz tudo).
' Enquanto liga (mLinking), SyncThinSegmentLabels nao mexe em rotulo nenhum: cada escrita dispara o Calculate do grafico e o refresh do
' observador, no meio de LinkOnePoint, apagaria o rotulo que esta sendo ligado.
Public Sub Labels_LinkPoints(s As Series, valueCells As Range, numFmt As String, Optional fixedPos As Long = 0, Optional fontColor As Long = -1, Optional horizontal As Boolean = False)
    Dim i As Long, n As Long, sheetRef As String, plotted As Variant, isNa As Boolean, en As Long, es As String, ed As String
    n = valueCells.Rows.Count
    mLinking = mLinking + 1
    On Error GoTo Fail
    s.HasDataLabels = True
    valueCells.NumberFormatLocal = Style_LocalizeCode(numFmt)
    sheetRef = "='" & Replace(valueCells.Worksheet.Name, "'", "''") & "'!"
    On Error Resume Next
    plotted = s.Values
    If Err.Number <> 0 Then Err.Clear: plotted = Empty
    On Error GoTo Fail
    For i = 1 To n
        isNa = False
        If IsArray(plotted) Then
            If LBound(plotted) + i - 1 <= UBound(plotted) Then isNa = IsError(plotted(LBound(plotted) + i - 1)) Or IsEmpty(plotted(LBound(plotted) + i - 1))
        End If
        If isNa Then
            On Error Resume Next
            LinkOnePoint s, i, sheetRef & valueCells.Cells(i, 1).Address(True, True), valueCells.Cells(i, 1).Value, numFmt, fixedPos, fontColor, horizontal
            If Err.Number <> 0 Then Errors_Log "Labels_LinkPoints" & vbTab & "ponto " & i & " (#N/D)" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
            On Error GoTo Fail
        Else
            LinkOnePoint s, i, sheetRef & valueCells.Cells(i, 1).Address(True, True), valueCells.Cells(i, 1).Value, numFmt, fixedPos, fontColor, horizontal
        End If
    Next
    mLinking = mLinking - 1
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    mLinking = mLinking - 1
    Err.Raise en, es, ed
End Sub

Private Sub LinkOnePoint(s As Series, i As Long, cellRef As String, v As Variant, numFmt As String, fixedPos As Long, fontColor As Long, horizontal As Boolean)
    With s.Points(i)
        .HasDataLabel = True
        .DataLabel.Formula = cellRef
        Style_AssignNumberFormat .DataLabel, numFmt
        .DataLabel.Font.Size = 9
        If fontColor >= 0 Then .DataLabel.Font.Color = fontColor Else .DataLabel.Font.Color = RGB(64, 64, 64)
        If fixedPos <> 0 Then
            .DataLabel.Position = fixedPos
        ElseIf horizontal Then
            .DataLabel.Position = IIf(Util_ToDouble(v) < 0, xlLabelPositionLeft, xlLabelPositionRight)
        Else
            .DataLabel.Position = IIf(Util_ToDouble(v) < 0, xlLabelPositionBelow, xlLabelPositionAbove)
        End If
    End With
End Sub

' Largura estimada (pt) de um rótulo de dados de 9 pt em negrito com o texto dado: 6 pt de margem interna do rótulo +
' dígitos (5,0 cada) + separadores e espaço (2,6) + sinais e parênteses (3,2) + demais caracteres (6,0). Medido no Excel
' (Calibri 9 negrito): dígito 4,56, separador 2,4, margem 6,0 ("1.500" = 26,6 pt; "1.234.567" = 42,7 pt); as constantes
' superestimam um pouco de propósito (outras fontes de tema são mais largas). É uma ESTIMATIVA e não DataLabel.Width porque
' ler a geometria de um rótulo fixa a posição lida até o próximo desenho do gráfico (o Excel invisível dos testes nunca
' redesenha): leitura feita na criação deixaria Left/Top defasados depois de mudar dados ou tamanho.
Public Function Labels_EstimateWidth(txt As String) As Double
    Dim i As Long, ch As String, w As Double
    w = 6
    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        Select Case ch
            Case "0" To "9": w = w + 5
            Case ".", ",", " ", ChrW(160): w = w + 2.6
            Case "-", "+", "(", ")": w = w + 3.2
            Case Else: w = w + 6
        End Select
    Next
    Labels_EstimateWidth = w
End Function

' Maior largura estimada entre os rótulos da série auxiliar de totais (0 sem a série ou sem rótulo legível).
' Lê só o TEXTO de cada rótulo (não a geometria).
Public Function Labels_MaxTotalWidth(cht As Chart) As Double
    Labels_MaxTotalWidth = Labels_MaxLabelWidth(cht, "tqq_total")
End Function

' Maior largura estimada (Labels_EstimateWidth) entre os rótulos da série auxiliar seriesName (0 sem a série ou sem rótulo legível).
' Lê só o TEXTO de cada rótulo (não a geometria).
Public Function Labels_MaxLabelWidth(cht As Chart, seriesName As String) As Double
    Dim s As Series, i As Long, w As Double, best As Double, t As String
    Set s = Chart_FindSeries(cht, seriesName)
    If s Is Nothing Then Exit Function
    On Error Resume Next
    For i = 1 To s.Points.Count
        t = vbNullString
        t = s.Points(i).DataLabel.Text
        If Len(t) > 0 Then
            w = Labels_EstimateWidth(t)
            If w > best Then best = w
        End If
    Next
    On Error GoTo 0
    Labels_MaxLabelWidth = best
End Function

' Combinação: a série de LINHA mostra o formato da própria célula de origem (NumberFormatLinked), não o numfmt do gráfico.
' O numfmt é um só para o gráfico e vale para as colunas e os totais (milhares); a linha costuma estar em outra unidade
' (percentual sobre colunas de milhares) e com o numfmt "int" mostraria "0" para 0,42. Quem quiser outro formato muda a
' célula. Falso em qualquer outro tipo e nas séries auxiliares.
Public Function Labels_IsSourceFormatSeries(cht As Chart, s As Series) As Boolean
    If Style_IsHelperSeries(s) Then Exit Function
    If s.ChartType <> xlLine And s.ChartType <> xlLineMarkers Then Exit Function
    Labels_IsSourceFormatSeries = (Registry_ChartType(cht) = "combo")
End Function

' Células de valores da coluna da linha da combinação (a última coluna da tabela de origem, sem o cabeçalho), lidas do endereço
' da origem gravado na meta do bloco; Nothing se não for possível ler.
Private Function ComboLineCells(cht As Chart) As Range
    Dim blk As Range, a As String, bang As Long, sh As String, src As Range
    On Error GoTo Fail
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    a = CStr(blk.Cells(1, 3).Value)
    bang = InStrRev(a, "!")
    If bang < 2 Then Exit Function
    sh = Left$(a, bang - 1)
    If Left$(sh, 1) = "'" And Right$(sh, 1) = "'" Then sh = Replace(Mid$(sh, 2, Len(sh) - 2), "''", "'")
    Set src = Registry_Workbook(cht).Worksheets(sh).Range(Mid$(a, bang + 1))
    If src.Rows.Count < 2 Or src.Columns.Count < 2 Then Exit Function
    Set ComboLineCells = src.Cells(2, src.Columns.Count).Resize(src.Rows.Count - 1, 1)
    Exit Function
Fail:
    Set ComboLineCells = Nothing
End Function

' Formato (código neutro de Style_NumberFormat) de reserva dos rótulos da linha da combinação, ou "" para seguir a célula.
' Célula de origem COM formato (percentual, moeda, ...) ou formatos misturados: "" (o rótulo é vinculado ao formato da célula).
' Célula em Geral: o rótulo mostraria os dígitos crus (0,423456789), então a regra de reserva olha os valores da linha:
' todos inteiros (ou nenhum número) = "int"; todos com módulo <= 1 = "pct1" (razão/percentual: 0,42 vira 42,3%); os demais
' = "dec2". Quem quiser outro formato muda o formato da célula.
Public Function Labels_ComboLineFormat(cht As Chart) As String
    Dim c As Range, f As Variant, v As Variant, one() As Variant, i As Long, d As Double, anyNum As Boolean, allInt As Boolean, allSmall As Boolean
    On Error GoTo Fail
    Set c = ComboLineCells(cht)
    If c Is Nothing Then Exit Function
    f = c.NumberFormat
    If IsNull(f) Then Exit Function
    If CStr(f) <> "General" Then Exit Function
    allInt = True: allSmall = True
    If c.Cells.Count = 1 Then
        ReDim one(1 To 1, 1 To 1)
        one(1, 1) = c.Value2
        v = one
    Else
        v = c.Value2
    End If
    For i = 1 To c.Rows.Count
        If IsNumeric(v(i, 1)) And Not IsEmpty(v(i, 1)) And VarType(v(i, 1)) <> vbString And Not IsError(v(i, 1)) Then
            anyNum = True: d = CDbl(v(i, 1))
            If d <> Int(d) Then allInt = False
            If Abs(d) > 1 Then allSmall = False
        End If
    Next
    If Not anyNum Or allInt Then
        Labels_ComboLineFormat = Style_NumberFormat("int")
    ElseIf allSmall Then
        Labels_ComboLineFormat = Style_NumberFormat("pct1")
    Else
        Labels_ComboLineFormat = Style_NumberFormat("dec2")
    End If
    Exit Function
Fail:
    Labels_ComboLineFormat = ""
End Function

' Aplica a um DataLabels/DataLabel da linha da combinação o formato certo: vinculado à célula, ou o de reserva (Geral).
Public Sub Labels_ApplyLineSourceFormat(target As Object, cht As Chart)
    Dim fb As String
    fb = Labels_ComboLineFormat(cht)
    If Len(fb) = 0 Then target.NumberFormatLinked = True Else Style_AssignNumberFormat target, fb
End Sub

' Mantém o formato da linha da combinação em dia com a célula (o usuário muda o formato da célula depois de inserir o gráfico):
' só escreve quando mudou. A série inteira e o último ponto (o rótulo do fim da linha tem formato próprio).
' Nunca propaga erro (chamada por eventos).
Public Sub Labels_RefreshComboLineFormat(cht As Chart)
    Dim s As Series, want As String, n As Long
    On Error GoTo Done
    If Registry_ChartType(cht) <> "combo" Then Exit Sub
    For Each s In cht.SeriesCollection
        If Labels_IsSourceFormatSeries(cht, s) Then
            want = Labels_ComboLineFormat(cht)
            If s.HasDataLabels Then SyncFormat s.DataLabels, want
            n = s.Points.Count
            If n >= 1 Then
                If s.Points(n).HasDataLabel Then SyncFormat s.Points(n).DataLabel, want
            End If
        End If
    Next
Done:
    If Err.Number <> 0 Then Errors_Log "Labels_RefreshComboLineFormat" & vbTab & Err.Number & vbTab & Err.Description
End Sub

Private Sub SyncFormat(target As Object, want As String)
    If Len(want) = 0 Then
        If Not target.NumberFormatLinked Then target.NumberFormatLinked = True
    Else
        If target.NumberFormatLinked Then
            Style_AssignNumberFormat target, want
        ElseIf target.NumberFormatLocal <> Style_LocalizeCode(want) Then
            Style_AssignNumberFormat target, want
        End If
    End If
End Sub

' Formato (código neutro de Style_NumberFormat) dos rótulos de segmento. Em 100% os segmentos mostram a
' PARTICIPAÇÃO (valor / total da linha): pct0 sempre, pct1 só quando o usuário escolhe pct1. Nos demais, numfmt.
Public Function Labels_SegmentFormat(cht As Chart) As String
    Dim key As String
    key = Registry_GetOption(cht, "numfmt", "int")
    If Kinds_Is100(Registry_ChartType(cht)) Then key = IIf(key = "pct1", "pct1", "pct0")
    Labels_SegmentFormat = Style_NumberFormat(key)
End Function

' Formato dos rótulos de total. Em 100% os totais são valores absolutos: um formato percentual (pct0/pct1,
' que ali vale para as participações) não se aplica a eles e os totais ficam em "int".
Public Function Labels_TotalFormat(cht As Chart) As String
    Dim key As String
    key = Registry_GetOption(cht, "numfmt", "int")
    If Kinds_Is100(Registry_ChartType(cht)) And (key = "pct0" Or key = "pct1") Then key = "int"
    Labels_TotalFormat = Style_NumberFormat(key)
End Function

' Reaplica a posição conforme o sinal atual das séries auxiliares de rótulo: acima/abaixo (verticais) ou
' direita/esquerda (horizontais). Cada ponto é isolado: um ponto não plotado (NA()) não interrompe os demais.
' Cascata horizontal: um rótulo "à esquerda" (barra que desce, total negativo) cujo ponto fica perto do mínimo do eixo e que não cabe
' até a borda esquerda da área interna (onde começam os nomes das categorias; largura pelo texto, Labels_EstimateWidth) passa para o
' outro lado do ponto, sobre a própria barra, com fundo branco para ficar legível; com espaço volta à esquerda, sem o fundo.
' Cascata com várias séries: rótulos de segmento mais finos que o texto ficam ocultos (SyncThinSegmentLabels).
Public Sub Labels_RefreshPositions(cht As Chart)
    Dim blk As Range, body As Range, s As Series, col As Long, i As Long, n As Long, en As Long, ed As String, want As Long
    Dim kind As String, horiz As Boolean, sideCheck As Boolean, inside As Boolean
    If mRefreshing Then Exit Sub
    mRefreshing = True
    On Error GoTo Done
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then GoTo Done
    kind = CStr(blk.Cells(1, 2).Value)
    If Kinds_Family(kind) = "pie" Then GoTo Done            ' rotulos nativos da pizza: sem auxiliar e sem posicao por sinal
    If Kinds_Family(kind) = "mekko" Then GoTo Done          ' Mekko: posicoes fixas (centro, acima do topo, abaixo do zero); a coluna 3 e a largura
    If Kinds_Family(kind) = "gantt" Then GoTo Done          ' Gantt: rotulos sempre a direita (barras) e acima (escala); a coluna 3 e o Fim
    horiz = Kinds_IsHorizontal(kind)
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    If Kinds_Family(kind) = "waterfall" Then
        SyncThinSegmentLabels cht, body, horiz          ' varias series: rotulos de segmento finos demais ficam ocultos (independe de tqq_lbl)
        Set s = Chart_FindSeries(cht, "tqq_lbl"): col = WF_COL_LABELVAL
        sideCheck = horiz
    Else
        Set s = Chart_FindSeries(cht, "tqq_total"): col = 3     ' acima/abaixo segue a posição plotada (TotalPos), não o líquido
    End If
    If s Is Nothing Then GoTo Done
    If Not s.HasDataLabels Then GoTo Done
    For i = 1 To n
        If i <= s.Points.Count Then
            inside = False
            If horiz Then
                want = IIf(Util_ToDouble(body.Cells(i + 1, col).Value) < 0, xlLabelPositionLeft, xlLabelPositionRight)
                If sideCheck And want = xlLabelPositionLeft Then
                    If Not LeftLabelFits(cht, s.Points(i), Util_ToDouble(body.Cells(i + 1, WF_COL_LABELPOS).Value)) Then want = xlLabelPositionRight: inside = True
                End If
            Else
                want = IIf(Util_ToDouble(body.Cells(i + 1, col).Value) < 0, xlLabelPositionBelow, xlLabelPositionAbove)
            End If
            On Error Resume Next
            If s.Points(i).DataLabel.Position <> want Then s.Points(i).DataLabel.Position = want    ' só escreve se mudou (evita redesenho a cada cálculo)
            If sideCheck Then SyncLabelBackground s.Points(i).DataLabel, inside
            Err.Clear
            On Error GoTo Done
        End If
    Next
Done:
    If Err.Number <> 0 Then en = Err.Number: ed = Err.Description
    mRefreshing = False
    If en <> 0 Then Errors_Log "Labels_RefreshPositions" & vbTab & en & vbTab & ed
End Sub

' True quando o rótulo "Left" do ponto (no valor xv do eixo horizontal) cabe entre o ponto e a borda esquerda da área interna; True em
' qualquer dúvida (rótulo ilegível, sem texto): a posição pelo sinal fica.
Private Function LeftLabelFits(cht As Chart, p As Point, xv As Double) As Boolean
    Dim t As String
    LeftLabelFits = True
    On Error GoTo Done
    t = p.DataLabel.Text
    If Len(t) = 0 Then Exit Function
    LeftLabelFits = (Geom_ChartValCoord(cht, xv) - SIDE_LABEL_OFFSET - Labels_EstimateWidth(t) >= cht.PlotArea.InsideLeft)
Done:
End Function

' Cascata com várias séries: o rótulo de um segmento (tqq_lbl_j, centrado, ligado a D_j) mais fino que o texto encostaria nas bordas do
' segmento e ficaria ilegível: fica oculto. "Fino" pela escala atual: nas colunas, altura do segmento < SEGMENT_TEXT_HEIGHT; nas barras,
' largura < largura estimada do texto (Labels_EstimateWidth do valor no formato do gráfico). Quando o segmento volta a caber (dados ou tamanho
' mudaram) o rótulo é religado como em Labels_LinkPoints. Rótulos de texto vazio (delta zero) ficam como estão. Só escreve quando muda; cada
' ponto é isolado; nunca propaga erro.
Private Sub SyncThinSegmentLabels(cht As Chart, body As Range, horiz As Boolean)
    Dim ns As Long, j As Long, i As Long, n As Long, s As Series, mn As Double, mx As Double, k As Double, dcol As Long, c As Range
    Dim txt As String, need As Double, want As Boolean, fmt As String, sheetRef As String, fc As Long, v As Double
    On Error GoTo Done
    If mLinking > 0 Then Exit Sub                   ' refresh reentrante no meio de Labels_LinkPoints: o refresh final refaz
    ns = Waterfall_SeriesCount(cht)
    If ns < 2 Then Exit Sub
    mn = cht.Axes(xlValue, xlPrimary).MinimumScale: mx = cht.Axes(xlValue, xlPrimary).MaximumScale
    If mx <= mn Then Exit Sub
    If horiz Then k = cht.PlotArea.InsideWidth / (mx - mn) Else k = cht.PlotArea.InsideHeight / (mx - mn)
    n = body.Rows.Count - 1
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    sheetRef = "='" & Replace(body.Worksheet.Name, "'", "''") & "'!"
    For j = 1 To ns
        Set s = Chart_FindSeries(cht, "tqq_lbl_" & j)
        If Not s Is Nothing Then
            dcol = Waterfall_SeriesCol(j, "D")
            fc = Style_ContrastFont(Style_DataSeriesColor(cht, j))
            For i = 1 To n
                If i <= s.Points.Count Then
                    Set c = body.Cells(i + 1, dcol)
                    v = Util_ToDouble(c.Value)
                    txt = Format$(v, fmt)
                    If Len(txt) > 0 Then
                        If horiz Then need = Labels_EstimateWidth(txt) Else need = SEGMENT_TEXT_HEIGHT
                        want = (Abs(v) * k >= need)
                        On Error Resume Next
                        If want <> s.Points(i).HasDataLabel Then
                            If want Then LinkOnePoint s, i, sheetRef & c.Address(True, True), c.Value, fmt, xlLabelPositionCenter, fc, horiz Else s.Points(i).HasDataLabel = False
                        End If
                        If Err.Number <> 0 Then Errors_Log "Labels/segmento fino" & vbTab & s.Name & " ponto " & i & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
                        On Error GoTo Done
                    End If
                End If
            Next
        End If
    Next
Done:
    If Err.Number <> 0 Then Errors_Log "Labels/segmentos finos" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Fundo branco no rótulo que ficou sobre a própria barra (inside) e sem fundo nos demais; só escreve quando muda. Medido: o preenchimento
' só passa a valer (Fill.Visible lido de volta True) depois de escrever a cor, mesmo que ela já seja o branco: as duas escritas vão juntas.
Private Sub SyncLabelBackground(dl As DataLabel, inside As Boolean)
    If inside Then
        If dl.Format.Fill.Visible <> msoTrue Or dl.Format.Fill.ForeColor.RGB <> RGB(255, 255, 255) Then
            dl.Format.Fill.Visible = msoTrue
            dl.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End If
    ElseIf dl.Format.Fill.Visible = msoTrue Then
        dl.Format.Fill.Visible = msoFalse
    End If
End Sub

' Número de colunas de participação (cabeçalho "Part1", "Part2", ...) seguidas a partir da coluna 4 do corpo do bloco.
Private Function PartColumnCount(body As Range) As Long
    Dim j As Long
    For j = 4 To body.Columns.Count
        If Left$(CStr(body.Cells(1, j).Value), 4) <> "Part" Then Exit For
        PartColumnCount = PartColumnCount + 1
    Next
End Function

' Rótulos ligados a células (Formula = célula da planilha _ThinQQ) exibem o TEXTO da célula, então o formato
' numérico precisa estar também na célula (gravado por NumberFormatLocal, como nos rótulos: o texto exibido segue os separadores do Excel).
Public Sub Labels_ApplyCellFormat(cht As Chart, fmt As String)
    Dim blk As Range, body As Range, col As Long, n As Long, kind As String, nPart As Long, j As Long, nSeries As Long, segFmt As String
    On Error GoTo Fail
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    kind = CStr(blk.Cells(1, 2).Value)
    If Kinds_Family(kind) = "pie" Then Exit Sub                 ' rotulos nativos da pizza (Pie_ApplyNumberFormat): nenhuma celula do bloco e rotulo
    If Kinds_Family(kind) = "mekko" Then                        ' Mekko: LblVal_j e TotLbl (a coluna 2 e a Categoria, que nao recebe formato)
        Mekko_ApplyCellFormat cht
        Exit Sub
    End If
    If Kinds_Family(kind) = "gantt" Then Exit Sub              ' Gantt: sem formato numerico; os rotulos mostram datas (dd/mm, mmm/yy) ou texto da tabela
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    If n < 1 Then Exit Sub
    If Kinds_Family(kind) = "waterfall" Then col = WF_COL_LABELVAL Else col = 2
    body.Cells(2, col).Resize(n, 1).NumberFormatLocal = Style_LocalizeCode(fmt)
    ' cascata com varias series: os rotulos de segmento ligam a D_j e o rotulo de total a coluna Total (Waterfall_SetLabels)
    If Kinds_Family(kind) = "waterfall" Then
        nSeries = Waterfall_SeriesCount(cht)
        If nSeries > 1 Then
            body.Cells(2, WF_COL_TOTAL).Resize(n, 1).NumberFormatLocal = Style_LocalizeCode(fmt)
            For j = 1 To nSeries
                body.Cells(2, Waterfall_SeriesCol(j, "D")).Resize(n, 1).NumberFormatLocal = Style_LocalizeCode(fmt)
            Next
        End If
    End If
    ' 100%: as colunas de participação Part_j (a partir da 4) seguem o formato dos segmentos, não o dos totais. Depois
    ' delas vêm Pos e VLine (que não são rótulos): em blocos com Pos/VLine são Columns.Count - 5 colunas; os cabeçalhos
    ' "Part" decidem, o que também serve a blocos da Fase 1 (sem Pos/VLine).
    If Kinds_Family(kind) = "stacked" And Kinds_Is100(kind) Then
        nPart = PartColumnCount(body)
        If nPart > 0 Then body.Cells(2, 4).Resize(n, nPart).NumberFormatLocal = Style_LocalizeCode(Labels_SegmentFormat(cht))
    End If
    ' area: a coluna Total (col 2, acima) e as colunas Val_j (rotulos de segmento: o valor; em area100, a participacao, com o
    ' formato dos segmentos). A familia "combo" e a empilhada (col 2; sem Part, ela nao e 100%).
    If Kinds_Family(kind) = "area" Then
        segFmt = Style_LocalizeCode(Labels_SegmentFormat(cht))
        For j = 1 To Area_SeriesCount(blk)
            body.Cells(2, Area_ValCol(blk, j)).Resize(n, 1).NumberFormatLocal = segFmt
        Next
    End If
    If Registry_GetOption(cht, "vline", "0") = "1" Then Deco_ApplyVLineFormat cht     ' a media da linha de valor ("Média N") segue o formato dos numeros (ligar a linha ja grava o formato)
    Exit Sub
Fail:
    Errors_Log "Labels_ApplyCellFormat" & vbTab & Err.Number & vbTab & Err.Description
End Sub
