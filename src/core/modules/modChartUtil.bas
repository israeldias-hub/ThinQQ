Attribute VB_Name = "modChartUtil"
Option Explicit

' Gancho de teste: quando True (e BUILD_CONFIG = "Test"), os Insert falham logo apos alocar o bloco,
' para verificar o rollback. Nunca e usado fora dos testes.
Public Util_FailAfterBlockForTest As Boolean
' Gancho de teste (so com BUILD_CONFIG = "Test"): quando True, Area_Insert e Combo_Insert falham logo DEPOIS de o grafico existir
' (o gancho acima falha antes de existir), para verificar que o rollback apaga o grafico e o bloco.
Public Util_FailAfterChartForTest As Boolean

' Ganchos de teste da limpeza de blocos (CalcSheet_Cleanup); todos so tem efeito com BUILD_CONFIG = "Test".
' Util_FailChartReadForTest = True: a leitura de um grafico (no nivel do grafico) falha.
Public Util_FailChartReadForTest As Boolean
' Util_FailChartReadSiteForTest = READFAIL_*: a leitura falha no ponto indicado (0 = nenhum). Cada ponto e um lugar
' onde a limpeza le o objeto Excel e, se a leitura falhar, nada pode ser apagado (falha fechada).
Public Util_FailChartReadSiteForTest As Long
Public Const READFAIL_SHAPE As Long = 1         ' ler uma forma da planilha (Type/HasChart)
Public Const READFAIL_GROUP As Long = 2         ' ler um item de grupo de formas
Public Const READFAIL_CHARTSHEET As Long = 3    ' ler uma planilha de grafico (wb.Charts)
Public Const READFAIL_CHART As Long = 4         ' ler as series de um grafico
Public Const READFAIL_SERIES As Long = 5        ' ler Series.Formula
Public Const READFAIL_HASLABEL As Long = 6      ' ler Point.HasDataLabel
Public Const READFAIL_FORMULA As Long = 7       ' ler DataLabel.Formula
' Contador de pontos cujo rotulo a limpeza leu (so conta em Test) e limite de pontos por passada (0 = padrao).
Public Util_LabelReadsForTest As Long
Public Util_LabelBudgetForTest As Long

' Gancho de teste (so com BUILD_CONFIG = "Test"): nome do ChartObject cuja leitura da escala automatica (Deco_SameScale) falha ("" = nenhum).
Public Util_FailScaleReadChartForTest As String
' Gancho de teste (so com BUILD_CONFIG = "Test") dos rotulos de segmento finos da cascata (SyncThinSegmentLabels): "read:j.i" falha a leitura
' de Points(i).HasDataLabel da serie j, "link:j.i" falha ao religar esse ponto, "series:j" falha ao comecar a serie j ("" = nenhum).
Public Util_FailThinSyncForTest As String

Public Sub Util_TestHookAfterBlock()
    If BUILD_CONFIG = "Test" Then
        If Util_FailAfterBlockForTest Then Err.Raise vbObjectError + 9999, "ThinQQ", "Falha simulada (teste)"
    End If
End Sub

Public Sub Util_TestHookScaleRead(cht As Chart)
    If BUILD_CONFIG = "Test" Then
        If Len(Util_FailScaleReadChartForTest) > 0 Then
            If cht.Parent.Name = Util_FailScaleReadChartForTest Then Err.Raise vbObjectError + 9997, "ThinQQ", "Falha simulada na leitura da escala (teste)"
        End If
    End If
End Sub

Public Sub Util_TestHookAfterChart()
    If BUILD_CONFIG = "Test" Then
        If Util_FailAfterChartForTest Then Err.Raise vbObjectError + 9998, "ThinQQ", "Falha simulada depois do grafico (teste)"
    End If
End Sub

' Primeira posicao vertical livre (sem sobrepor outros graficos da planilha) para o retangulo proposto.
Public Function Chart_FreeTop(ws As Worksheet, l As Double, t As Double, w As Double, h As Double) As Double
    Dim co As ChartObject, moved As Boolean
    Do
        moved = False
        For Each co In ws.ChartObjects
            If l < co.Left + co.Width And co.Left < l + w And t < co.Top + co.Height And co.Top < t + h Then
                t = co.Top + co.Height + 12
                moved = True
            End If
        Next
    Loop While moved
    Chart_FreeTop = t
End Function

Public Sub Util_EnsureDir(path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub
    If Not fso.FolderExists(fso.GetParentFolderName(path)) Then Util_EnsureDir fso.GetParentFolderName(path)
    fso.CreateFolder path
End Sub

Public Function Util_FileName(path As String) As String
    Util_FileName = Mid$(path, InStrRev(path, "\") + 1)
End Function

Public Function Util_ToDouble(v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then
        If VarType(v) <> vbString Then Util_ToDouble = CDbl(v) Else Util_ToDouble = Val(v)
    End If
End Function

' Coluna de células -> array 1..n de Double (texto/vazio = 0).
Public Function Util_ColumnToDoubles(rng As Range) As Double()
    Dim out() As Double, i As Long, n As Long
    n = rng.Rows.Count
    ReDim out(1 To n)
    For i = 1 To n
        out(i) = Util_ToDouble(rng.Cells(i, 1).Value)
    Next
    Util_ColumnToDoubles = out
End Function

Public Function Chart_FindSeries(cht As Chart, name As String) As Series
    Dim s As Series
    For Each s In cht.SeriesCollection
        If s.Name = name Then Set Chart_FindSeries = s: Exit Function
    Next
End Function

' Texto da célula; evita '#####' de colunas estreitas.
Public Function Util_CellLabel(c As Range) As String
    Dim v As Variant, s As String
    v = c.Value2
    If IsError(v) Then Util_CellLabel = CStr(c.Text): Exit Function
    If VarType(v) = vbString Then Util_CellLabel = v: Exit Function
    s = CStr(c.Text)
    If Left$(s, 1) = "#" Then
        If VarType(c.Value) = vbDate Then s = Format$(c.Value, "Short Date") Else s = CStr(v)
    End If
    Util_CellLabel = s
End Function

' Nome do aplicativo no registro do atualizador (HKCU\Software\VB and VBA Program Settings\<nome>\Updater), o mesmo do loader
' (Settings_AppName): o build Test usa um ramo proprio e nunca le nem grava as configuracoes reais do usuario.
Public Function Util_UpdaterRegApp() As String
    If BUILD_CONFIG = "Test" Then Util_UpdaterRegApp = "ThinQQ.Test" Else Util_UpdaterRegApp = "ThinQQ"
End Function

' SHA-256 (hex minusculo) de um arquivo, em VBA puro (modSha256; sem processo filho). Retorna "" em qualquer falha
' (arquivo ausente/ilegivel): quem chama trata como nao verificado.
Public Function Util_Sha256File(path As String) As String
    On Error GoTo Fail
    Util_Sha256File = Sha256_File(path)
    Exit Function
Fail:
    Util_Sha256File = ""
End Function

' Quantidade de células da referência de categorias (2o argumento de SERIES) da série, lida na fórmula da série:
' reflete linhas inseridas/removidas na origem mesmo quando Series.XValues não mostra a diferença.
' Retorna -1 se o argumento não for uma referência simples a uma planilha da pasta wb.
Public Function Util_SeriesCategoryRefCount(s As Series, wb As Workbook) As Long
    Dim r As Range
    On Error GoTo Fail
    Util_SeriesCategoryRefCount = -1
    Set r = Util_SeriesArgRange(s, wb, 2)
    If Not r Is Nothing Then Util_SeriesCategoryRefCount = r.Cells.Count
    Exit Function
Fail:
    Util_SeriesCategoryRefCount = -1
End Function

' Faixa do argumento argN (1 nome, 2 categorias, 3 valores) da fórmula VIVA da série (=SERIES(...)), lida no texto da fórmula: acompanha a
' planilha renomeada e as linhas inseridas. Nothing se o argumento não for uma referência simples a uma planilha da pasta wb (constante,
' outra pasta, nome definido) ou se a leitura falhar. Nunca propaga erro.
Public Function Util_SeriesArgRange(s As Series, wb As Workbook, argN As Long) As Range
    Dim f As String, i As Long, ch As String, inQ As Boolean, inS As Boolean, depth As Long
    Dim k As Long, startPos As Long, arg As String, bang As Long, sh As String
    On Error GoTo Fail
    f = s.Formula
    k = 1
    startPos = InStr(f, "(") + 1
    If startPos < 2 Then Exit Function
    For i = startPos To Len(f)
        ch = Mid$(f, i, 1)
        If ch = """" And Not inS Then
            inQ = Not inQ
        ElseIf ch = "'" And Not inQ Then
            inS = Not inS
        ElseIf Not inQ And Not inS Then
            If ch = "(" Or ch = "{" Then
                depth = depth + 1
            ElseIf ch = "}" Then
                depth = depth - 1
            ElseIf ch = ")" And depth > 0 Then
                depth = depth - 1
            ElseIf (ch = "," And depth = 0) Or (ch = ")" And depth = 0) Then
                If k = argN Then arg = Mid$(f, startPos, i - startPos): Exit For
                k = k + 1
                startPos = i + 1
            End If
        End If
    Next
    If Len(arg) = 0 Then Exit Function
    bang = InStrRev(arg, "!")
    If bang < 2 Then Exit Function
    sh = Left$(arg, bang - 1)
    If InStr(sh, "[") > 0 Then Exit Function
    If Left$(sh, 1) = "'" And Right$(sh, 1) = "'" Then sh = Replace(Mid$(sh, 2, Len(sh) - 2), "''", "'")
    Set Util_SeriesArgRange = wb.Worksheets(sh).Range(Mid$(arg, bang + 1))
    Exit Function
Fail:
    Set Util_SeriesArgRange = Nothing
End Function

' Serie auxiliar (prefixo tqq_) para rotulos/totais/marcadores, orientavel:
'  - posCells Is Nothing (graficos verticais): linha invisivel (sem linha e sem marcador) no grupo primario, com
'    Values = valueCells (um ponto por categoria).
'  - posCells informado (barras horizontais): em barras o Excel nao combina com linha, so com dispersao (XY) no grupo
'    SECUNDARIO. X = valueCells (posicao no eixo de valores das barras), Y = posCells (posicao da categoria, n - i + 0,5;
'    o eixo Y secundario fixo em 0..n e oculto). Sem marcador e sem linha. Chama Chart_SetupSecondaryAxes.
' Nao mexe na legenda: quem chama decide (Chart_HideHelperLegendEntries). Em falha desfaz a serie e repropaga.
Public Function Chart_AddHelperSeries(cht As Chart, name As String, valueCells As Range, Optional posCells As Range = Nothing) As Series
    Dim s As Series, en As Long, es As String, ed As String
    On Error GoTo Fail
    Set s = cht.SeriesCollection.NewSeries
    s.Name = name
    If posCells Is Nothing Then
        s.Values = Block_Ref(valueCells)
        s.ChartType = xlLine
        s.Format.Line.Visible = msoFalse
        s.MarkerStyle = xlMarkerStyleNone
    Else
        s.ChartType = xlXYScatter          ' antes de XValues: uma serie XY precisa existir como XY para aceitar X numerico
        s.AxisGroup = xlSecondary
        s.XValues = Block_Ref(valueCells)
        s.Values = Block_Ref(posCells)
        s.MarkerStyle = xlMarkerStyleNone
        s.Format.Line.Visible = msoFalse
        Chart_SetupSecondaryAxes cht, posCells.Rows.Count
    End If
    Set Chart_AddHelperSeries = s
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not s Is Nothing Then s.Delete
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' Eixos secundarios das auxiliares XY de barras: eixo Y (posicao da categoria) fixo em 0..n, nao invertido,
' totalmente oculto (sem marcas, rotulos, linha e grade). Se existir eixo X secundario, alinha-o ao primario.
Public Sub Chart_SetupSecondaryAxes(cht As Chart, n As Long)
    cht.HasAxis(xlValue, xlSecondary) = True
    With cht.Axes(xlValue, xlSecondary)
        .MinimumScale = 0
        .MaximumScale = n
        .ReversePlotOrder = False
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabelPosition = xlTickLabelPositionNone
        .Format.Line.Visible = msoFalse
        .HasMajorGridlines = False
        .HasMinorGridlines = False
    End With
    If cht.HasAxis(xlCategory, xlSecondary) Then Chart_SyncSecondaryX cht
End Sub

' Alinha o eixo X secundario (se existir) ao eixo de valores primario: copia minimo/maximo e o oculta.
' Sem eixo X secundario nao faz nada (a serie XY usa o eixo horizontal primario). Nunca propaga erro.
Public Sub Chart_SyncSecondaryX(cht As Chart)
    Dim mn As Double, mx As Double
    On Error Resume Next
    If Not cht.HasAxis(xlCategory, xlSecondary) Then Exit Sub
    mn = cht.Axes(xlValue, xlPrimary).MinimumScale
    mx = cht.Axes(xlValue, xlPrimary).MaximumScale
    If Err.Number = 0 Then
        With cht.Axes(xlCategory, xlSecondary)
            .MinimumScale = mn
            .MaximumScale = mx
        End With
    End If
    Err.Clear
    With cht.Axes(xlCategory, xlSecondary)
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabelPosition = xlTickLabelPositionNone
        .Format.Line.Visible = msoFalse
        .HasMajorGridlines = False
        .HasMinorGridlines = False
    End With
    Err.Clear
End Sub

' Depois de uma mudanca posterior a criacao (legenda ligada/desligada, formato numerico dos rotulos, decoracoes, corte de eixo, Alterar
' dados; o redimensionamento a mao NAO a chama: Cht_Resize so reposiciona), a folga que o Insert deu aos totais (barras), as setas e colchetes, o rotulo "nome valor" do fim da linha, o indicador 100%, o
' rotulo da linha de valor, os totais acima do Mekko e a escala de datas do Gantt so valia para o estado do momento, e a area de plotagem de
' tamanho manual pode ter ficado sob a legenda: reexecuta as contas (legenda, totais, setas, fim da linha, decoracoes, Mekko, Gantt). So reduz a
' area de plotagem; sem nada disso ligado nao faz nada.
' Nunca propaga erro (so registra).
Public Sub Chart_EnsureRoomAfterChange(cht As Chart)
    On Error Resume Next
    Deco_FitPlotToLegend cht                      ' primeiro: as folgas abaixo contam a partir da borda livre (a legenda a direita)
    If Err.Number <> 0 Then Errors_Log "Chart_EnsureRoomAfterChange/legenda" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Stacked_EnsureRoomAfterChange cht
    If Err.Number <> 0 Then Errors_Log "Chart_EnsureRoomAfterChange/totais" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Ann_EnsureRoomAfterChange cht
    If Err.Number <> 0 Then Errors_Log "Chart_EnsureRoomAfterChange/setas" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Deco_EnsureLineEndRoom cht
    If Err.Number <> 0 Then Errors_Log "Chart_EnsureRoomAfterChange/fim da linha" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Deco_EnsureDecorationRoom cht                 ' indicador 100% e rotulo da linha de valor (Deco_SetLegend volta o layout ao automatico)
    If Err.Number <> 0 Then Errors_Log "Chart_EnsureRoomAfterChange/decoracoes" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Mekko_EnsureRoom cht                          ' Mekko: folga acima da area interna para os totais (nos demais tipos nao faz nada)
    If Err.Number <> 0 Then Errors_Log "Chart_EnsureRoomAfterChange/mekko" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Gantt_EnsureRoom cht                          ' Gantt: folga acima da area interna para a escala de datas (os colchetes: nas setas, acima)
    If Err.Number <> 0 Then Errors_Log "Chart_EnsureRoomAfterChange/gantt" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Break_EnsureRoom cht                          ' corte de eixo: folga para o rotulo do total da coluna cortada (sem corte nao faz nada)
    If Err.Number <> 0 Then Errors_Log "Chart_EnsureRoomAfterChange/corte" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Wfp_RefreshAfterDataChange cht                ' cascata comparativa: valor por ponto das setas e nomes/legenda no tamanho novo (nos demais tipos nao faz nada)
    If Err.Number <> 0 Then Errors_Log "Chart_EnsureRoomAfterChange/comparativa" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
End Sub

' Remove da legenda toda entrada de serie auxiliar (tqq_*). Idempotente e correta com varias auxiliares.
' Legenda sem entrada removida OU parcialmente removida (uma auxiliar nova entrou numa legenda que ja tinha
' outras removidas) e religada preservando posicao e fonte; depois cada entrada tqq_* e apagada do maior indice
' para o menor (apagar renumera as seguintes). O indice de cada serie segue a ordem dos grupos do grafico
' (ChartGroups, inclusive os do eixo secundario) e das series dentro de cada grupo, com a inversao descrita em
' LegendGroupReversed.
Public Sub Chart_HideHelperLegendEntries(cht As Chart)
    Dim lp As Long, lft As Double, tp As Double, wd As Double, ht As Double, en As Long, es As String, ed As String
    If Not cht.HasLegend Then Exit Sub
    lp = cht.Legend.Position
    If lp = xlLegendPositionRight Or lp = xlLegendPositionLeft Or lp = xlLegendPositionTop Or lp = xlLegendPositionBottom Then
        HideHelperEntries cht
        Exit Sub
    End If
    ' Canto (xlLegendPositionCorner) ou posicao livre (legenda arrastada; Position = xlLegendPositionCustom): a posicao
    ' nao diz se a lista esta na vertical (pilhas em colunas/areas listam invertidas) ou na horizontal (medido: a legenda
    ' arrastada a partir da direita/esquerda e o canto listam como a direita; a arrastada de cima/baixo, como embaixo). Para o
    ' mapeamento ser deterministico a legenda vai para a direita so durante o trabalho e volta ao lugar (canto: Position;
    ' livre: Left/Top/Width/Height). Nada a fazer (e a legenda nao se mexe) se as entradas auxiliares ja foram removidas.
    If Not HelperEntriesPending(cht) Then Exit Sub
    lft = cht.Legend.Left: tp = cht.Legend.Top: wd = cht.Legend.Width: ht = cht.Legend.Height
    cht.Legend.Position = xlLegendPositionRight
    On Error Resume Next
    HideHelperEntries cht
    en = Err.Number: es = Err.Source: ed = Err.Description
    Err.Clear
    If lp = xlLegendPositionCorner Then
        cht.Legend.Position = xlLegendPositionCorner
    Else
        cht.Legend.Left = lft: cht.Legend.Top = tp: cht.Legend.Width = wd: cht.Legend.Height = ht
    End If
    Err.Clear
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, es, ed
End Sub

' True quando ha entradas de serie auxiliar (tqq_*) ainda na legenda.
Private Function HelperEntriesPending(cht As Chart) As Boolean
    Dim s As Series, h As Long
    For Each s In cht.SeriesCollection
        If Style_IsHelperSeries(s) Then h = h + 1
    Next
    If h = 0 Then Exit Function
    HelperEntriesPending = (cht.Legend.LegendEntries.Count <> cht.SeriesCollection.Count - h)
End Function

Private Sub HideHelperEntries(cht As Chart)
    Dim g As ChartGroup, s As Series, k As Long, m As Long, h As Long, total As Long, cnt As Long
    Dim idx As Long, nIdx As Long, i As Long, j As Long, t As Long, pos As Long, fsz As Double, hidden() As Long, rev As Boolean
    If Not cht.HasLegend Then Exit Sub
    total = cht.SeriesCollection.Count
    For Each s In cht.SeriesCollection
        If Style_IsHelperSeries(s) Then h = h + 1
    Next
    If h = 0 Then Exit Sub
    cnt = cht.Legend.LegendEntries.Count
    If cnt = total - h Then Exit Sub                 ' ja esta feito
    If cnt <> total Then                             ' estado parcial: recomeca com todas as entradas
        pos = cht.Legend.Position
        fsz = cht.Legend.Font.Size
        cht.HasLegend = False
        cht.HasLegend = True
        cht.Legend.Position = pos
        cht.Legend.Font.Size = fsz
    End If
    ReDim hidden(1 To h)
    For Each g In cht.ChartGroups
        m = g.SeriesCollection.Count
        rev = LegendGroupReversed(cht, g)
        For k = 1 To m
            If Style_IsHelperSeries(g.SeriesCollection(k)) Then
                nIdx = nIdx + 1
                If nIdx <= h Then hidden(nIdx) = idx + IIf(rev, m - k + 1, k)
            End If
        Next
        idx = idx + m
    Next
    If idx <> total Or nIdx <> h Then                ' os grupos nao cobrem todas as series: nao arrisca apagar entrada de dados
        Errors_Log "Chart_HideHelperLegendEntries" & vbTab & "grupos x series" & vbTab & idx & "/" & total & " aux " & nIdx & "/" & h
        Exit Sub
    End If
    For i = 2 To nIdx                                ' ordem decrescente (a inversao por grupo desordena a lista)
        t = hidden(i)
        j = i - 1
        Do While j >= 1
            If hidden(j) >= t Then Exit Do
            hidden(j + 1) = hidden(j)
            j = j - 1
        Loop
        hidden(j + 1) = t
    Next
    For i = 1 To nIdx
        cht.Legend.LegendEntries(hidden(i)).Delete
    Next
End Sub

' O Excel lista as entradas da legenda na ordem VISUAL das series (medido no Excel 2019): em pilhas verticais
' (colunas e areas empilhadas, inclusive 100%) com a legenda a esquerda/direita a ultima serie aparece primeiro
' (indice 1); barras agrupadas com o eixo de categorias nao invertido tambem aparecem invertidas (com o eixo
' invertido, como o Style_ApplyBase deixa as barras, nao). Barras empilhadas, colunas agrupadas e linhas: ordem normal (medido
' tambem na cascata horizontal de varias series, com auxiliares intercaladas no grupo empilhado, nas 5 posicoes da legenda).
Private Function LegendGroupReversed(cht As Chart, g As ChartGroup) As Boolean
    Dim ct As Long, lp As Long
    On Error GoTo Fail
    If g.SeriesCollection.Count = 0 Then Exit Function
    ct = g.SeriesCollection(1).ChartType
    Select Case ct
        Case xlColumnStacked, xlColumnStacked100, xlAreaStacked, xlAreaStacked100
            lp = cht.Legend.Position
            LegendGroupReversed = (lp = xlLegendPositionRight Or lp = xlLegendPositionLeft)
        Case xlBarClustered
            LegendGroupReversed = Not cht.Axes(xlCategory, xlPrimary).ReversePlotOrder
    End Select
    Exit Function
Fail:
    LegendGroupReversed = False
End Function
