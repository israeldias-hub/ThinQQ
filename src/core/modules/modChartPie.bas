Attribute VB_Name = "modChartPie"
Option Explicit
' Pizza e rosca (kinds pie, doughnut). Uma serie so: a PRIMEIRA coluna de valores da tabela (as demais sao ignoradas; Core_Insert avisa).
' Bloco (corpo): linha 1 = cabecalhos; colunas 1 Categoria | 2 Valor (a origem se for numero positivo; texto, vazio, erro, zero e negativo
' valem 0: a fatia some e a tabela do usuario nao e tocada) | 3 Share (Valor / soma dos Valor; 0 se a soma e 0: informativo, os rotulos sao
' NATIVOS e nao ligados a celulas).
' Opcoes (Registry): numfmt, palette, legend (right|none), labels (0|1), n e, so na pizza, pieout (0|1: rotulos fora da fatia).
' Rotulos: valor (ShowValue) ou participacao (ShowPercentage: numfmt pct0/pct1); dentro (pizza: InsideEnd; rosca: posicao padrao, sobre
' o anel: o Excel recusa DataLabels.Position InsideEnd/OutsideEnd na rosca, erro 1004, e por isso a rosca nunca define a posicao e nao
' tem pieout) ou fora (pieout=1: OutsideEnd + nome da categoria + linhas de chamada, sem legenda).
' Legenda: HasLegend = (legend <> "none") And (pieout <> "1"). Rotulos fora vencem a opcao de legenda (os nomes ja estao nos rotulos);
' desligar os rotulos fora devolve a legenda que a opcao pede. Pie_ApplyLegend e o unico lugar que aplica essa regra.

Public Function Pie_BuildBlock(wb As Workbook, id As String, data As clsChartData, kind As String) As Range
    Dim blk As Range, body As Range, i As Long, n As Long, srcAddr As String, valAddr As String, valRange As String
    n = data.CategoryCount
    Set blk = CalcSheet_AllocateBlock(wb, id, kind, Mid$(ChartData_ExternalAddress(data.Source), 2), n + 1, 3)
    Set body = Block_Body(blk)
    body.Cells(1, 1).Value = "Categoria": body.Cells(1, 2).Value = "Valor": body.Cells(1, 3).Value = "Share"
    valRange = body.Cells(2, 2).Resize(n, 1).Address(True, True)       ' mesma planilha: sem nome de planilha
    For i = 1 To n
        body.Cells(i + 1, 1).Formula = ChartData_ExternalAddress(data.Source.Cells(i + 1, 1))
        srcAddr = Mid$(ChartData_ExternalAddress(data.Source.Cells(i + 1, 2)), 2)
        ' IF aninhado (e nao AND): AND avalia os dois lados e propagaria o #N/D de uma celula de erro para o Valor
        body.Cells(i + 1, 2).Formula = "=IF(ISNUMBER(" & srcAddr & "),IF(" & srcAddr & ">0," & srcAddr & ",0),0)"
        valAddr = body.Cells(i + 1, 2).Address(False, False)
        body.Cells(i + 1, 3).Formula = "=IF(SUM(" & valRange & ")>0," & valAddr & "/SUM(" & valRange & "),0)"
    Next
    Set Pie_BuildBlock = blk
End Function

Public Function Pie_Insert(ws As Worksheet, data As clsChartData, kind As String) As ChartObject
    Dim id As String, co As ChartObject, cht As Chart, s As Series, anchor As Range, blk As Range, body As Range
    Dim en As Long, es As String, ed As String
    If Kinds_Family(kind) <> "pie" Then Err.Raise 5, "Pie_Insert", "Tipo de pizza desconhecido: " & kind        ' antes de alocar o bloco
    On Error GoTo Fail
    id = Registry_NewId()
    Set blk = Pie_BuildBlock(ws.Parent, id, data, kind)
    Util_TestHookAfterBlock
    Set body = Block_Body(blk)
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, Chart_FreeTop(ws, anchor.Left, anchor.Top, 460, 280), 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    Set s = cht.SeriesCollection.NewSeries
    s.Name = ChartData_ExternalAddress(data.SeriesNameCell(1))
    s.Values = Block_Ref(body.Cells(2, 2).Resize(data.CategoryCount, 1))
    s.XValues = ChartData_ExternalAddress(data.CategoryCells())
    cht.ChartType = IIf(kind = "doughnut", xlDoughnut, xlPie)
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlZero
    Style_ApplyBase cht                              ' tolera pizza (sem eixos)
    Util_TestHookAfterChart
    cht.ChartGroups(1).FirstSliceAngle = 0
    If kind = "doughnut" Then cht.ChartGroups(1).DoughnutHoleSize = 50
    Registry_SetOption cht, "numfmt", "int"
    Registry_SetOption cht, "palette", "padrao"
    Registry_SetOption cht, "n", CStr(data.CategoryCount)
    Registry_SetOption cht, "legend", "right"
    If Kinds_Supports(kind, "pieout") Then Registry_SetOption cht, "pieout", "0"
    Style_ApplyPointColors cht, "padrao"
    Pie_ApplyLegend cht
    Pie_SetLabels cht, True
    Set Pie_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' A serie de dados (a unica) da pizza; Nothing se o grafico nao tem serie.
Private Function PieSeries(cht As Chart) As Series
    If cht.SeriesCollection.Count > 0 Then Set PieSeries = cht.SeriesCollection(1)
End Function

' Liga ou desliga os rotulos (nativos). Ligar aplica o conteudo pelas opcoes numfmt e pieout (religar mantem participacao e rotulos fora).
Public Sub Pie_SetLabels(cht As Chart, turnOn As Boolean)
    Dim s As Series
    Set s = PieSeries(cht)
    If s Is Nothing Then Exit Sub
    s.HasDataLabels = turnOn
    Registry_SetOption cht, "labels", IIf(turnOn, "1", "0")
    If turnOn Then Pie_ApplyLabelContent cht
End Sub

' Rotulos fora da pizza (pieout): OutsideEnd, nome da categoria e linhas de chamada, sem legenda (Pie_ApplyLegend). So na pizza: a rosca
' recusa a posicao do rotulo (erro 1004 do Excel) e aqui levanta 7002 sem gravar nada (Core_SetToggle mostra o aviso).
Public Sub Pie_SetOutside(cht As Chart, turnOn As Boolean)
    If Not Kinds_Supports(Registry_ChartType(cht), "pieout") Then
        Err.Raise vbObjectError + 7002, "Pie_SetOutside", "'Rótulos fora da pizza' não se aplica a este tipo de gráfico."
    End If
    Registry_SetOption cht, "pieout", IIf(turnOn, "1", "0")
    Pie_ApplyLabelContent cht
    Pie_ApplyLegend cht
End Sub

' Formato dos rotulos pela CHAVE (int, dec1, dec2, pct0, pct1, mil, mi): grava numfmt e reaplica o conteudo. pct0/pct1 mostram a
' participacao (ShowPercentage); os demais, o valor.
Public Sub Pie_ApplyNumberFormat(cht As Chart, key As String)
    Registry_SetOption cht, "numfmt", key
    Pie_ApplyLabelContent cht
End Sub

' Paleta: cores das fatias (Style_ApplyPointColors) e, depois, a cor da fonte dos rotulos de dentro por contraste com a fatia nova.
Public Sub Pie_ApplyColors(cht As Chart, paletteName As String)
    Style_ApplyPointColors cht, paletteName
    Pie_ApplyLabelContent cht
End Sub

' A legenda da pizza: (legend <> "none") And (pieout <> "1"). Aplica o estado e nao grava opcao nenhuma.
Public Sub Pie_ApplyLegend(cht As Chart)
    Dim want As Boolean
    want = (Registry_GetOption(cht, "legend", "right") <> "none") And (Registry_GetOption(cht, "pieout", "0") <> "1")
    If want Then
        If Not cht.HasLegend Then cht.HasLegend = True
        cht.Legend.Position = xlLegendPositionRight
        cht.Legend.Font.Size = 9
    Else
        If cht.HasLegend Then cht.HasLegend = False
    End If
End Sub

' Conteudo e aparencia dos rotulos de uma vez, pelas opcoes do registro (numfmt, pieout): valor ou participacao, formato, posicao (so na
' pizza), nome da categoria e linhas de chamada (fora), cor da fonte (fora: cinza-escuro; dentro: por contraste com a fatia). Sem rotulos
' ligados nao faz nada (o religar aplica). A rosca nunca recebe Position.
Private Sub Pie_ApplyLabelContent(cht As Chart)
    Dim s As Series, key As String, isPct As Boolean, outside As Boolean, isPie As Boolean, k As Long
    Set s = PieSeries(cht)
    If s Is Nothing Then Exit Sub
    If Not s.HasDataLabels Then Exit Sub
    key = Registry_GetOption(cht, "numfmt", "int")
    isPct = (key = "pct0" Or key = "pct1")
    isPie = Kinds_Supports(Registry_ChartType(cht), "pieout")      ' so a pizza escolhe a posicao
    outside = isPie And (Registry_GetOption(cht, "pieout", "0") = "1")
    With s.DataLabels
        .ShowSeriesName = False
        .ShowCategoryName = outside
        .ShowPercentage = isPct
        .ShowValue = Not isPct
        If outside Then .Separator = vbLf
        Style_AssignNumberFormat s.DataLabels, Style_NumberFormat(key)
        .Font.Size = 9
        If isPie Then .Position = IIf(outside, xlLabelPositionOutsideEnd, xlLabelPositionInsideEnd)
    End With
    If isPie Then s.HasLeaderLines = outside
    For k = 1 To s.Points.Count         ' (as atribuicoes acima, no nivel da serie, recriam os rotulos que o usuario tenha apagado)
        With s.Points(k).DataLabel.Font
            If outside Then
                .Color = RGB(64, 64, 64)
            Else
                .Color = Style_ContrastFont(CLng(s.Points(k).Format.Fill.ForeColor.RGB))
            End If
        End With
    Next
    If outside Then SyncZeroSliceNames cht, s
End Sub

' Com rotulos fora, a fatia de valor 0 (zero, negativo, texto ou vazio na origem) fica sem o nome da categoria: o rotulo vira um texto vazio
' (o formato de valor tem a secao de zero vazia). Sem isso os nomes das fatias que nao existem se amontoam uns sobre os outros na borda
' (medido: tres fatias zero na mesma posicao). Dentro nao precisa: o rotulo de valor zero ja sai vazio. Escreve so o que mudou.
Private Sub SyncZeroSliceNames(cht As Chart, s As Series)
    Dim blk As Range, body As Range, n As Long, k As Long, want As Boolean
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    For k = 1 To s.Points.Count
        want = True
        If k <= n Then want = (Util_ToDouble(body.Cells(k + 1, 2).Value) > 0)
        If s.Points(k).HasDataLabel Then         ' o rotulo que o usuario apagou fica apagado (escrever uma propriedade o recriaria)
            If s.Points(k).DataLabel.ShowCategoryName <> want Then s.Points(k).DataLabel.ShowCategoryName = want
        End If
    Next
End Sub

' O que na pizza depende dos DADOS e muda depois de inserir (Core_SafeRefresh): com rotulos fora, a fatia que passou a ter valor ganha o
' nome e a que passou a 0 o perde (SyncZeroSliceNames). Sem rotulos fora nao faz nada. Nunca propaga erro.
Public Sub Pie_RefreshAfterDataChange(cht As Chart)
    Dim s As Series
    On Error GoTo Fail
    If Not Kinds_Supports(Registry_ChartType(cht), "pieout") Then Exit Sub
    If Registry_GetOption(cht, "pieout", "0") <> "1" Then Exit Sub
    Set s = PieSeries(cht)
    If s Is Nothing Then Exit Sub
    If Not s.HasDataLabels Then Exit Sub
    SyncZeroSliceNames cht, s
    Exit Sub
Fail:
    Errors_Log "Pie_RefreshAfterDataChange" & vbTab & Err.Number & vbTab & Err.Description
End Sub
