Attribute VB_Name = "modChartMekko"
Option Explicit
' Mekko (kinds mekko = eixo percentual, mekkounits = unidades): area empilhada em eixo de DATAS (design da Fase 2, secao 3.2).
' Cada categoria da tabela vira uma coluna de largura proporcional a W (percentual: o total da linha; unidades: a segunda coluna da
' tabela) desenhada por 5 pontos seguidos no eixo X inteiro 1..1001 (o eixo de datas trunca fracoes): L0 (x = borda esquerda, y = 0),
' L (borda esquerda, y = valor), C (centro, y = valor), R (borda direita, y = valor) e R0 (borda direita, y = 0). A descida a zero com
' largura nula, somada a borda branca das areas (1,5 pt), desenha o separador entre colunas encostadas (sem espaco, como no think-cell).
' Bloco (corpo): linha 1 = cabecalhos; 5n linhas; colunas (S = numero de series de dados):
'  1 X | 2 Categoria (so em C; vazia se a coluna tem largura 0) | 3 W (largura, so em C) | 4 Total (so em C) | 5..4+S Y_j (altura da faixa j:
'  participacao no percentual, valor nas unidades; 0 em L0/R0) | 5+S..4+2S Lbl_j (meio da faixa j, so em C; NA() nas demais) |
'  5+2S..4+3S LblVal_j (texto do rotulo: Y_j, ou 0 se a coluna tem largura 0) | 5+3S TotPos (topo: 1 ou Total) | 6+3S TotLbl (largura em %
'  ou Total) | 7+3S CatPos (0) | 8+3S VLine (unidades: media dos Totais ponderada por W; NA() no percentual).
' Rotulos: series auxiliares de LINHA invisivel no mesmo eixo de datas, com valor so nas linhas C. Um ponto NA() NAO esconde o rotulo
' (DisplayBlanksAs = xlZero: o Excel o poe no zero com o texto da celula; medido na Tarefa 4), entao Mekko_LinkLabels liga so os pontos
' C e apaga o rotulo dos demais um a um; nada escreve depois propriedades de rotulo no nivel da SERIE (escrever DataLabels.* recria os
' rotulos apagados): formato e cores sao aplicados por ponto.
Public Const MK_ROWS_PER_CAT As Long = 5
Public Const MK_COL_X As Long = 1
Public Const MK_COL_CAT As Long = 2
Public Const MK_COL_W As Long = 3
Public Const MK_COL_TOTAL As Long = 4
Private Const MK_X0 As Long = 1                  ' X da borda esquerda (o eixo de datas aceita MinimumScale = 1: medido)
Private Const MK_SPAN As Long = 1000             ' resolucao do eixo (0,1%): a borda direita e MK_X0 + MK_SPAN
Private Const MK_TOP_ROOM As Double = 24         ' folga minima (pt) acima da area interna: rotulos dos totais acima do topo (100%)
Private Const MK_BOTTOM_ROOM As Double = 21      ' folga minima (pt) abaixo da area interna: nomes das categorias (medido: o rotulo "abaixo"
' comeca 5,5 pt sob o zero e tem 14 pt; o layout automatico deixa ~13 pt e o Excel prende o rotulo na borda, colado ao eixo e 1 pt para fora)
Private Const MK_BORDER As String = "1.5"        ' espessura (pt) da borda branca entre colunas e faixas, gravada na opcao "border"

' ---------------------------------------------------------------------------------------------------------------------
' Colunas do bloco
' ---------------------------------------------------------------------------------------------------------------------

Public Function Mekko_SeriesCount(blk As Range) As Long
    Mekko_SeriesCount = (blk.Columns.Count - 8) \ 3
End Function

Public Function Mekko_YCol(j As Long) As Long
    Mekko_YCol = 4 + j
End Function

Public Function Mekko_LblCol(blk As Range, j As Long) As Long
    Mekko_LblCol = 4 + Mekko_SeriesCount(blk) + j
End Function

Public Function Mekko_LblValCol(blk As Range, j As Long) As Long
    Mekko_LblValCol = 4 + 2 * Mekko_SeriesCount(blk) + j
End Function

Public Function Mekko_TotPosCol(blk As Range) As Long
    Mekko_TotPosCol = 5 + 3 * Mekko_SeriesCount(blk)
End Function

Public Function Mekko_TotLblCol(blk As Range) As Long
    Mekko_TotLblCol = 6 + 3 * Mekko_SeriesCount(blk)
End Function

Public Function Mekko_CatPosCol(blk As Range) As Long
    Mekko_CatPosCol = 7 + 3 * Mekko_SeriesCount(blk)
End Function

Public Function Mekko_VLineCol(blk As Range) As Long
    Mekko_VLineCol = 8 + 3 * Mekko_SeriesCount(blk)
End Function

' Linha do corpo (contando o cabecalho como linha 1) onde fica o ponto C da categoria i e o ponto da serie, entre 1 e 5n: o ponto C e
' body.Cells(1 + Mekko_CenterRow(i), ...) e e o ponto Mekko_CenterRow(i) das series.
Public Function Mekko_CenterRow(i As Long) As Long
    Mekko_CenterRow = MK_ROWS_PER_CAT * (i - 1) + 3
End Function

' ---------------------------------------------------------------------------------------------------------------------
' Bloco
' ---------------------------------------------------------------------------------------------------------------------

' Mekko (unidades): a segunda coluna da tabela e a largura; sem nenhuma serie depois dela o grafico nao existe (erro de validacao dos
' dados, 6009: Core_Insert mostra a mensagem como aviso). Chamada antes de alocar qualquer coisa.
Private Sub Mekko_CheckData(data As clsChartData, kind As String)
    If kind = "mekkounits" And data.SeriesCount < 2 Then
        Err.Raise vbObjectError + 6009, "Mekko_Insert", "Mekko (unidades): a segunda coluna da tabela é a largura; são necessárias pelo menos 3 colunas."
    End If
End Sub

Public Function Mekko_BuildBlock(wb As Workbook, id As String, data As clsChartData, kind As String) As Range
    Dim blk As Range, body As Range, n As Long, ns As Long, m As Long, units As Boolean, i As Long, j As Long, rc As Long, r As Long
    Dim firstCol As Long, wAll As String, tAll As String, wt As String, prevSum As String, curSum As String, wAddr As String, totAddr As String
    Dim srcAddr As String, catAddr As String, xLeft As String, yAddr As String, rowCells As Range
    If Kinds_Family(kind) <> "mekko" Then Err.Raise 5, "Mekko_BuildBlock", "Tipo de Mekko desconhecido: " & kind
    Mekko_CheckData data, kind
    units = (kind = "mekkounits")
    n = data.CategoryCount
    firstCol = IIf(units, 2, 1)                      ' a serie j e a coluna firstCol + j da tabela (unidades: a 2a coluna e a largura)
    ns = data.SeriesCount - IIf(units, 1, 0)
    m = MK_ROWS_PER_CAT * n
    Set blk = CalcSheet_AllocateBlock(wb, id, kind, Mid$(ChartData_ExternalAddress(data.Source), 2), m + 1, 8 + 3 * ns)
    Set body = Block_Body(blk)
    body.Cells(1, MK_COL_X).Value = "X": body.Cells(1, MK_COL_CAT).Value = "Categoria"
    body.Cells(1, MK_COL_W).Value = "W": body.Cells(1, MK_COL_TOTAL).Value = "Total"
    For j = 1 To ns
        body.Cells(1, Mekko_YCol(j)).Value = "Y" & j
        body.Cells(1, Mekko_LblCol(blk, j)).Value = "Lbl" & j
        body.Cells(1, Mekko_LblValCol(blk, j)).Value = "LblVal" & j
    Next
    body.Cells(1, Mekko_TotPosCol(blk)).Value = "TotPos": body.Cells(1, Mekko_TotLblCol(blk)).Value = "TotLbl"
    body.Cells(1, Mekko_CatPosCol(blk)).Value = "CatPos": body.Cells(1, Mekko_VLineCol(blk)).Value = "VLine"
    ' fora das linhas C: NA() nas colunas de posicao/texto dos rotulos (de Lbl_1 a CatPos); Categoria, W e Total ficam vazias
    body.Cells(2, Mekko_LblCol(blk, 1)).Resize(m, Mekko_CatPosCol(blk) - Mekko_LblCol(blk, 1) + 1).Formula = "=NA()"
    ' W e Total so existem nas linhas C: a soma da coluna inteira e a soma das linhas C
    wAll = body.Cells(2, MK_COL_W).Resize(m, 1).Address(True, True)
    tAll = body.Cells(2, MK_COL_TOTAL).Resize(m, 1).Address(True, True)
    wt = "SUM(" & wAll & ")"
    For i = 1 To n
        rc = 1 + Mekko_CenterRow(i)                  ' linha C no corpo; L0 = rc - 2, L = rc - 1, R = rc + 1, R0 = rc + 2
        Set rowCells = data.Source.Cells(i + 1, firstCol + 1).Resize(1, ns)
        body.Cells(rc, MK_COL_TOTAL).Formula = "=AGGREGATE(9,6," & Mid$(ChartData_ExternalAddress(rowCells), 2) & ")"     ' ignora erros e texto
        totAddr = body.Cells(rc, MK_COL_TOTAL).Address(False, False)
        If units Then
            ' largura: numero positivo; texto, vazio, erro, zero e negativo valem 0 (IF aninhado: AND avaliaria os dois lados e propagaria #N/D)
            srcAddr = Mid$(ChartData_ExternalAddress(data.Source.Cells(i + 1, 2)), 2)
            body.Cells(rc, MK_COL_W).Formula = "=IF(ISNUMBER(" & srcAddr & "),IF(" & srcAddr & ">0," & srcAddr & ",0),0)"
        Else
            body.Cells(rc, MK_COL_W).Formula = "=MAX(0," & totAddr & ")"     ' total da linha; negativo vale 0 (sem coluna para tras)
        End If
        wAddr = body.Cells(rc, MK_COL_W).Address(False, False)
        catAddr = Mid$(ChartData_ExternalAddress(data.Source.Cells(i + 1, 1)), 2)
        body.Cells(rc, MK_COL_CAT).Formula = "=IF(" & wAddr & ">0," & catAddr & ","""")"     ' coluna de largura 0: sem nome solto sobre as vizinhas
        ' o rotulo ligado mostra o TEXTO da celula: o formato da categoria de origem (data, ano) vem junto (Range.NumberFormat: o codigo como esta)
        If Not IsNull(data.Source.Cells(i + 1, 1).NumberFormat) Then body.Cells(rc, MK_COL_CAT).NumberFormat = data.Source.Cells(i + 1, 1).NumberFormat
        ' X: bordas e centro pelas somas acumuladas das larguras (W so nas linhas C: a faixa do inicio do corpo ate a linha C anterior)
        curSum = "SUM(" & body.Cells(2, MK_COL_W).Address(True, True) & ":" & wAddr & ")"
        If i = 1 Then
            prevSum = "0"
            xLeft = "=" & MK_X0
        Else
            prevSum = "SUM(" & body.Cells(2, MK_COL_W).Address(True, True) & ":" & body.Cells(rc - MK_ROWS_PER_CAT, MK_COL_W).Address(False, False) & ")"
            xLeft = "=" & MK_X0 & "+ROUND(" & MK_SPAN & "*IF(" & wt & "=0,0," & prevSum & "/" & wt & "),0)"
        End If
        body.Cells(rc - 2, MK_COL_X).Formula = xLeft
        body.Cells(rc - 1, MK_COL_X).Formula = xLeft
        body.Cells(rc, MK_COL_X).Formula = "=" & MK_X0 & "+ROUND(" & MK_SPAN & "*IF(" & wt & "=0,0,(" & prevSum & "+" & curSum & ")/2/" & wt & "),0)"
        body.Cells(rc + 1, MK_COL_X).Formula = "=" & MK_X0 & "+ROUND(" & MK_SPAN & "*IF(" & wt & "=0,0," & curSum & "/" & wt & "),0)"
        body.Cells(rc + 2, MK_COL_X).Formula = "=" & MK_X0 & "+ROUND(" & MK_SPAN & "*IF(" & wt & "=0,0," & curSum & "/" & wt & "),0)"
        For j = 1 To ns
            srcAddr = Mid$(ChartData_ExternalAddress(data.Source.Cells(i + 1, firstCol + j)), 2)
            If units Then
                body.Cells(rc, Mekko_YCol(j)).Formula = "=IF(ISNUMBER(" & srcAddr & ")," & srcAddr & ",0)"
            Else
                ' participacao; linha de total zero: 0 (sem #DIV/0!)
                body.Cells(rc, Mekko_YCol(j)).Formula = "=IF(" & totAddr & "=0,0,IF(ISNUMBER(" & srcAddr & ")," & srcAddr & ",0)/" & totAddr & ")"
            End If
            yAddr = body.Cells(rc, Mekko_YCol(j)).Address(False, False)
            body.Cells(rc - 2, Mekko_YCol(j)).Formula = "=0"
            body.Cells(rc - 1, Mekko_YCol(j)).Formula = "=" & yAddr
            body.Cells(rc + 1, Mekko_YCol(j)).Formula = "=" & yAddr
            body.Cells(rc + 2, Mekko_YCol(j)).Formula = "=0"
            body.Cells(rc, Mekko_LblCol(blk, j)).Formula = "=SUM(" & body.Cells(rc, Mekko_YCol(1)).Address(False, False) & ":" & yAddr & ")-" & yAddr & "/2"
            body.Cells(rc, Mekko_LblValCol(blk, j)).Formula = "=IF(" & wAddr & ">0," & yAddr & ",0)"      ' coluna de largura 0: rotulo vazio (secao de zero)
        Next
        If units Then
            body.Cells(rc, Mekko_TotPosCol(blk)).Formula = "=" & totAddr
            body.Cells(rc, Mekko_TotLblCol(blk)).Formula = "=IF(" & wAddr & ">0," & totAddr & ",0)"
        Else
            body.Cells(rc, Mekko_TotPosCol(blk)).Formula = "=1"
            body.Cells(rc, Mekko_TotLblCol(blk)).Formula = "=IF(" & wt & "=0,0," & wAddr & "/" & wt & ")"     ' largura da coluna (% do total geral)
        End If
        body.Cells(rc, Mekko_CatPosCol(blk)).Formula = "=0"
    Next
    If units Then
        body.Cells(2, Mekko_VLineCol(blk)).Resize(m, 1).Formula = "=IF(" & wt & "=0,0,SUMPRODUCT(" & wAll & "," & tAll & ")/" & wt & ")"
    Else
        body.Cells(2, Mekko_VLineCol(blk)).Resize(m, 1).Formula = "=NA()"          ' percentual: sem linha de valor
    End If
    Set Mekko_BuildBlock = blk
End Function

' ---------------------------------------------------------------------------------------------------------------------
' Insercao
' ---------------------------------------------------------------------------------------------------------------------

Public Function Mekko_Insert(ws As Worksheet, data As clsChartData, kind As String) As ChartObject
    Dim id As String, co As ChartObject, cht As Chart, s As Series, j As Long, anchor As Range, blk As Range, body As Range
    Dim m As Long, ns As Long, xRef As String, units As Boolean
    Dim en As Long, es As String, ed As String
    If Kinds_Family(kind) <> "mekko" Then Err.Raise 5, "Mekko_Insert", "Tipo de Mekko desconhecido: " & kind     ' antes de alocar o bloco
    Mekko_CheckData data, kind                       ' recusa (6009) antes de alocar: nada a desfazer
    units = (kind = "mekkounits")
    On Error GoTo Fail
    id = Registry_NewId()
    Set blk = Mekko_BuildBlock(ws.Parent, id, data, kind)
    Util_TestHookAfterBlock
    Set body = Block_Body(blk)
    m = body.Rows.Count - 1
    ns = Mekko_SeriesCount(blk)
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, Chart_FreeTop(ws, anchor.Left, anchor.Top, 460, 280), 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    For j = 1 To ns
        Set s = cht.SeriesCollection.NewSeries
        s.Name = ChartData_ExternalAddress(data.SeriesNameCell(j + IIf(units, 1, 0)))
        s.Values = Block_Ref(body.Cells(2, Mekko_YCol(j)).Resize(m, 1))
    Next
    cht.ChartType = xlAreaStacked
    xRef = Block_Ref(body.Cells(2, MK_COL_X).Resize(m, 1))
    For j = 1 To ns
        cht.SeriesCollection(j).XValues = xRef       ' depois do tipo area: o eixo de datas le os X inteiros
    Next
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlZero
    Style_ApplyBase cht                              ' forca eixo de TEXTO (como nos outros tipos): o Mekko o troca abaixo
    Util_TestHookAfterChart
    Mekko_SetupAxes cht, kind
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionRight
    cht.Legend.Font.Size = 9
    Registry_SetOption cht, "numfmt", IIf(units, "int", "pct0")
    Registry_SetOption cht, "palette", "padrao"
    Registry_SetOption cht, "n", CStr(data.CategoryCount)
    Registry_SetOption cht, "legend", "right"
    Registry_SetOption cht, "border", MK_BORDER       ' separadores de 1,5 pt (texto com ponto: Style_ApplySeriesColors le a opcao)
    Mekko_ApplyColors cht, "padrao"
    Mekko_SetLabels cht, True
    Mekko_SetTotals cht, True
    Mekko_SetCategoryLabels cht
    If Kinds_Supports(kind, "ind100") Then Deco_SetInd100 cht, True        ' indicador 100%: ligado por padrao no Mekko (%)
    Mekko_EnsureRoom cht                             ' folgas de cima (totais) e de baixo (nomes)
    Ann_Refresh cht                                  ' o indicador foi posicionado antes da folga de cima (que desce o nivel 100%)
    Set Mekko_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' Eixo de categorias de DATAS com escala fixa MK_X0..MK_X0+MK_SPAN em dias, X = MK_X0 na borda esquerda da area de plotagem e
' MK_X0+MK_SPAN na direita (sem meia fatia), sem marcas nem rotulos, linha cinza. Eixo de valores (oculto por Style_ApplyBase): 0..1 no
' percentual, automatico nas unidades. Medido no Excel 2019: MinimumScale = 1 e aceito (lido de volta 1, sem automatico); BaseUnit le
' xlDays; AxisBetweenCategories = False e aceito sem erro mas e LIDO de volta como True e nao muda nada: no eixo de datas os pontos ficam
' nas posicoes exatas (centro do rotulo da 1a coluna a InsideLeft + InsideWidth * 0,2254 com largura 55/122, nas duas posicoes do
' eixo). Por isso Geom_ChartBetween devolve False em eixo de datas.
Private Sub Mekko_SetupAxes(cht As Chart, kind As String)
    With cht.Axes(xlCategory)
        .CategoryType = xlTimeScale
        .BaseUnitIsAuto = False
        .BaseUnit = xlDays
        .AxisBetweenCategories = False
        .MaximumScale = MK_X0 + MK_SPAN
        .MinimumScale = MK_X0
        .MajorUnit = MK_SPAN
        .MajorTickMark = xlTickMarkNone
        .MinorTickMark = xlTickMarkNone
        .TickLabelPosition = xlTickLabelPositionNone
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(166, 166, 166)
    End With
    With cht.Axes(xlValue)
        If kind = "mekkounits" Then
            .MinimumScaleIsAuto = True
            .MaximumScaleIsAuto = True
        Else
            .MinimumScale = 0
            .MaximumScale = 1
        End If
    End With
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Rotulos, totais e categorias
' ---------------------------------------------------------------------------------------------------------------------

' Liga o rotulo de cada ponto que tem posicao (valueCells nao e erro: so as linhas C) e texto (labelCells nao e erro) a celula de
' labelCells, com o formato (fmt = codigo de Style_NumberFormat; "" = sem formato, o texto da celula como esta), a posicao, a fonte
' de 9 pt na cor dada e o negrito; os demais pontos ficam SEM rotulo (um ponto NA() mostraria o rotulo no zero). Cada ponto e isolado:
' um erro e registrado e os demais continuam.
Public Sub Mekko_LinkLabels(s As Series, valueCells As Range, labelCells As Range, fmt As String, pos As Long, fontColor As Long, Optional bold As Boolean = False)
    Dim r As Long, m As Long, linkIt As Boolean
    m = valueCells.Rows.Count
    If Len(fmt) > 0 Then labelCells.NumberFormatLocal = Style_LocalizeCode(fmt)       ' o rotulo ligado mostra o TEXTO da celula
    s.HasDataLabels = True
    For r = 1 To m
        linkIt = Not IsError(valueCells.Cells(r, 1).Value)
        If linkIt Then linkIt = Not IsError(labelCells.Cells(r, 1).Value)
        On Error Resume Next
        If linkIt Then
            With s.Points(r)
                .HasDataLabel = True
                .DataLabel.Formula = Block_Ref(labelCells.Cells(r, 1))
                If Len(fmt) > 0 Then Style_AssignNumberFormat .DataLabel, fmt
                .DataLabel.Position = pos
                .DataLabel.Font.Size = 9
                .DataLabel.Font.Bold = bold
                .DataLabel.Font.Color = fontColor
            End With
        Else
            s.Points(r).HasDataLabel = False
        End If
        If Err.Number <> 0 Then Errors_Log "Mekko_LinkLabels" & vbTab & s.Name & " ponto " & r & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
        On Error GoTo 0
    Next
End Sub

' Rotulos de segmento: uma auxiliar tqq_lbl_j por serie, com os pontos em Lbl_j (meio da faixa, so em C) e o rotulo ligado a LblVal_j
' (participacao no percentual, valor nas unidades), centrado, fonte por contraste com a faixa. Formato dos segmentos
' (Labels_SegmentFormat: pct0/pct1 no percentual, numfmt nas unidades). Desligar apaga as auxiliares.
Public Sub Mekko_SetLabels(cht As Chart, turnOn As Boolean)
    Dim s As Series, blk As Range, body As Range, m As Long, ns As Long, j As Long, i As Long, fmt As String, posCells As Range
    If Not turnOn Then
        For i = cht.SeriesCollection.Count To 1 Step -1
            Set s = cht.SeriesCollection(i)
            If Left$(s.Name, 8) = "tqq_lbl_" Then s.Delete
        Next
        Registry_SetOption cht, "labels", "0"
        Exit Sub
    End If
    Set blk = Registry_Block(cht)
    Set body = Block_Body(blk)
    m = body.Rows.Count - 1
    ns = Mekko_SeriesCount(blk)
    fmt = Labels_SegmentFormat(cht)
    For j = 1 To ns
        Set posCells = body.Cells(2, Mekko_LblCol(blk, j)).Resize(m, 1)
        Set s = Chart_FindSeries(cht, "tqq_lbl_" & j)
        If s Is Nothing Then Set s = Chart_AddHelperSeries(cht, "tqq_lbl_" & j, posCells)
        Mekko_LinkLabels s, posCells, body.Cells(2, Mekko_LblValCol(blk, j)).Resize(m, 1), fmt, xlLabelPositionCenter, Style_ContrastFont(Style_DataSeriesColor(cht, j))
    Next
    Chart_HideHelperLegendEntries cht
    Registry_SetOption cht, "labels", "1"
End Sub

' Totais no topo de cada coluna: a auxiliar tqq_total em TotPos (1 no percentual; o Total nas unidades) com o rotulo ligado a TotLbl (a
' LARGURA da coluna em % do total geral no percentual; o Total nas unidades), acima, em negrito. Desligar apaga a auxiliar.
Public Sub Mekko_SetTotals(cht As Chart, turnOn As Boolean)
    Dim s As Series, blk As Range, body As Range, m As Long, posCells As Range
    Set s = Chart_FindSeries(cht, "tqq_total")
    If Not turnOn Then
        If Not s Is Nothing Then s.Delete
        Registry_SetOption cht, "totals", "0"
        Exit Sub
    End If
    Set blk = Registry_Block(cht)
    Set body = Block_Body(blk)
    m = body.Rows.Count - 1
    Set posCells = body.Cells(2, Mekko_TotPosCol(blk)).Resize(m, 1)
    If s Is Nothing Then Set s = Chart_AddHelperSeries(cht, "tqq_total", posCells)
    Mekko_LinkLabels s, posCells, body.Cells(2, Mekko_TotLblCol(blk)).Resize(m, 1), Mekko_TotalFormat(cht), xlLabelPositionAbove, RGB(64, 64, 64), True
    Chart_HideHelperLegendEntries cht
    Registry_SetOption cht, "totals", "1"
End Sub

' Nomes das categorias embaixo de cada coluna (o eixo de datas nao tem rotulos): a auxiliar tqq_cat em CatPos (y = 0) com o rotulo ligado
' a celula Categoria, abaixo, 9 pt cinza. Sempre presente (nao e uma opcao).
Private Sub Mekko_SetCategoryLabels(cht As Chart)
    Dim s As Series, blk As Range, body As Range, m As Long, posCells As Range
    Set blk = Registry_Block(cht)
    Set body = Block_Body(blk)
    m = body.Rows.Count - 1
    Set posCells = body.Cells(2, Mekko_CatPosCol(blk)).Resize(m, 1)
    Set s = Chart_FindSeries(cht, "tqq_cat")
    If s Is Nothing Then Set s = Chart_AddHelperSeries(cht, "tqq_cat", posCells)
    Mekko_LinkLabels s, posCells, body.Cells(2, MK_COL_CAT).Resize(m, 1), "", xlLabelPositionBelow, RGB(64, 64, 64)
    Chart_HideHelperLegendEntries cht
End Sub

' Formato (codigo de Style_NumberFormat) dos rotulos de total: no percentual a largura da coluna e sempre pct0 (numfmt vale para as
' participacoes); nas unidades, numfmt.
Public Function Mekko_TotalFormat(cht As Chart) As String
    If Registry_ChartType(cht) = "mekkounits" Then
        Mekko_TotalFormat = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    Else
        Mekko_TotalFormat = Style_NumberFormat("pct0")
    End If
End Function

' ---------------------------------------------------------------------------------------------------------------------
' Formato, cores, folga e linha de valor
' ---------------------------------------------------------------------------------------------------------------------

' Formato das celulas que sao texto de rotulo (LblVal_j: segmentos; TotLbl: totais) e, com a linha de valor ligada, da coluna VLine.
' Chamada por Labels_ApplyCellFormat (familia mekko). Nunca propaga erro.
Public Sub Mekko_ApplyCellFormat(cht As Chart)
    Dim blk As Range, body As Range, m As Long, j As Long, segFmt As String
    On Error GoTo Fail
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    Set body = Block_Body(blk)
    m = body.Rows.Count - 1
    If m < 1 Then Exit Sub
    segFmt = Style_LocalizeCode(Labels_SegmentFormat(cht))
    For j = 1 To Mekko_SeriesCount(blk)
        body.Cells(2, Mekko_LblValCol(blk, j)).Resize(m, 1).NumberFormatLocal = segFmt
    Next
    body.Cells(2, Mekko_TotLblCol(blk)).Resize(m, 1).NumberFormatLocal = Style_LocalizeCode(Mekko_TotalFormat(cht))
    If Registry_GetOption(cht, "vline", "0") = "1" Then Deco_ApplyVLineFormat cht
    Exit Sub
Fail:
    Errors_Log "Mekko_ApplyCellFormat" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Formato numerico (Style_ApplyNumberFormat, familia mekko): as celulas (Mekko_ApplyCellFormat) e, POR PONTO com rotulo, os rotulos
' de segmento e de total (escrever DataLabels.NumberFormatLocal na serie recriaria os rotulos dos pontos NA()).
Public Sub Mekko_ApplyNumberFormat(cht As Chart)
    Dim s As Series, segFmt As String, totFmt As String
    Labels_ApplyCellFormat cht, Mekko_TotalFormat(cht)         ' familia mekko: Mekko_ApplyCellFormat (o rotulo ligado mostra o texto da celula)
    segFmt = Labels_SegmentFormat(cht)
    totFmt = Mekko_TotalFormat(cht)
    For Each s In cht.SeriesCollection
        If Left$(s.Name, 8) = "tqq_lbl_" Then
            PointsFormat s, segFmt
        ElseIf s.Name = "tqq_total" Then
            PointsFormat s, totFmt
        End If
    Next
End Sub

Private Sub PointsFormat(s As Series, fmt As String)
    Dim r As Long
    For r = 1 To s.Points.Count
        If s.Points(r).HasDataLabel Then Style_AssignNumberFormat s.Points(r).DataLabel, fmt
    Next
End Sub

' Cores do Mekko: paleta nas areas com a borda branca de 1,5 pt (a opcao "border"; Style_ApplySeriesColors) e, depois, a cor da
' fonte de cada rotulo de segmento (por ponto) por contraste com a cor nova da faixa. ApplyOption "palette" chama esta (senao a troca
' de paleta perderia o contraste dos rotulos).
Public Sub Mekko_ApplyColors(cht As Chart, paletteName As String)
    Dim j As Long, ls As Series, r As Long, fc As Long, bw As Double
    bw = Val(Replace(Registry_GetOption(cht, "border", MK_BORDER), ",", "."))
    If bw <= 0 Then bw = Val(MK_BORDER)
    Style_ApplySeriesColors cht, paletteName, bw
    For j = 1 To Mekko_SeriesCount(Registry_Block(cht))
        Set ls = Chart_FindSeries(cht, "tqq_lbl_" & j)
        If Not ls Is Nothing Then
            fc = Style_ContrastFont(Style_DataSeriesColor(cht, j))
            For r = 1 To ls.Points.Count
                If ls.Points(r).HasDataLabel Then ls.Points(r).DataLabel.Font.Color = fc
            Next
        End If
    Next
End Sub

' Folga acima da area interna para os rotulos dos totais (acima do topo da coluna: no percentual, acima do nivel 100%): InsideTop >=
' MK_TOP_ROOM; e abaixo dela para os nomes das categorias (o eixo de datas nao tem rotulos e o layout automatico nao reserva o espaco):
' MK_BOTTOM_ROOM ate a borda do grafico. So aumenta (desce o topo e/ou reduz a altura, como Ann_EnsureRoom); 3 passadas porque a area
' interna nem sempre acompanha 1:1 a externa. Fora do Mekko nao faz nada. Chamada no Insert e por Chart_EnsureRoomAfterChange
' (Deco_SetLegend volta o layout ao automatico e a folga se perderia). Nunca propaga erro.
Public Sub Mekko_EnsureRoom(cht As Chart)
    Dim k As Long, delta As Double
    On Error Resume Next
    If Kinds_Family(Registry_ChartType(cht)) <> "mekko" Then Exit Sub
    For k = 1 To 3
        delta = MK_TOP_ROOM - cht.PlotArea.InsideTop
        If delta <= 0.01 Then Exit For
        cht.PlotArea.Height = cht.PlotArea.Height - delta
        cht.PlotArea.Top = cht.PlotArea.Top + delta
    Next
    For k = 1 To 3
        delta = MK_BOTTOM_ROOM - (cht.ChartArea.Height - (cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight))
        If delta <= 0.01 Then Exit For
        cht.PlotArea.Height = cht.PlotArea.Height - delta
    Next
    If Err.Number <> 0 Then Errors_Log "Mekko_EnsureRoom" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Numero de linhas da tabela de origem entre a primeira e a ultima categoria, lido nas formulas da coluna Categoria do bloco (linhas C).
' As series do Mekko apontam para o BLOCO (sempre 5n pontos), nao para a tabela: o tamanho da tabela so aparece nessas referencias.
' Linhas inseridas no meio da tabela afastam as duas pontas; uma linha apagada vira #REF!. -1 se nao for possivel ler (Core_ChartSourceMismatch
' trata como tabela mudada). Nunca propaga erro.
Public Function Mekko_SourceRowsSpanned(cht As Chart) As Long
    Dim blk As Range, body As Range, n As Long, r1 As Long, r2 As Long
    On Error GoTo Fail
    Mekko_SourceRowsSpanned = -1
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    Set body = Block_Body(blk)
    n = (body.Rows.Count - 1) \ MK_ROWS_PER_CAT
    If n < 1 Then Exit Function
    r1 = CategoryRefRow(CStr(body.Cells(1 + Mekko_CenterRow(1), MK_COL_CAT).Formula))
    r2 = CategoryRefRow(CStr(body.Cells(1 + Mekko_CenterRow(n), MK_COL_CAT).Formula))
    If r1 < 1 Or r2 < r1 Then Exit Function
    Mekko_SourceRowsSpanned = r2 - r1 + 1
    Exit Function
Fail:
    Mekko_SourceRowsSpanned = -1
End Function

' Linha da referencia a tabela na formula da Categoria (=IF(W>0,'Planilha'!$A$7,"")): o numero depois do ultimo "$" do trecho que segue o
' ultimo "!" (o nome da planilha pode ter "!" ou ","; o restante da formula nao tem). 0 se nao houver referencia (ex.: #REF!).
Private Function CategoryRefRow(f As String) As Long
    Dim p As Long, s As String
    If InStr(f, "#REF!") > 0 Then Exit Function
    p = InStrRev(f, "!")
    If p = 0 Then Exit Function
    s = Mid$(f, p + 1)
    p = InStr(s, ",")
    If p > 0 Then s = Left$(s, p - 1)
    p = InStrRev(s, "$")
    If p = 0 Then Exit Function
    CategoryRefRow = CLng(Val(Mid$(s, p + 1)))
End Function

' Celulas da coluna VLine (5n, uma por ponto) do bloco do Mekko; Nothing se o bloco nao tem a coluna (cabecalho) ou a leitura falha.
' Usada por Deco_VLineCells (linha de valor das unidades: a media dos Totais ponderada pela largura, de X = 1 a 1001).
Public Function Mekko_VLineCells(cht As Chart) As Range
    Dim blk As Range, body As Range, col As Long
    On Error GoTo Fail
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    If Kinds_Family(CStr(blk.Cells(1, 2).Value)) <> "mekko" Then Exit Function
    Set body = Block_Body(blk)
    col = Mekko_VLineCol(blk)
    If body.Rows.Count < 2 Or col < 1 Or col > body.Columns.Count Then Exit Function
    If CStr(body.Cells(1, col).Value) <> "VLine" Then Exit Function
    Set Mekko_VLineCells = body.Cells(2, col).Resize(body.Rows.Count - 1, 1)
    Exit Function
Fail:
    Set Mekko_VLineCells = Nothing
End Function
