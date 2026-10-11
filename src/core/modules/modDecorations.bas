Attribute VB_Name = "modDecorations"
Option Explicit
Private Const LEGEND_GAP As Double = 10         ' folga entre a area de plotagem e a legenda a direita (o layout automatico deixa ~11 pt)
Private Const LINEEND_OFFSET As Double = 6      ' o rotulo "Right" de um ponto comeca ~5,5 pt depois dele (medido nos totais XY)
Private Const LINEEND_MARGIN As Double = 4      ' respiro entre o fim do rotulo e a borda do grafico
Private Const LINEEND_WRAP_FRACTION As Double = 0.2     ' o Excel quebra em linhas um rotulo de dados mais largo que ~20% da largura do grafico
' (medido: grafico de 455 pt, "Regiao Sudeste 1.400.000,00" em 2 linhas; de 295 pt, em 3), entao o rotulo nunca ocupa mais que isso
Private Const MIN_PLOT_FRACTION As Double = 0.3         ' a area de plotagem nunca fica mais estreita que 30% da largura do grafico (nem <= 0)
Private Const VLINE_NAME As String = "tqq_vline"        ' serie auxiliar da linha de valor
Private Const IND100_NAME As String = "tqq_ann_ind100"  ' linha do indicador 100% (o rotulo e IND100_NAME & "_lbl"; ver modAnnotations)
Private Const IND_MIN_LEFT As Double = 46               ' folga minima (pt) a esquerda da area interna para o indicador 100% das colunas e areas
Private Const IND_MIN_TOP As Double = 20                ' folga minima (pt) acima da area interna para o indicador 100% das barras
Private Const VLINE_PREFIX As String = "Média "         ' texto antes do numero no rotulo da linha de valor (formato da celula VLine)
' Decoracoes do grafico (design da Fase 2, secao 6): a opcao de legenda (right, none, lineend), a linha de valor na media
' (Deco_SetValueLine), o indicador 100% (Deco_SetInd100) e a mesma escala entre graficos (Deco_SameScale / Deco_ResetScale /
' Deco_ApplyFixedScale).

' Opcao de legenda: right (a direita, 9 pt, sem as entradas das auxiliares tqq_*), none (sem legenda) e lineend (sem legenda e
' o nome de cada serie de linha no rotulo do ultimo ponto; so em linha e combinacao). Valida, aplica e SO ENTAO grava a opcao
' "legend": uma recusa (erro 7002, tratado por Core_SetOption como aviso) nao deixa a opcao gravada. right e none desfazem o
' nome no ultimo ponto. Depois de mudar a legenda a area de plotagem volta ao layout AUTOMATICO (o Excel so a reflui sozinho
' enquanto o layout dela e automatico: as folgas das setas, dos totais e do fim da linha a fixam em tamanho manual, e com ela
' fixa tirar a legenda nao devolvia o espaco) e as folgas, que so encolhem, sao refeitas por cima (Chart_EnsureRoomAfterChange:
' legenda a direita, totais, setas, fim da linha). Mesma opcao duas vezes da o mesmo resultado.
Public Sub Deco_SetLegend(cht As Chart, mode As String)
    Dim kind As String
    Select Case mode
        Case "right", "none", "lineend"
        Case Else: Err.Raise 5, "Deco_SetLegend", "Modo de legenda desconhecido: " & mode
    End Select
    kind = Registry_ChartType(cht)
    If mode = "lineend" Then
        If Not Kinds_Supports(kind, "lineend") Then Err.Raise vbObjectError + 7002, "Deco_SetLegend", "Rótulos no fim da linha só existem em gráficos de linha e combinação."
    End If
    If Not Kinds_SupportsChart(cht, "legend") Then Err.Raise vbObjectError + 7002, "Deco_SetLegend", "Legenda não se aplica a este tipo de gráfico."
    If Kinds_Family(kind) = "pie" Then
        ' Pizza e rosca: grava a opcao e aplica a regra unica (legend <> "none") And (pieout <> "1"): com rotulos fora a legenda continua oculta,
        ' qualquer que seja a opcao (os nomes ja estao nos rotulos). "lineend" ja foi recusado acima (so linha e combinacao).
        Registry_SetOption cht, "legend", mode
        Pie_ApplyLegend cht
        Exit Sub
    End If
    Select Case mode
        Case "right"
            LineEnd cht, False
            cht.HasLegend = True
            cht.Legend.Position = xlLegendPositionRight
            cht.Legend.Font.Size = 9
            Chart_HideHelperLegendEntries cht
        Case "none"
            LineEnd cht, False
            cht.HasLegend = False
        Case Else
            cht.HasLegend = False
            LineEnd cht, True
    End Select
    Registry_SetOption cht, "legend", mode
    On Error Resume Next
    cht.PlotArea.Position = xlChartElementPositionAutomatic     ' layout automatico: com a legenda nova o Excel recalcula a area interna
    If Err.Number <> 0 Then Errors_Log "Deco_SetLegend/PlotArea.Position" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    On Error GoTo 0
    Chart_EnsureRoomAfterChange cht
End Sub

' Reaplica o nome no fim da linha quando a opcao e lineend (chamada no fim de Stacked_SetLabels: ligar/desligar os rotulos
' refaz os rotulos da serie e apagaria o nome do ultimo ponto). Nao faz nada com outra opcao.
Public Sub Deco_ApplyLineEnd(cht As Chart)
    If Registry_GetOption(cht, "legend", "") = "lineend" Then LineEnd cht, True
End Sub

' Em cada serie de linha nao auxiliar, o ultimo ponto. Ligado: rotulo com o nome da serie (separador " ", a direita do
' ponto) e, se os rotulos estao ligados (opcao labels), o valor depois do nome. Desligado: volta ao rotulo normal do ponto
' (so o valor, acima) ou, sem rotulos, sem rotulo.
Private Sub LineEnd(cht As Chart, turnOn As Boolean)
    Dim s As Series, n As Long, showVal As Boolean
    showVal = (Registry_GetOption(cht, "labels", "1") = "1")
    For Each s In cht.SeriesCollection
        If IsLineSeries(s) Then
            n = s.Points.Count
            If n >= 1 Then
                If turnOn Then
                    s.Points(n).HasDataLabel = True
                    With s.Points(n).DataLabel
                        .ShowSeriesName = True
                        .ShowValue = showVal
                        .ShowCategoryName = False
                        .Separator = " "
                        .Position = xlLabelPositionRight
                        .Font.Size = 9
                        .Font.Color = RGB(64, 64, 64)
                        If Labels_IsSourceFormatSeries(cht, s) Then      ' combinacao: fundo branco, como os rotulos da linha
                            .Format.Fill.Visible = msoTrue
                            .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
                        End If
                        If showVal Then
                            If Labels_IsSourceFormatSeries(cht, s) Then Labels_ApplyLineSourceFormat s.Points(n).DataLabel, cht Else Style_AssignNumberFormat s.Points(n).DataLabel, Labels_SegmentFormat(cht)
                        End If
                    End With
                ElseIf showVal Then
                    s.Points(n).HasDataLabel = True
                    With s.Points(n).DataLabel
                        .ShowSeriesName = False
                        .ShowValue = True
                        .Position = xlLabelPositionAbove
                    End With
                Else
                    s.Points(n).HasDataLabel = False
                End If
            End If
        End If
    Next
    If turnOn Then LineEndRoom cht
End Sub

' Refaz a folga do rotulo do fim da linha depois de uma mudanca posterior (formato numerico mais largo, grafico redimensionado);
' so com a opcao lineend. Ver LineEndRoom.
Public Sub Deco_EnsureLineEndRoom(cht As Chart)
    If Registry_GetOption(cht, "legend", "") = "lineend" Then LineEndRoom cht
End Sub

' A legenda ligada (ou arrastada) ou o grafico redimensionado depois que a area de plotagem ganhou tamanho manual (a folga das
' setas e dos totais a fixa) nao a empurra: com a legenda a direita a area de plotagem nao pode avancar sobre ela. So reduz a
' largura (nunca abaixo de MIN_PLOT_FRACTION da largura do grafico). Chamada por Chart_EnsureRoomAfterChange. Nunca propaga erro.
Public Sub Deco_FitPlotToLegend(cht As Chart)
    Dim k As Long, delta As Double
    On Error Resume Next
    If Not cht.HasLegend Then Exit Sub
    If cht.Legend.Position <> xlLegendPositionRight Then Exit Sub
    For k = 1 To 3
        delta = (cht.PlotArea.Left + cht.PlotArea.Width) - (cht.Legend.Left - LEGEND_GAP)
        If delta <= 0.5 Then Exit For
        If Not Deco_ShrinkPlotWidth(cht, delta) Then Exit For
    Next
    If Err.Number <> 0 Then Errors_Log "Deco_FitPlotToLegend" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Reduz a largura da area de plotagem em delta pt sem deixa-la menor que MIN_PLOT_FRACTION da largura do grafico (rede de seguranca
' para graficos minusculos: o Excel recusa largura <= 0 e a area ficaria sob a legenda); False quando nao ha mais o que reduzir
' (ja no minimo). Publica para o teste direto do minimo.
Public Function Deco_ShrinkPlotWidth(cht As Chart, delta As Double) As Boolean
    Dim w As Double, minW As Double, nw As Double
    w = cht.PlotArea.Width
    minW = MIN_PLOT_FRACTION * cht.ChartArea.Width
    If w <= minW + 0.01 Then Exit Function
    nw = w - delta
    If nw < minW Then nw = minW
    cht.PlotArea.Width = nw
    Deco_ShrinkPlotWidth = True
End Function

' O rotulo "nome valor" do ultimo ponto fica a direita dele: sem a legenda a area de plotagem se expande e o ultimo ponto
' fica a meia fatia da borda, pouco para um nome comprido. Reduz a largura da area de plotagem ate o maior rotulo (largura
' estimada pelo texto, Labels_EstimateWidth) caber antes da borda do grafico. O rotulo que o Excel quebra em linhas (mais largo que
' ~20% do grafico) ocupa so essa fracao, e a area de plotagem nunca fica abaixo de MIN_PLOT_FRACTION. Nunca propaga erro.
Private Sub LineEndRoom(cht As Chart)
    Dim s As Series, best As Double, w As Double, need As Double, k As Long, n As Long, f As Double, xLast As Double, delta As Double
    On Error Resume Next
    For Each s In cht.SeriesCollection
        If IsLineSeries(s) Then
            w = 0
            w = Labels_EstimateWidth(s.Points(s.Points.Count).DataLabel.Text)
            If w > best Then best = w
        End If
    Next
    n = Geom_ChartCategoryCount(cht)
    If best <= 0 Or n < 1 Then Exit Sub
    If best > LINEEND_WRAP_FRACTION * cht.ChartArea.Width Then best = LINEEND_WRAP_FRACTION * cht.ChartArea.Width
    need = LINEEND_OFFSET + best + LINEEND_MARGIN
    f = 1                                                        ' fracao da largura interna onde fica o ultimo ponto
    If Geom_ChartBetween(cht) Then f = (n - 0.5) / n
    For k = 1 To 3
        xLast = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * f
        delta = need - (cht.ChartArea.Width - xLast)
        If delta <= 0.5 Then Exit For
        If Not Deco_ShrinkPlotWidth(cht, delta / f) Then Exit For
    Next
    If Err.Number <> 0 Then Errors_Log "Deco_SetLegend/LineEndRoom" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Linha de valor (media dos totais)
' ---------------------------------------------------------------------------------------------------------------------

' Celulas da coluna VLine (uma por categoria; no Mekko, uma por ponto: 5 por categoria) e da coluna Pos do bloco do grafico, por
' familia: stacked e combo -> Stacked_VLineCol/PosCol, area -> Area_VLineCol/PosCol, mekko -> Mekko_VLineCells (sem Pos: o Mekko e
' vertical). Valida o CABECALHO da coluna: um bloco da
' Fase 1 (colunas; nunca publicado) nao tem Pos/VLine e "as ultimas colunas" cairiam em outra coluna (Part, TotalPos). Nothing para
' familia sem a coluna, bloco antigo ou leitura que falha. Nunca propaga erro.
Public Function Deco_VLineCells(cht As Chart) As Range
    Set Deco_VLineCells = BlockColumnCells(cht, "VLine")
End Function

Public Function Deco_PosCells(cht As Chart) As Range
    Set Deco_PosCells = BlockColumnCells(cht, "Pos")
End Function

Private Function BlockColumnCells(cht As Chart, header As String) As Range
    Dim blk As Range, body As Range, col As Long
    On Error GoTo Fail
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    Select Case Kinds_Family(Registry_ChartType(cht))
        Case "stacked", "combo"
            If header = "VLine" Then col = Stacked_VLineCol(blk) Else col = Stacked_PosCol(blk)
        Case "area"
            If header = "VLine" Then col = Area_VLineCol(blk) Else col = Area_PosCol(blk)
        Case "mekko"
            If header = "VLine" Then Set BlockColumnCells = Mekko_VLineCells(cht)      ' valida o cabecalho
            Exit Function
        Case Else
            Exit Function
    End Select
    Set body = Block_Body(blk)
    If body.Rows.Count < 2 Or col < 1 Then Exit Function
    If CStr(body.Cells(1, col).Value) <> header Then Exit Function
    Set BlockColumnCells = body.Cells(2, col).Resize(body.Rows.Count - 1, 1)
    Exit Function
Fail:
    Set BlockColumnCells = Nothing
End Function

' Codigo de formato (neutro, em ingles) da celula VLine: "Média " antes do numero no formato do grafico (numfmt), em TODAS as secoes
' (a media negativa e a nula tambem mostram o texto; a secao de zero dos formatos dos rotulos e vazia e esconderia "Média 0").
' O rotulo da linha de valor e ligado a esta celula e mostra o texto dela.
Private Function VLineFormatCode(cht As Chart) As String
    Dim parts() As String, q As String
    q = """" & VLINE_PREFIX & """"
    parts = Split(Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int")) & ";;", ";")
    VLineFormatCode = q & parts(0) & ";" & q & parts(1) & ";" & q & parts(0)
End Function

' Grava na coluna VLine do bloco o formato da media (VLineFormatCode, localizado). Chamada ao ligar a linha de valor e quando o
' formato numerico do grafico muda (Labels_ApplyCellFormat). Sem a coluna (bloco antigo) nao faz nada. Nunca propaga erro.
Public Sub Deco_ApplyVLineFormat(cht As Chart)
    Dim vcells As Range
    On Error Resume Next
    If Not Kinds_Supports(Registry_ChartType(cht), "vline") Then Exit Sub        ' 100%: a coluna e NA() e nao ha linha
    Set vcells = Deco_VLineCells(cht)
    If vcells Is Nothing Then Exit Sub
    vcells.NumberFormatLocal = Style_LocalizeCode(VLineFormatCode(cht))
    If Err.Number <> 0 Then Errors_Log "Deco_ApplyVLineFormat" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Linha de valor na media dos totais (opcao vline). Ligar: refaz a serie auxiliar tqq_vline (colunas, linha, area e combinacao: linha
' sobre a coluna VLine do bloco; barras: XY no grupo secundario, X = VLine e Y = Pos, do centro da 1a faixa ao da ultima), tracejada
' cinza de 1 pt, sem marcadores, fora da legenda, com o rotulo "Média N" (ligado a celula do ultimo ponto) a direita (colunas) ou acima
' (barras). A media e FORMULA do bloco (=AVERAGE(Total)): muda sozinha com os dados. Recusa (erro 7002) nos tipos sem o recurso (100%,
' cascata, Mekko (%), pizza, Gantt: Kinds_Supports "vline") e em bloco da Fase 1 (sem a coluna VLine). Nos tipos 100% a coluna VLine
' e NA() e NAO ha linha nem rotulo: um ponto NA() nao esconderia o rotulo (DisplayBlanksAs = xlZero o poe no zero com o texto da celula).
' Desligar: apaga a serie. Quem chama atualiza o grafico (Core_SafeRefresh).
Public Sub Deco_SetValueLine(cht As Chart, turnOn As Boolean)
    Dim kind As String, s As Series, vcells As Range, pcells As Range, horiz As Boolean, m As Long
    Dim en As Long, es As String, ed As String
    kind = Registry_ChartType(cht)
    If Not turnOn Then
        Set s = Chart_FindSeries(cht, VLINE_NAME)
        If Not s Is Nothing Then s.Delete
        If Kinds_Supports(kind, "vline") Then Registry_SetOption cht, "vline", "0"
        If Not s Is Nothing Then ReturnRoom cht                  ' a folga do rotulo (so reduz) volta: layout automatico e as demais folgas refeitas
        Exit Sub
    End If
    If Not Kinds_Supports(kind, "vline") Then Err.Raise vbObjectError + 7002, "Deco_SetValueLine", "A linha de valor não está disponível para este tipo de gráfico."
    horiz = Kinds_IsHorizontal(kind)
    Set vcells = Deco_VLineCells(cht)
    If horiz And Not vcells Is Nothing Then Set pcells = Deco_PosCells(cht)
    If vcells Is Nothing Then
        Err.Raise vbObjectError + 7002, "Deco_SetValueLine", "Este gráfico foi criado numa versão anterior e não tem os dados da linha de valor. Use 'Alterar dados' para reconstruí-lo."
    ElseIf horiz And pcells Is Nothing Then
        Err.Raise vbObjectError + 7002, "Deco_SetValueLine", "Este gráfico foi criado numa versão anterior e não tem os dados da linha de valor. Use 'Alterar dados' para reconstruí-lo."
    End If
    On Error GoTo Fail
    Set s = Chart_FindSeries(cht, VLINE_NAME)
    If Not s Is Nothing Then s.Delete: Set s = Nothing           ' religar refaz (idempotente)
    Deco_ApplyVLineFormat cht
    If horiz Then Set s = Chart_AddHelperSeries(cht, VLINE_NAME, vcells, pcells) Else Set s = Chart_AddHelperSeries(cht, VLINE_NAME, vcells)
    With s.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(127, 127, 127)
        .Weight = 1
        .DashStyle = msoLineDash
    End With
    s.MarkerStyle = xlMarkerStyleNone
    m = vcells.Rows.Count                                        ' nao o numero de categorias: no Mekko sao 5 pontos por categoria
    s.Points(m).HasDataLabel = True
    With s.Points(m).DataLabel
        .Formula = Block_Ref(vcells.Cells(m, 1))
        .Position = IIf(horiz, xlLabelPositionAbove, xlLabelPositionRight)
        .Font.Size = 9
        .Font.Color = RGB(89, 89, 89)
        .Format.Fill.Visible = msoTrue                           ' fundo branco: o rotulo cai sobre a ultima coluna/barra (legivel, medido no PNG)
        .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    Chart_HideHelperLegendEntries cht
    Registry_SetOption cht, "vline", "1"
    VLineRoom cht
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not s Is Nothing Then s.Delete
    On Error GoTo 0
    Err.Raise en, es, ed
End Sub

' O rotulo "Média N" do ultimo ponto fica a direita dele (colunas; o ultimo ponto esta a meia fatia da borda da area de plotagem,
' pouco para "Média 1.234.567") ou acima do ponto, centrado na linha (barras). Reduz a largura da area de plotagem ate o rotulo
' (largura estimada pelo texto; o Excel quebra em linhas o que passa de ~20% do grafico) caber antes da borda livre (a do grafico
' ou, com a legenda a direita, a dela). So reduz. Nunca propaga erro.
Private Sub VLineRoom(cht As Chart)
    Dim s As Series, m As Long, w As Double, need As Double, k As Long, f As Double, n As Long, delta As Double, xp As Double
    Dim mn As Double, mx As Double, avg As Double, vcells As Range
    On Error Resume Next
    Set s = Chart_FindSeries(cht, VLINE_NAME)
    If s Is Nothing Then Exit Sub
    m = s.Points.Count
    If m < 1 Then Exit Sub
    w = Labels_EstimateWidth(s.Points(m).DataLabel.Text)
    If w <= 0 Then Exit Sub
    If w > LINEEND_WRAP_FRACTION * cht.ChartArea.Width Then w = LINEEND_WRAP_FRACTION * cht.ChartArea.Width
    If Geom_ChartIsHorizontal(cht) Then
        ' centrado em x = media: meia largura de cada lado
        Set vcells = Deco_VLineCells(cht)
        If vcells Is Nothing Then Exit Sub
        avg = Util_ToDouble(vcells.Cells(1, 1).Value)
        mn = cht.Axes(xlValue, xlPrimary).MinimumScale: mx = cht.Axes(xlValue, xlPrimary).MaximumScale
        If mx <= mn Then Exit Sub
        f = (avg - mn) / (mx - mn)
        If f < 0.2 Then f = 0.2
        If f > 1 Then f = 1
        need = w / 2 + LINEEND_MARGIN
        For k = 1 To 3
            xp = Geom_ChartValCoord(cht, avg)
            delta = need - (Geom_ChartRightEdge(cht) - xp)
            If delta <= 0.5 Then Exit For
            If Not Deco_ShrinkPlotWidth(cht, delta / f) Then Exit For
        Next
    Else
        n = Geom_ChartCategoryCount(cht)
        If n < 1 Then Exit Sub
        need = LINEEND_OFFSET + w + LINEEND_MARGIN
        f = 1                                                    ' fracao da largura interna onde fica o ultimo ponto
        If Geom_ChartBetween(cht) Then f = (n - 0.5) / n
        For k = 1 To 3
            xp = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * f
            delta = need - (Geom_ChartRightEdge(cht) - xp)
            If delta <= 0.5 Then Exit For
            If Not Deco_ShrinkPlotWidth(cht, delta / f) Then Exit For
        Next
    End If
    If Err.Number <> 0 Then Errors_Log "Deco_SetValueLine/VLineRoom" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Indicador 100%
' ---------------------------------------------------------------------------------------------------------------------

' Liga/desliga o indicador "100%" (opcao ind100): um tique no nivel 100% do eixo e o rotulo, esquerda nas colunas e areas, em cima nas
' barras (formas tqq_ann_ind100*, posicionadas por Ann_Refresh). Ligado por padrao nos graficos 100% (os Insert chamam). Recusa (erro
' 7002) fora dos tipos 100% e Mekko (%). Ligar abre a folga fora da area de plotagem (Deco_EnsureIndicatorRoom). Ann_RemoveAll nao o remove.
Public Sub Deco_SetInd100(cht As Chart, turnOn As Boolean)
    Dim kind As String, en As Long, es As String, ed As String, had As Boolean
    kind = Registry_ChartType(cht)
    If Not turnOn Then
        had = Ann_HasInd100(cht)
        Ann_RemoveInd100 cht
        If Kinds_Supports(kind, "ind100") Then Registry_SetOption cht, "ind100", "0"
        If had Then ReturnRoom cht                               ' a folga do indicador (so reduz) volta: layout automatico e as demais folgas refeitas
        Exit Sub
    End If
    If Not Kinds_Supports(kind, "ind100") Then Err.Raise vbObjectError + 7002, "Deco_SetInd100", "O indicador 100% só existe em gráficos 100% e Mekko (%)."
    On Error GoTo Fail
    Ann_AddInd100 cht
    Deco_EnsureIndicatorRoom cht
    Registry_SetOption cht, "ind100", "1"
    Ann_Refresh cht
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    Ann_RemoveInd100 cht
    On Error GoTo 0
    Err.Raise en, es, ed
End Sub

' As folgas (Chart_EnsureRoomAfterChange) so REDUZEM a area de plotagem: ao desligar uma decoracao (indicador 100%, linha de valor) ou voltar a
' escala ao automatico, o espaco reservado nao voltaria sozinho. Como Deco_SetLegend: o layout da area de plotagem volta ao AUTOMATICO (descarta
' um layout manual do usuario) e as folgas das decoracoes e opcoes que continuam ligadas sao refeitas por cima; as setas sao reposicionadas.
' Nunca propaga erro.
Private Sub ReturnRoom(cht As Chart)
    On Error Resume Next
    cht.PlotArea.Position = xlChartElementPositionAutomatic
    If Err.Number <> 0 Then Errors_Log "Deco/ReturnRoom.Position" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Chart_EnsureRoomAfterChange cht
    Ann_Refresh cht
    If Err.Number <> 0 Then Errors_Log "Deco/ReturnRoom" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Folga para o indicador 100% fora da area interna: colunas e areas, a esquerda (o rotulo "100%" a 10 pt do tique: no minimo 46 pt
' ate a borda esquerda da area interna, ou a largura do rotulo + 16 se maior); barras, em cima (no minimo 20 pt, ou 12 + a altura do
' rotulo). So AUMENTA a folga (nunca a diminui): encolhe a area de plotagem pelo lado. Chamada ao ligar o indicador e por
' Chart_EnsureRoomAfterChange (legenda, formato, redimensionamento descartam o layout manual). Nunca propaga erro.
Public Sub Deco_EnsureIndicatorRoom(cht As Chart)
    Dim lbl As Shape, need As Double, k As Long, delta As Double, w0 As Double, h0 As Double, shrunk As Double
    On Error Resume Next
    Set lbl = cht.Shapes(IND100_NAME & "_lbl")
    If Err.Number <> 0 Then Err.Clear: Set lbl = Nothing
    If Geom_ChartIsHorizontal(cht) Then
        need = IND_MIN_TOP
        If Not lbl Is Nothing Then
            If 12 + lbl.Height > need Then need = 12 + lbl.Height
        End If
        For k = 1 To 3
            delta = need - cht.PlotArea.InsideTop
            If delta <= 0.5 Then Exit For
            h0 = cht.PlotArea.Height
            If h0 - delta < MIN_PLOT_FRACTION * cht.ChartArea.Height Then Exit For
            cht.PlotArea.Height = h0 - delta
            cht.PlotArea.Top = cht.PlotArea.Top + delta
        Next
    Else
        need = IND_MIN_LEFT
        If Not lbl Is Nothing Then
            If 16 + lbl.Width > need Then need = 16 + lbl.Width
        End If
        For k = 1 To 3
            delta = need - cht.PlotArea.InsideLeft
            If delta <= 0.5 Then Exit For
            w0 = cht.PlotArea.Width
            If Not Deco_ShrinkPlotWidth(cht, delta) Then Exit For
            shrunk = w0 - cht.PlotArea.Width                     ' o minimo de largura pode ter limitado a reducao
            cht.PlotArea.Left = cht.PlotArea.Left + shrunk
        Next
    End If
    If Err.Number <> 0 Then Errors_Log "Deco_EnsureIndicatorRoom" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' As folgas das decoracoes ligadas (indicador 100% e rotulo da linha de valor); chamada por Chart_EnsureRoomAfterChange, que e o que
' Deco_SetLegend (volta o layout ao automatico), o formato numerico e o redimensionamento executam: folga registrada so na criacao
' se perderia. Sem decoracao ligada nao faz nada. Nunca propaga erro.
Public Sub Deco_EnsureDecorationRoom(cht As Chart)
    On Error Resume Next
    If Registry_GetOption(cht, "ind100", "0") = "1" Then Deco_EnsureIndicatorRoom cht
    If Err.Number <> 0 Then Errors_Log "Deco_EnsureDecorationRoom/ind100" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    If Registry_GetOption(cht, "vline", "0") = "1" Then VLineRoom cht
    If Err.Number <> 0 Then Errors_Log "Deco_EnsureDecorationRoom/vline" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Mesma escala
' ---------------------------------------------------------------------------------------------------------------------

' Texto neutro de uma escala para a opcao (vmin/vmax): ponto decimal em qualquer localidade (Str$ nunca usa a virgula; Val le de volta).
Public Function Deco_ScaleText(v As Double) As String
    Deco_ScaleText = Trim$(Str$(v))
End Function

' Mesma escala (design da Fase 2, secao 6): para cada grafico com eixo de valores livre (Kinds_Supports "samescale": nao os 100%,
' Mekko (%), pizza e Gantt) devolve o eixo ao automatico, le o minimo e o maximo e aplica a TODOS a uniao (o menor minimo e o maior
' maximo), persistindo em vmin/vmax para o refresh reaplicar (Deco_ApplyFixedScale). Chamar de novo recalcula a partir das escalas
' automaticas de agora (a escala e fixa, nao dinamica). Devolve quantos receberam a escala; skipped = quantos foram ignorados pelo
' TIPO. Com menos de dois graficos aceitos nada e feito (uma escala "igual a de si mesma" so fixaria o eixo) e devolve 0. Um grafico
' cuja leitura falha e registrado e fica de fora. Nunca propaga erro por grafico.
Public Function Deco_SameScale(charts As Collection, ByRef skipped As Long) As Long
    Dim c As Variant, cht As Chart, ok As Collection, got As Collection, los() As Double, his() As Double, n As Long, i As Long
    Dim lo As Double, hi As Double, gLo As Double, gHi As Double
    skipped = 0
    If charts Is Nothing Then Exit Function
    Set ok = New Collection
    For Each c In charts
        Set cht = c
        Select Case FreeScaleKind(cht)
            Case 1: ok.Add cht
            Case 0: skipped = skipped + 1
        End Select                                                ' -1 (grafico ilegivel, registrado): fica de fora sem contar como "tipo"
    Next
    If ok.Count < 2 Then Exit Function
    ReDim los(1 To ok.Count): ReDim his(1 To ok.Count)
    Set got = New Collection
    For Each c In ok
        Set cht = c
        If ReadAutoScale(cht, lo, hi) Then
            n = n + 1
            los(n) = lo: his(n) = hi
            got.Add cht
        End If
    Next
    If n < 2 Then                                                 ' so um grafico legivel: nao fixa a escala dele nela mesma
        For Each c In got
            Set cht = c
            Deco_ApplyFixedScale cht                              ' a leitura o devolveu ao automatico: volta ao que estava gravado
        Next
        Exit Function
    End If
    gLo = los(1): gHi = his(1)
    For i = 2 To n
        If los(i) < gLo Then gLo = los(i)
        If his(i) > gHi Then gHi = his(i)
    Next
    For Each c In got
        Set cht = c
        If SetScale(cht, gLo, gHi) Then
            Registry_SetOption cht, "vmin", Deco_ScaleText(gLo)
            Registry_SetOption cht, "vmax", Deco_ScaleText(gHi)
            Deco_SameScale = Deco_SameScale + 1
            Deco_AfterScaleChange cht
        End If
    Next
End Function

' 1 = o grafico tem escala livre (Kinds_Supports "samescale"), 0 = nao tem (tipo), -1 = nao foi possivel ler o tipo (grafico invalido; registrado).
Private Function FreeScaleKind(cht As Chart) As Long
    On Error GoTo Fail
    FreeScaleKind = IIf(Kinds_Supports(Registry_ChartType(cht), "samescale"), 1, 0)
    Exit Function
Fail:
    Errors_Log "Deco_SameScale/tipo" & vbTab & Err.Number & vbTab & Err.Description
    FreeScaleKind = -1
End Function

' Devolve o eixo de valores ao automatico e le a escala que o Excel escolheu (ler DEPOIS de ligar o automatico: o Excel recalcula
' a escala na leitura). False (registrado) se a leitura falha.
Private Function ReadAutoScale(cht As Chart, ByRef lo As Double, ByRef hi As Double) As Boolean
    Dim ax As Axis
    On Error GoTo Fail
    Util_TestHookScaleRead cht
    Set ax = cht.Axes(xlValue, xlPrimary)
    ax.MinimumScaleIsAuto = True
    ax.MaximumScaleIsAuto = True
    lo = ax.MinimumScale
    hi = ax.MaximumScale
    ReadAutoScale = True
    Exit Function
Fail:
    Errors_Log "Deco_SameScale/ReadAutoScale" & vbTab & Err.Number & vbTab & Err.Description
End Function

' Fixa o minimo e o maximo do eixo de valores primario. A ordem importa quando o novo intervalo nao se sobrepoe ao atual (o Excel
' recusa minimo acima do maximo fixo): se uma ordem falha tenta a outra. False (registrado) se as duas falham.
Private Function SetScale(cht As Chart, lo As Double, hi As Double) As Boolean
    Dim ax As Axis
    On Error Resume Next
    Set ax = cht.Axes(xlValue, xlPrimary)
    ax.MinimumScale = lo
    ax.MaximumScale = hi
    If Err.Number <> 0 Then
        Err.Clear
        ax.MaximumScale = hi
        ax.MinimumScale = lo
    End If
    If Err.Number <> 0 Then
        Errors_Log "Deco_SameScale/SetScale" & vbTab & Err.Number & vbTab & Err.Description
        Exit Function
    End If
    SetScale = True
End Function

' O que muda quando a escala do eixo de valores muda por aqui: a auxiliar XY das barras (eixo X secundario, se o Excel criou um) segue
' o eixo primario, as folgas (setas com o maximo fixo, totais, indicador) sao refeitas e rotulos e setas reposicionados. Nunca propaga erro.
Public Sub Deco_AfterScaleChange(cht As Chart)
    On Error Resume Next
    If Geom_ChartIsHorizontal(cht) Then Chart_SyncSecondaryX cht
    Chart_EnsureRoomAfterChange cht
    Labels_RefreshPositions cht
    Ann_Refresh cht
    If Err.Number <> 0 Then Errors_Log "Deco_AfterScaleChange" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Volta o eixo de valores ao automatico e limpa vmin/vmax. Tipo sem escala livre: so limpa as opcoes (nunca ha o que desfazer).
' Nunca propaga erro.
Public Sub Deco_ResetScale(cht As Chart)
    Dim ax As Axis
    On Error Resume Next
    Registry_SetOption cht, "vmin", ""
    Registry_SetOption cht, "vmax", ""
    If Kinds_Supports(Registry_ChartType(cht), "samescale") Then
        Set ax = cht.Axes(xlValue, xlPrimary)
        ax.MinimumScaleIsAuto = True
        ax.MaximumScaleIsAuto = True
        cht.PlotArea.Position = xlChartElementPositionAutomatic   ' a folga das setas com o maximo fixo (so reduz) volta; refeita para o eixo automatico abaixo
        Deco_AfterScaleChange cht
    End If
    If Err.Number <> 0 Then Errors_Log "Deco_ResetScale" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Reaplica a escala fixa gravada (vmin/vmax) quando o eixo mudou (o refresh depois de uma mudanca de dados: o automatico recalcula).
' Sem a opcao, em tipo sem escala livre ou com texto invalido: nao faz nada. So escreve quando o eixo e automatico ou difere do gravado
' (com tolerancia: o texto da opcao tem 15 digitos). Mudou: refaz as folgas e o alinhamento (Deco_AfterScaleChange). Nunca propaga erro.
Public Sub Deco_ApplyFixedScale(cht As Chart)
    Dim sMin As String, sMax As String, lo As Double, hi As Double, ax As Axis, need As Boolean
    On Error GoTo Fail
    sMin = Registry_GetOption(cht, "vmin", "")
    If Len(sMin) = 0 Then Exit Sub
    sMax = Registry_GetOption(cht, "vmax", "")
    If Len(sMax) = 0 Then Exit Sub
    If Not Kinds_Supports(Registry_ChartType(cht), "samescale") Then Exit Sub
    lo = Val(sMin): hi = Val(sMax)
    If hi <= lo Then Exit Sub
    Set ax = cht.Axes(xlValue, xlPrimary)
    If ax.MinimumScaleIsAuto Or Abs(ax.MinimumScale - lo) > 0.000000001 * (1 + Abs(lo)) Then need = True
    If ax.MaximumScaleIsAuto Or Abs(ax.MaximumScale - hi) > 0.000000001 * (1 + Abs(hi)) Then need = True
    If Not need Then Exit Sub
    If SetScale(cht, lo, hi) Then Deco_AfterScaleChange cht
    Exit Sub
Fail:
    Errors_Log "Deco_ApplyFixedScale" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' True quando o grafico tem escala fixa (Mesma escala) e os dados passaram do maximo gravado (vmax): as barras ficariam cortadas. O maior valor
' plotado e o topo da pilha (Ann_GetCategoryData: empilhadas, area, cascata) ou, em agrupadas e linha, o maior valor de qualquer serie. False em
' qualquer duvida. Usada por Core_SafeRefresh para o aviso unico.
Public Function Deco_ScaleClipped(cht As Chart) As Boolean
    Dim sMax As String, kind As String, values() As Double, tops() As Double, i As Long, s As Series, v As Variant, mx As Double
    On Error GoTo Fail
    sMax = Registry_GetOption(cht, "vmax", "")
    kind = Registry_ChartType(cht)
    If Len(sMax) = 0 Or Not Kinds_Supports(kind, "samescale") Then Exit Function
    If kind = "clustered" Or kind = "barclustered" Or kind = "line" Then
        For Each s In cht.SeriesCollection
            If Not Style_IsHelperSeries(s) Then
                v = s.Values
                For i = LBound(v) To UBound(v)
                    If Util_ToDouble(v(i)) > mx Then mx = Util_ToDouble(v(i))
                Next
            End If
        Next
    Else
        Ann_GetCategoryData cht, values, tops
        For i = 1 To UBound(tops)
            If tops(i) > mx Then mx = tops(i)
        Next
    End If
    Deco_ScaleClipped = (mx > Val(sMax) + 0.000001)
Fail:
End Function

Private Function IsLineSeries(s As Series) As Boolean
    If Style_IsHelperSeries(s) Then Exit Function
    IsLineSeries = (s.ChartType = xlLine Or s.ChartType = xlLineMarkers)
End Function
