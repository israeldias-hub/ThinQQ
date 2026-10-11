Attribute VB_Name = "modDecorations"
Option Explicit
Private Const LEGEND_GAP As Double = 10         ' folga entre a area de plotagem e a legenda a direita (o layout automatico deixa ~11 pt)
Private Const LINEEND_OFFSET As Double = 6      ' o rotulo "Right" de um ponto comeca ~5,5 pt depois dele (medido nos totais XY)
Private Const LINEEND_MARGIN As Double = 4      ' respiro entre o fim do rotulo e a borda do grafico
Private Const LINEEND_WRAP_FRACTION As Double = 0.2     ' o Excel quebra em linhas um rotulo de dados mais largo que ~20% da largura do grafico
' (medido: grafico de 455 pt, "Regiao Sudeste 1.400.000,00" em 2 linhas; de 295 pt, em 3), entao o rotulo nunca ocupa mais que isso
Private Const MIN_PLOT_FRACTION As Double = 0.3         ' a area de plotagem nunca fica mais estreita que 30% da largura do grafico (nem <= 0)
' Decoracoes do grafico (design da Fase 2, secao 6). Nesta tarefa: a opcao de legenda (right, none, lineend). As demais
' (linha de valor, indicador 100%, mesma escala) chegam na Tarefa 7.

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

Private Function IsLineSeries(s As Series) As Boolean
    If Style_IsHelperSeries(s) Then Exit Function
    IsLineSeries = (s.ChartType = xlLine Or s.ChartType = xlLineMarkers)
End Function
