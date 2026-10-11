Attribute VB_Name = "modChartCombo"
Option Explicit
' Combinacao: colunas empilhadas (as series 1..S-1 da tabela) + uma linha com marcadores (a ultima coluna da tabela), com
' eixo secundario opcional para a linha (kind combo). O bloco e o da familia empilhada (Stacked_BuildBlock com kind combo:
' Total, TotalPos e VLine so das colunas) e totais, rotulos, conectores, cores e o numero de categorias vem dos Stacked_*.

' Exige pelo menos 2 colunas de valores (uma de colunas e a da linha). Erro de validacao dos dados (6008): Core_Insert mostra a
' mensagem como aviso, sem registrar falha; nada e criado. O eixo secundario nasce ligado quando as magnitudes diferem muito
' (Combo_DefaultSecondary). Em falha depois da insercao desfaz o grafico e o bloco.
Public Function Combo_Insert(ws As Worksheet, data As clsChartData) As ChartObject
    Dim co As ChartObject, id As String, en As Long, es As String, ed As String
    If data.SeriesCount < 2 Then
        Err.Raise vbObjectError + 6008, "Combo_Insert", "A combinação precisa de pelo menos 2 colunas de valores: as primeiras viram colunas empilhadas e a última vira linha."
    End If
    On Error GoTo Fail
    Set co = Stacked_Insert(ws, data, "combo")      ' desfaz o que criou em caso de falha
    Combo_SetSecondary co.Chart, Combo_DefaultSecondary(data)
    Set Combo_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then
        id = Mid$(co.Name, Len(ID_PREFIX) + 1)
        co.Delete
        If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    End If
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' Eixo secundario por padrao quando as magnitudes diferem por mais de 5 vezes (a linha de percentuais sobre colunas de
' milhares ficaria achatada no eixo comum): maior |valor| da ultima coluna (linha) x maior |soma das demais| por linha
' (colunas). Valores lidos por Util_ToDouble (texto/vazio = 0). Falso se alguma das magnitudes for zero.
Public Function Combo_DefaultSecondary(data As clsChartData) As Boolean
    Dim i As Long, j As Long, ns As Long, lineMax As Double, colMax As Double, v As Double, rowSum As Double
    ns = data.SeriesCount
    If ns < 2 Then Exit Function
    For i = 1 To data.CategoryCount
        v = Abs(Util_ToDouble(data.ValueCell(i, ns).Value))
        If v > lineMax Then lineMax = v
        rowSum = 0
        For j = 1 To ns - 1
            rowSum = rowSum + Util_ToDouble(data.ValueCell(i, j).Value)
        Next
        If Abs(rowSum) > colMax Then colMax = Abs(rowSum)
    Next
    If lineMax > 0 And colMax > 0 Then Combo_DefaultSecondary = (lineMax > 5 * colMax Or lineMax * 5 < colMax)
End Function

' Liga ou desliga o eixo secundario da linha. O eixo de valores secundario fica oculto como o primario (a escala ainda existe
' para a geometria). So em combinacao (nos demais tipos nao faz nada nem grava a opcao). A mudanca de grupo pode devolver as
' entradas das auxiliares a legenda: reescondidas no fim.
Public Sub Combo_SetSecondary(cht As Chart, turnOn As Boolean)
    Dim ln As Series
    If Not Kinds_Supports(Registry_ChartType(cht), "secondary") Then Exit Sub
    Set ln = Combo_LineSeries(cht)
    If ln Is Nothing Then Exit Sub
    If turnOn Then
        ln.AxisGroup = xlSecondary
        cht.HasAxis(xlValue, xlSecondary) = True
        Style_HideValueAxis cht.Axes(xlValue, xlSecondary)
    Else
        ln.AxisGroup = xlPrimary
    End If
    Chart_HideHelperLegendEntries cht
    Registry_SetOption cht, "secondary", IIf(turnOn, "1", "0")
End Sub

' A serie de linha da combinacao: a ultima serie nao auxiliar (Nothing se nao ha nenhuma).
Public Function Combo_LineSeries(cht As Chart) As Series
    Dim s As Series
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then Set Combo_LineSeries = s
    Next
End Function
