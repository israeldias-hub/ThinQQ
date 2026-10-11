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
    Util_TestHookAfterChart                          ' (teste) falha com o grafico ja criado: o Fail: abaixo o desfaz
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
    Combo_SyncSecondaryScale cht
End Sub

' Escala do eixo secundario (oculto): com todos os valores da linha >= 0 o minimo fica em 0 (o automatico do Excel escolheria um
' minimo perto do menor valor, ex.: 0,39 para 0,42..0,47, e exageraria a inclinacao de uma linha que o usuario nao ve medida);
' com algum valor negativo volta ao automatico. O 0 fixado aqui fica registrado na opcao "secmin": um minimo FIXO que nao e esse 0 foi
' definido a mao pelo usuario e nao e sobrescrito (devolvido ao automatico pelo usuario, volta a valer a regra). Só escreve quando precisa.
' Sem eixo secundario nao faz nada. Nunca propaga erro (chamada por eventos).
Public Sub Combo_SyncSecondaryScale(cht As Chart)
    Dim ln As Series, v As Variant, i As Long, nonNeg As Boolean, ax As Axis
    On Error GoTo Done
    If Registry_ChartType(cht) <> "combo" Then Exit Sub
    If Registry_GetOption(cht, "secondary", "0") <> "1" Then Exit Sub
    If Not cht.HasAxis(xlValue, xlSecondary) Then Exit Sub
    Set ln = Combo_LineSeries(cht)
    If ln Is Nothing Then Exit Sub
    nonNeg = True
    v = ln.Values
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            If Not IsError(v(i)) And Not IsEmpty(v(i)) Then
                If VarType(v(i)) <> vbString Then
                    If IsNumeric(v(i)) Then
                        If CDbl(v(i)) < 0 Then nonNeg = False
                    End If
                End If
            End If
        Next
    End If
    Set ax = cht.Axes(xlValue, xlSecondary)
    If Not ax.MinimumScaleIsAuto Then
        If Not (Registry_GetOption(cht, "secmin", "") = "0" And ax.MinimumScale = 0) Then GoTo Done      ' minimo do usuario: fica
    End If
    If nonNeg Then
        If ax.MinimumScaleIsAuto Then ax.MinimumScale = 0
        If Registry_GetOption(cht, "secmin", "") <> "0" Then Registry_SetOption cht, "secmin", "0"
    Else
        If Not ax.MinimumScaleIsAuto Then ax.MinimumScaleIsAuto = True
        If Len(Registry_GetOption(cht, "secmin", "")) > 0 Then Registry_SetOption cht, "secmin", ""
    End If
Done:
    If Err.Number <> 0 Then Errors_Log "Combo_SyncSecondaryScale" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' O que numa combinacao depende dos DADOS e muda depois de inserir (Core_SafeRefresh): a escala do eixo secundario e o formato
' dos rotulos da linha (que segue o formato da celula). Nunca propaga erro.
Public Sub Combo_RefreshAfterDataChange(cht As Chart)
    On Error Resume Next
    If Registry_ChartType(cht) <> "combo" Then Exit Sub
    Combo_SyncSecondaryScale cht
    Labels_RefreshComboLineFormat cht
End Sub

' A serie de linha da combinacao: a ultima serie nao auxiliar (Nothing se nao ha nenhuma).
Public Function Combo_LineSeries(cht As Chart) As Series
    Dim s As Series
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then Set Combo_LineSeries = s
    Next
End Function
