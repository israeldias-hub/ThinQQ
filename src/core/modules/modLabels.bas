Attribute VB_Name = "modLabels"
Option Explicit
Private mRefreshing As Boolean

' Liga o rótulo de cada ponto à célula correspondente (rótulo "vivo"). Posição: acima/abaixo conforme o sinal
' (fixedPos = 0) ou a posição fixa informada (ex.: xlLabelPositionCenter em segmentos empilhados). Com horizontal
' (auxiliares XY de barras), a posição por sinal é à direita (>= 0) ou à esquerda (< 0) do ponto.
' fontColor >= 0 define a cor da fonte (contraste com o preenchimento do segmento).
Public Sub Labels_LinkPoints(s As Series, valueCells As Range, numFmt As String, Optional fixedPos As Long = 0, Optional fontColor As Long = -1, Optional horizontal As Boolean = False)
    Dim i As Long, n As Long, sheetRef As String, v As Variant
    n = valueCells.Rows.Count
    s.HasDataLabels = True
    valueCells.NumberFormatLocal = Style_LocalizeCode(numFmt)
    sheetRef = "='" & Replace(valueCells.Worksheet.Name, "'", "''") & "'!"
    For i = 1 To n
        With s.Points(i)
            .HasDataLabel = True
            .DataLabel.Formula = sheetRef & valueCells.Cells(i, 1).Address(True, True)
            Style_AssignNumberFormat .DataLabel, numFmt
            .DataLabel.Font.Size = 9
            If fontColor >= 0 Then .DataLabel.Font.Color = fontColor Else .DataLabel.Font.Color = RGB(64, 64, 64)
            If fixedPos <> 0 Then
                .DataLabel.Position = fixedPos
            Else
                v = valueCells.Cells(i, 1).Value
                If horizontal Then
                    .DataLabel.Position = IIf(Util_ToDouble(v) < 0, xlLabelPositionLeft, xlLabelPositionRight)
                Else
                    .DataLabel.Position = IIf(Util_ToDouble(v) < 0, xlLabelPositionBelow, xlLabelPositionAbove)
                End If
            End If
        End With
    Next
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
    Dim s As Series, i As Long, w As Double, best As Double, t As String
    Set s = Chart_FindSeries(cht, "tqq_total")
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
    Labels_MaxTotalWidth = best
End Function

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
Public Sub Labels_RefreshPositions(cht As Chart)
    Dim blk As Range, body As Range, s As Series, col As Long, i As Long, n As Long, en As Long, ed As String, want As Long
    Dim kind As String, horiz As Boolean
    If mRefreshing Then Exit Sub
    mRefreshing = True
    On Error GoTo Done
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then GoTo Done
    kind = CStr(blk.Cells(1, 2).Value)
    horiz = Kinds_IsHorizontal(kind)
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    If Kinds_Family(kind) = "waterfall" Then
        Set s = Chart_FindSeries(cht, "tqq_lbl"): col = WF_COL_LABELVAL
    Else
        Set s = Chart_FindSeries(cht, "tqq_total"): col = 3     ' acima/abaixo segue a posição plotada (TotalPos), não o líquido
    End If
    If s Is Nothing Then GoTo Done
    If Not s.HasDataLabels Then GoTo Done
    For i = 1 To n
        If i <= s.Points.Count Then
            If horiz Then
                want = IIf(Util_ToDouble(body.Cells(i + 1, col).Value) < 0, xlLabelPositionLeft, xlLabelPositionRight)
            Else
                want = IIf(Util_ToDouble(body.Cells(i + 1, col).Value) < 0, xlLabelPositionBelow, xlLabelPositionAbove)
            End If
            On Error Resume Next
            If s.Points(i).DataLabel.Position <> want Then s.Points(i).DataLabel.Position = want    ' só escreve se mudou (evita redesenho a cada cálculo)
            Err.Clear
            On Error GoTo Done
        End If
    Next
Done:
    If Err.Number <> 0 Then en = Err.Number: ed = Err.Description
    mRefreshing = False
    If en <> 0 Then Errors_Log "Labels_RefreshPositions" & vbTab & en & vbTab & ed
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
    Dim blk As Range, body As Range, col As Long, n As Long, kind As String, nPart As Long
    On Error GoTo Fail
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    kind = CStr(blk.Cells(1, 2).Value)
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    If n < 1 Then Exit Sub
    If Kinds_Family(kind) = "waterfall" Then col = WF_COL_LABELVAL Else col = 2
    body.Cells(2, col).Resize(n, 1).NumberFormatLocal = Style_LocalizeCode(fmt)
    ' 100%: as colunas de participação Part_j (a partir da 4) seguem o formato dos segmentos, não o dos totais. Depois
    ' delas vêm Pos e VLine (que não são rótulos): em blocos com Pos/VLine são Columns.Count - 5 colunas; os cabeçalhos
    ' "Part" decidem, o que também serve a blocos da Fase 1 (sem Pos/VLine).
    If Kinds_Family(kind) = "stacked" And Kinds_Is100(kind) Then
        nPart = PartColumnCount(body)
        If nPart > 0 Then body.Cells(2, 4).Resize(n, nPart).NumberFormatLocal = Style_LocalizeCode(Labels_SegmentFormat(cht))
    End If
    Exit Sub
Fail:
    Errors_Log "Labels_ApplyCellFormat" & vbTab & Err.Number & vbTab & Err.Description
End Sub
