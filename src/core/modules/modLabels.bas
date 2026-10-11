Attribute VB_Name = "modLabels"
Option Explicit
Private mRefreshing As Boolean

' Liga o rótulo de cada ponto à célula correspondente (rótulo "vivo"). Posição: acima/abaixo conforme o sinal
' (fixedPos = 0) ou a posição fixa informada (ex.: xlLabelPositionCenter em segmentos empilhados).
' fontColor >= 0 define a cor da fonte (contraste com o preenchimento do segmento).
Public Sub Labels_LinkPoints(s As Series, valueCells As Range, numFmt As String, Optional fixedPos As Long = 0, Optional fontColor As Long = -1)
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
                .DataLabel.Position = IIf(Util_ToDouble(v) < 0, xlLabelPositionBelow, xlLabelPositionAbove)
            End If
        End With
    Next
End Sub

' Formato (código neutro de Style_NumberFormat) dos rótulos de segmento. Em 100% os segmentos mostram a
' PARTICIPAÇÃO (valor / total da linha): pct0 sempre, pct1 só quando o usuário escolhe pct1. Nos demais, numfmt.
Public Function Labels_SegmentFormat(cht As Chart) As String
    Dim key As String
    key = Registry_GetOption(cht, "numfmt", "int")
    If Registry_ChartType(cht) = "stacked100" Then key = IIf(key = "pct1", "pct1", "pct0")
    Labels_SegmentFormat = Style_NumberFormat(key)
End Function

' Formato dos rótulos de total. Em 100% os totais são valores absolutos: um formato percentual (pct0/pct1,
' que ali vale para as participações) não se aplica a eles e os totais ficam em "int".
Public Function Labels_TotalFormat(cht As Chart) As String
    Dim key As String
    key = Registry_GetOption(cht, "numfmt", "int")
    If Registry_ChartType(cht) = "stacked100" And (key = "pct0" Or key = "pct1") Then key = "int"
    Labels_TotalFormat = Style_NumberFormat(key)
End Function

' Reaplica acima/abaixo conforme o sinal atual das séries auxiliares de rótulo.
Public Sub Labels_RefreshPositions(cht As Chart)
    Dim blk As Range, body As Range, s As Series, col As Long, i As Long, n As Long, en As Long, ed As String, want As Long
    If mRefreshing Then Exit Sub
    mRefreshing = True
    On Error GoTo Done
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then GoTo Done
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    If Registry_ChartType(cht) = "waterfall" Then
        Set s = Chart_FindSeries(cht, "tqq_lbl"): col = 17   ' coluna 17 = WF_COL_LABELVAL (constante definida na Tarefa 9)
    Else
        Set s = Chart_FindSeries(cht, "tqq_total"): col = 3     ' acima/abaixo segue a posição plotada (TotalPos), não o líquido
    End If
    If s Is Nothing Then GoTo Done
    If Not s.HasDataLabels Then GoTo Done
    For i = 1 To n
        If i <= s.Points.Count Then
            want = IIf(Util_ToDouble(body.Cells(i + 1, col).Value) < 0, xlLabelPositionBelow, xlLabelPositionAbove)
            If s.Points(i).DataLabel.Position <> want Then s.Points(i).DataLabel.Position = want    ' so escreve se mudou (evita redesenho a cada calculo)
        End If
    Next
Done:
    If Err.Number <> 0 Then en = Err.Number: ed = Err.Description
    mRefreshing = False
    If en <> 0 Then Errors_Log "Labels_RefreshPositions" & vbTab & en & vbTab & ed
End Sub

' Rótulos ligados a células (Formula = célula da planilha _ThinQQ) exibem o TEXTO da célula, então o formato
' numérico precisa estar também na célula (gravado por NumberFormatLocal, como nos rótulos: o texto exibido segue os separadores do Excel).
Public Sub Labels_ApplyCellFormat(cht As Chart, fmt As String)
    Dim blk As Range, body As Range, col As Long, n As Long
    On Error GoTo Fail
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    If n < 1 Then Exit Sub
    If Registry_ChartType(cht) = "waterfall" Then col = WF_COL_LABELVAL Else col = 2
    body.Cells(2, col).Resize(n, 1).NumberFormatLocal = Style_LocalizeCode(fmt)
    ' 100%: colunas de participação (4..) seguem o formato dos segmentos, não o dos totais
    If Registry_ChartType(cht) = "stacked100" And body.Columns.Count > 3 Then
        body.Cells(2, 4).Resize(n, body.Columns.Count - 3).NumberFormatLocal = Style_LocalizeCode(Labels_SegmentFormat(cht))
    End If
    Exit Sub
Fail:
    Errors_Log "Labels_ApplyCellFormat" & vbTab & Err.Number & vbTab & Err.Description
End Sub
