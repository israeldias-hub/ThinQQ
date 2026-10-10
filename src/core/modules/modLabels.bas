Attribute VB_Name = "modLabels"
Option Explicit
Private mRefreshing As Boolean

' Liga o rótulo de cada ponto à célula correspondente (rótulo "vivo") e posiciona pelo sinal.
Public Sub Labels_LinkPoints(s As Series, valueCells As Range, numFmt As String)
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
            .DataLabel.Font.Color = RGB(64, 64, 64)
            v = valueCells.Cells(i, 1).Value
            .DataLabel.Position = IIf(Util_ToDouble(v) < 0, xlLabelPositionBelow, xlLabelPositionAbove)
        End With
    Next
End Sub

' Reaplica acima/abaixo conforme o sinal atual das séries auxiliares de rótulo.
Public Sub Labels_RefreshPositions(cht As Chart)
    Dim blk As Range, body As Range, s As Series, col As Long, i As Long, n As Long
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
        Set s = Chart_FindSeries(cht, "tqq_total"): col = 2
    End If
    If s Is Nothing Then GoTo Done
    If Not s.HasDataLabels Then GoTo Done
    For i = 1 To n
        If i <= s.Points.Count Then
            s.Points(i).DataLabel.Position = IIf(Util_ToDouble(body.Cells(i + 1, col).Value) < 0, xlLabelPositionBelow, xlLabelPositionAbove)
        End If
    Next
Done:
    mRefreshing = False
End Sub

' Rótulos ligados a células (Formula = célula da planilha _ThinQQ) exibem o TEXTO da célula, então o formato
' numérico precisa estar também na célula (gravado por NumberFormatLocal, como nos rótulos: o texto exibido segue os separadores do Excel).
Public Sub Labels_ApplyCellFormat(cht As Chart, fmt As String)
    Dim blk As Range, body As Range, col As Long, n As Long
    On Error Resume Next
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    If n < 1 Then Exit Sub
    If Registry_ChartType(cht) = "waterfall" Then col = WF_COL_LABELVAL Else col = 2
    body.Cells(2, col).Resize(n, 1).NumberFormatLocal = Style_LocalizeCode(fmt)
End Sub
