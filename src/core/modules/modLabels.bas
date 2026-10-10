Attribute VB_Name = "modLabels"
Option Explicit

' Liga o rótulo de cada ponto à célula correspondente (rótulo "vivo") e posiciona pelo sinal.
Public Sub Labels_LinkPoints(s As Series, valueCells As Range, numFmt As String)
    Dim i As Long, n As Long, sheetRef As String, v As Variant
    n = valueCells.Rows.Count
    s.HasDataLabels = True
    sheetRef = "='" & Replace(valueCells.Worksheet.Name, "'", "''") & "'!"
    For i = 1 To n
        With s.Points(i)
            .HasDataLabel = True
            .DataLabel.Formula = sheetRef & valueCells.Cells(i, 1).Address(True, True)
            .DataLabel.NumberFormat = numFmt
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
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    If Registry_ChartType(cht) = "waterfall" Then
        Set s = Chart_FindSeries(cht, "tqq_lbl"): col = 17
    Else
        Set s = Chart_FindSeries(cht, "tqq_total"): col = 2
    End If
    If s Is Nothing Then Exit Sub
    If Not s.HasDataLabels Then Exit Sub
    For i = 1 To n
        If i <= s.Points.Count Then
            s.Points(i).DataLabel.Position = IIf(Util_ToDouble(body.Cells(i + 1, col).Value) < 0, xlLabelPositionBelow, xlLabelPositionAbove)
        End If
    Next
End Sub
