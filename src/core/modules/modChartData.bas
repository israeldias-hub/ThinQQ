Attribute VB_Name = "modChartData"
Option Explicit

Public Function ChartData_FromRange(rng As Range) As clsChartData
    Dim src As Range, d As clsChartData
    Set src = rng
    If src.Cells.CountLarge = 1 Then
        Set src = src.CurrentRegion
    ElseIf src.Areas.Count = 1 Then
        ' Colunas/linhas inteiras: limita à região usada da planilha.
        Dim used As Range
        Set used = Application.Intersect(src, src.Worksheet.UsedRange)
        If used Is Nothing Then Set src = src.Cells(1, 1) Else Set src = used
    End If
    Set d = New clsChartData
    d.Init src
    Set ChartData_FromRange = d
End Function

' ='Nome da planilha'!$A$1:$B$2 (apóstrofos dobrados), aceito por Series.Values/Name/XValues e por Range.Formula.
Public Function ChartData_ExternalAddress(r As Range) As String
    ChartData_ExternalAddress = "='" & Replace(r.Worksheet.Name, "'", "''") & "'!" & r.Address(True, True)
End Function
