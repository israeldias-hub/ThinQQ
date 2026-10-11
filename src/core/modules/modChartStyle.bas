Attribute VB_Name = "modChartStyle"
Option Explicit
' Aparência padrão "limpa": sem grade, eixo de valores invisível (mas presente, para a geometria),
' sem bordas, fonte 9 pt. Paletas de 6 cores.

Public Function Style_Palette(name As String) As Variant
    Select Case LCase$(name)
        Case "cinza"
            Style_Palette = Array(RGB(64, 64, 64), RGB(127, 127, 127), RGB(166, 166, 166), RGB(191, 191, 191), RGB(217, 217, 217), RGB(89, 89, 89))
        Case "tema"
            Style_Palette = Array(-1, -2, -3, -4, -5, -6)    ' accent1..accent6 do tema
        Case Else
            Style_Palette = Array(RGB(31, 56, 100), RGB(68, 114, 196), RGB(157, 195, 230), RGB(165, 165, 165), RGB(0, 150, 136), RGB(237, 162, 36))
    End Select
End Function

Public Sub Style_SetColor(cf As ColorFormat, colorValue As Variant)
    If colorValue < 0 Then
        cf.ObjectThemeColor = msoThemeColorAccent1 + (-colorValue - 1)
    Else
        cf.RGB = CLng(colorValue)
    End If
End Sub

Public Function Style_IsHelperSeries(s As Series) As Boolean
    Style_IsHelperSeries = (Left$(s.Name, Len(HELPER_PREFIX)) = HELPER_PREFIX)
End Function

' Branco para fundos escuros, cinza-escuro para claros (luminância relativa simples).
Public Function Style_ContrastFont(rgbValue As Long) As Long
    Dim r As Long, g As Long, b As Long, lum As Double
    r = rgbValue And &HFF: g = (rgbValue \ &H100) And &HFF: b = (rgbValue \ &H10000) And &HFF
    lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    If lum < 0.55 Then Style_ContrastFont = RGB(255, 255, 255) Else Style_ContrastFont = RGB(64, 64, 64)
End Function

Public Sub Style_ApplyBase(cht As Chart)
    With cht
        .HasTitle = True    ' liga e desliga: marca o título automático (gráfico de 1 série) como excluído
        .HasTitle = False
        .ChartArea.Format.Line.Visible = msoFalse
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Font.Size = 9
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse
        .HasAxis(xlValue) = True
        With .Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
            .MajorTickMark = xlTickMarkNone
            .MinorTickMark = xlTickMarkNone
            .TickLabelPosition = xlTickLabelPositionNone
            .Format.Line.Visible = msoFalse
        End With
        With .Axes(xlCategory)
            .MajorTickMark = xlTickMarkNone
            .MinorTickMark = xlTickMarkNone
            .TickLabelPosition = xlTickLabelPositionLow   ' rótulos de categoria abaixo do gráfico mesmo com valores negativos
            .Format.Line.ForeColor.RGB = RGB(166, 166, 166)
            .TickLabels.Font.Size = 9
            .TickLabels.Font.Color = RGB(64, 64, 64)
        End With
    End With
End Sub

Public Sub Style_ApplySeriesColors(cht As Chart, paletteName As String)
    Dim pal As Variant, s As Series, k As Long, c As Variant
    pal = Style_Palette(paletteName)
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            c = pal(k Mod 6)
            If s.ChartType = xlLine Or s.ChartType = xlLineMarkers Then
                Style_SetColor s.Format.Line.ForeColor, c
                s.Format.Line.Weight = 2
                s.MarkerStyle = xlMarkerStyleCircle
                s.MarkerSize = 5
                Style_SetColor s.Format.Fill.ForeColor, c
            Else
                Style_SetColor s.Format.Fill.ForeColor, c
                s.Format.Line.Visible = msoFalse
                If s.HasDataLabels Then If s.DataLabels.Position = xlLabelPositionCenter Then s.DataLabels.Font.Color = Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB))
            End If
            k = k + 1
        End If
    Next
End Sub

Public Function Style_NumberFormat(key As String) As String
    Select Case key
        Case "dec1": Style_NumberFormat = "#,##0.0;-#,##0.0;"
        Case "dec2": Style_NumberFormat = "#,##0.00;-#,##0.00;"
        Case "pct0": Style_NumberFormat = "0%;-0%;"
        Case "pct1": Style_NumberFormat = "0.0%;-0.0%;"
        Case "mil":  Style_NumberFormat = "#,##0,;-#,##0,;"
        Case "mi":   Style_NumberFormat = "#,##0.0,,;-#,##0.0,,;"
        Case Else:   Style_NumberFormat = "#,##0;-#,##0;"
    End Select
End Function

Public Sub Style_ApplyNumberFormat(cht As Chart, fmt As String)
    Dim s As Series, i As Long
    Labels_ApplyCellFormat cht, fmt
    For Each s In cht.SeriesCollection
        If s.HasDataLabels Then
            Style_AssignNumberFormat s.DataLabels, fmt
            ' rótulos ligados a células (por ponto) guardam formato próprio
            For i = 1 To s.Points.Count
                If s.Points(i).HasDataLabel Then Style_AssignNumberFormat s.Points(i).DataLabel, fmt
            Next
        End If
    Next
End Sub

' Tabela de formatos: os códigos em inglês (Style_NumberFormat) são a fonte neutra. Formatos de rótulos
' de gráfico são interpretados na notação do idioma do Excel neste ambiente, então os rótulos são sempre
' gravados por NumberFormatLocal com o código derivado por Style_LocalizeCode (nunca por NumberFormat).
' (A regra "códigos em inglês" do plano vale para formatos de célula/intervalo.)
Public Function Style_LocalizeCode(fmt As String) As String
    Dim d As String, t As String
    d = Application.International(xlDecimalSeparator)
    t = Application.International(xlThousandsSeparator)
    If Len(d) <> 1 Then d = "."
    If Len(t) <> 1 Then t = ","
    If t = d Then t = IIf(d = ",", ".", ",")
    ' NBSP (separador de milhar de alguns idiomas) é usado literalmente no código
    Style_LocalizeCode = Replace(Replace(Replace(fmt, ",", ChrW(1)), ".", d), ChrW(1), t)
End Function

' Atribui o formato (código em inglês de Style_NumberFormat) a um DataLabels/DataLabel, de forma determinística.
Public Sub Style_AssignNumberFormat(target As Object, fmt As String)
    target.NumberFormatLocal = Style_LocalizeCode(fmt)
End Sub
