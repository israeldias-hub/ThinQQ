Attribute VB_Name = "modChartWaterfall"
Option Explicit
' Bloco (corpo): linha 1 = cabeçalhos; 18 colunas:
'  1 Cat | 2 Src | 3 IsTotal | 4 Delta | 5 Start | 6 End | 7 Lo | 8 Hi | 9 Base | 10 PiecePos | 11 PieceNeg
' 12 UpPos | 13 UpNeg | 14 DownPos | 15 DownNeg | 16 Total | 17 LabelVal | 18 LabelPos
Public Const WF_COL_CAT As Long = 1
Public Const WF_COL_END As Long = 6
Public Const WF_COL_BASE As Long = 9
Public Const WF_COL_UPPOS As Long = 12
Public Const WF_COL_UPNEG As Long = 13
Public Const WF_COL_DOWNPOS As Long = 14
Public Const WF_COL_DOWNNEG As Long = 15
Public Const WF_COL_TOTAL As Long = 16
Public Const WF_COL_LABELVAL As Long = 17
Public Const WF_COL_LABELPOS As Long = 18
' Unica fonte da lista de marcadores de total (minusculas, sem espacos): usada pela funcao VBA e pela formula do bloco.
Public Const WF_MARKERS As String = "e,t,=,total,subtotal"
' Nota: total na primeira linha tem Start = 0 e End = 0 (nada desenhado); o calculo "para tras" (estilo
' think-cell, total define o ponto de partida) esta fora de escopo.

Public Function Waterfall_IsTotalMarker(v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If VarType(v) <> vbString Then Exit Function
    Waterfall_IsTotalMarker = (InStr("," & WF_MARKERS & ",", "," & LCase$(Trim$(CStr(v))) & ",") > 0) And Len(Trim$(CStr(v))) > 0
End Function

Public Function Waterfall_BuildBlock(wb As Workbook, id As String, data As clsChartData) As Range
    Dim blk As Range, body As Range, n As Long, i As Long, hdr As Variant, c As Long
    n = data.CategoryCount
    Set blk = CalcSheet_AllocateBlock(wb, id, "waterfall", Mid$(ChartData_ExternalAddress(data.Source), 2), n + 1, 18)
    Set body = Block_Body(blk)
    hdr = Array("Cat", "Src", "IsTotal", "Delta", "Start", "End", "Lo", "Hi", "Base", "PiecePos", "PieceNeg", _
                "UpPos", "UpNeg", "DownPos", "DownNeg", "Total", "LabelVal", "LabelPos")
    For c = 0 To 17
        body.Cells(1, c + 1).Value = hdr(c)
    Next
    For i = 1 To n
        body.Cells(i + 1, 1).Formula = ChartData_ExternalAddress(data.Source.Cells(i + 1, 1))
        body.Cells(i + 1, 2).Formula = ChartData_ExternalAddress(data.ValueCell(i, 1))
        body.Cells(i + 1, 5).FormulaR1C1 = IIf(i = 1, "=0", "=R[-1]C[1]")
    Next
    With body.Offset(1, 0).Resize(n, 1)
        .Offset(0, 2).FormulaR1C1 = "=IF(ISERROR(RC[-1]),FALSE,IF(ISNUMBER(RC[-1]),FALSE,OR(LOWER(TRIM(RC[-1]&""""))={""" & Replace(WF_MARKERS, ",", """,""") & """})))"   ' erro na origem = 0, nao total
        .Offset(0, 3).FormulaR1C1 = "=IF(RC[-1],0,IF(ISNUMBER(RC[-2]),RC[-2],0))"
        .Offset(0, 5).FormulaR1C1 = "=RC[-1]+RC[-2]"
        .Offset(0, 6).FormulaR1C1 = "=IF(RC[-4],MIN(0,RC[-1]),MIN(RC[-2],RC[-1]))"
        .Offset(0, 7).FormulaR1C1 = "=IF(RC[-5],MAX(0,RC[-2]),MAX(RC[-3],RC[-2]))"
        .Offset(0, 8).FormulaR1C1 = "=MAX(0,RC[-2])+MIN(0,RC[-1])"
        .Offset(0, 9).FormulaR1C1 = "=MAX(0,RC[-2])-MAX(0,RC[-3])"
        .Offset(0, 10).FormulaR1C1 = "=MIN(0,RC[-4])-MIN(0,RC[-3])"
        .Offset(0, 11).FormulaR1C1 = "=IF(AND(NOT(RC[-9]),RC[-8]>=0),RC[-2],0)"
        .Offset(0, 12).FormulaR1C1 = "=IF(AND(NOT(RC[-10]),RC[-9]>=0),RC[-2],0)"
        .Offset(0, 13).FormulaR1C1 = "=IF(AND(NOT(RC[-11]),RC[-10]<0),RC[-4],0)"
        .Offset(0, 14).FormulaR1C1 = "=IF(AND(NOT(RC[-12]),RC[-11]<0),RC[-4],0)"
        .Offset(0, 15).FormulaR1C1 = "=IF(RC[-13],RC[-10],0)"
        .Offset(0, 16).FormulaR1C1 = "=IF(RC[-14],RC[-11],RC[-13])"
        .Offset(0, 17).FormulaR1C1 = "=IF(RC[-1]>=0,RC[-10],RC[-11])"
    End With
    Set Waterfall_BuildBlock = blk
End Function

Public Function Waterfall_Insert(ws As Worksheet, data As clsChartData) As ChartObject
    Dim id As String, body As Range, n As Long, co As ChartObject, cht As Chart, anchor As Range
    Dim en As Long, es As String, ed As String
    Dim names As Variant, cols As Variant, k As Long, s As Series, sh As String
    On Error GoTo Fail
    id = Registry_NewId()
    Set body = Block_Body(Waterfall_BuildBlock(ws.Parent, id, data))
    Util_TestHookAfterBlock
    n = data.CategoryCount
    sh = "='" & CALC_SHEET & "'!"
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, Chart_FreeTop(ws, anchor.Left, anchor.Top, 460, 280), 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    names = Array("tqq_base", "tqq_uppos", "tqq_upneg", "tqq_downpos", "tqq_downneg", "tqq_total")
    cols = Array(WF_COL_BASE, WF_COL_UPPOS, WF_COL_UPNEG, WF_COL_DOWNPOS, WF_COL_DOWNNEG, WF_COL_TOTAL)
    For k = 0 To 5
        Set s = cht.SeriesCollection.NewSeries
        s.Name = names(k)
        s.Values = sh & body.Cells(2, cols(k)).Resize(n, 1).Address(True, True)
        s.XValues = ChartData_ExternalAddress(data.CategoryCells())
    Next
    cht.ChartType = xlColumnStacked
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlZero
    cht.ChartGroups(1).GapWidth = 50
    Style_ApplyBase cht
    cht.HasLegend = False
    Waterfall_ApplyColors cht, "padrao"
    Registry_SetOption cht, "numfmt", "int"
    Registry_SetOption cht, "palette", "padrao"
    Registry_SetOption cht, "gap", "50"
    Registry_SetOption cht, "totals", "0"
    Waterfall_SetLabels cht, True
    Ann_SetWaterfallConnectors cht, True
    Set Waterfall_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

Public Sub Waterfall_ApplyColors(cht As Chart, paletteName As String)
    Dim pal As Variant, s As Series
    pal = Style_Palette(paletteName)
    For Each s In cht.SeriesCollection
        Select Case s.Name
            Case "tqq_base"
                s.Format.Fill.Visible = msoFalse
                s.Format.Line.Visible = msoFalse
            Case "tqq_uppos", "tqq_upneg"
                Style_SetColor s.Format.Fill.ForeColor, pal(1): s.Format.Line.Visible = msoFalse
            Case "tqq_downpos", "tqq_downneg"
                Style_SetColor s.Format.Fill.ForeColor, pal(2): s.Format.Line.Visible = msoFalse
            Case "tqq_total"
                Style_SetColor s.Format.Fill.ForeColor, pal(0): s.Format.Line.Visible = msoFalse
        End Select
    Next
End Sub

Public Sub Waterfall_SetLabels(cht As Chart, turnOn As Boolean)
    Dim s As Series, body As Range, n As Long, sh As String, fmt As String
    Set s = Chart_FindSeries(cht, "tqq_lbl")
    If Not turnOn Then
        If Not s Is Nothing Then s.Delete
        Registry_SetOption cht, "labels", "0"
        Exit Sub
    End If
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    sh = "='" & CALC_SHEET & "'!"
    If s Is Nothing Then
        Set s = cht.SeriesCollection.NewSeries
        s.Name = "tqq_lbl"
        s.Values = sh & body.Cells(2, WF_COL_LABELPOS).Resize(n, 1).Address(True, True)
        s.ChartType = xlLine
        s.Format.Line.Visible = msoFalse
        s.MarkerStyle = xlMarkerStyleNone
    End If
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    Labels_LinkPoints s, body.Cells(2, WF_COL_LABELVAL).Resize(n, 1), fmt
    Registry_SetOption cht, "labels", "1"
End Sub

Public Function Waterfall_EndValues(cht As Chart) As Double()
    Dim body As Range, n As Long
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    Waterfall_EndValues = Util_ColumnToDoubles(body.Cells(2, WF_COL_END).Resize(n, 1))
End Function
