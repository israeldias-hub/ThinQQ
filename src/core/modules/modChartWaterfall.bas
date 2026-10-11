Attribute VB_Name = "modChartWaterfall"
Option Explicit
' Bloco (corpo): linha 1 = cabeçalhos; 20 colunas (S = 1) ou 20 + 10 S (S > 1 séries de dados):
'  1 Cat | 2 Src | 3 IsTotal | 4 Delta | 5 Start | 6 End | 7 Lo | 8 Hi | 9 Base | 10 PiecePos | 11 PieceNeg
' 12 UpPos | 13 UpNeg | 14 DownPos | 15 DownNeg | 16 Total | 17 LabelVal | 18 LabelPos | 19 Pos | 20 BD
' e, só com S > 1, 10 colunas por série j a partir da coluna WF_MS_BASE + WF_MS_WIDTH (j - 1):
' D | CP | CN | A | B | M | UP | DN | UPN | DNN (Waterfall_SeriesCol). Com S = 1 as colunas 10..15 alimentam o gráfico; com
' S > 1 elas ficam calculadas mas o gráfico usa as colunas UP/DN/UPN/DNN de cada série.
' Pos = n - i + 0,5 (posição da categoria no eixo Y secundário das auxiliares XY das barras); BD = "build-down" (o 'e' na
' primeira linha e um número na última: o primeiro total é calculado para trás, e a última linha vira total).
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
Public Const WF_COL_POS As Long = 19
Public Const WF_COL_BD As Long = 20
Public Const WF_MS_BASE As Long = 21
Public Const WF_MS_WIDTH As Long = 10
Public Const WF_MAX_SERIES As Long = 6
' Unica fonte da lista de marcadores de total (minusculas, sem espacos): usada pela funcao VBA e pela formula do bloco.
Public Const WF_MARKERS As String = "e,t,=,total,subtotal"
' Erro de validação dos dados (faixa 6007..6099: Core_Insert mostra a mensagem como aviso, sem registrar falha).
Private Const WF_ERR_TOO_MANY As String = "A cascata aceita no máximo 6 séries."

Public Function Waterfall_IsTotalMarker(v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If VarType(v) <> vbString Then Exit Function
    Waterfall_IsTotalMarker = (InStr("," & WF_MARKERS & ",", "," & LCase$(Trim$(CStr(v))) & ",") > 0) And Len(Trim$(CStr(v))) > 0
End Function

' Coluna (no corpo do bloco) da parte indicada da série j (1..WF_MAX_SERIES): D, CP, CN, A, B, M, UP, DN, UPN ou DNN.
' 0 para parte desconhecida.
Public Function Waterfall_SeriesCol(j As Long, part As String) As Long
    Dim k As Long
    k = PartIndex(part)
    If k >= 0 Then Waterfall_SeriesCol = WF_MS_BASE + WF_MS_WIDTH * (j - 1) + k
End Function

Private Function PartIndex(part As String) As Long
    Select Case part
        Case "D": PartIndex = 0
        Case "CP": PartIndex = 1
        Case "CN": PartIndex = 2
        Case "A": PartIndex = 3
        Case "B": PartIndex = 4
        Case "M": PartIndex = 5
        Case "UP": PartIndex = 6
        Case "DN": PartIndex = 7
        Case "UPN": PartIndex = 8
        Case "DNN": PartIndex = 9
        Case Else: PartIndex = -1
    End Select
End Function

' Número de séries de dados da cascata, lido da largura do bloco (20 colunas = 1 série; 20 + 10 S colunas = S séries).
Public Function Waterfall_SeriesCount(cht As Chart) As Long
    Dim blk As Range, k As Long
    Waterfall_SeriesCount = 1
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    k = (blk.Columns.Count - WF_COL_BD) \ WF_MS_WIDTH
    If k > 1 Then Waterfall_SeriesCount = k
End Function

' Referencia R1C1 absoluta (com a planilha) as celulas de origem da linha i (categoria i) das series jFrom..jTo.
Private Function SrcRef(data As clsChartData, i As Long, jFrom As Long, jTo As Long) As String
    Dim a As Range, b As Range, sh As String
    Set a = data.ValueCell(i, jFrom): Set b = data.ValueCell(i, jTo)
    sh = "'" & Replace(a.Worksheet.Name, "'", "''") & "'!"
    SrcRef = sh & "R" & a.Row & "C" & a.Column
    If jTo <> jFrom Then SrcRef = SrcRef & ":R" & b.Row & "C" & b.Column
End Function

' Constante de matriz com os marcadores de total ({"e","t",...}), montada a partir de WF_MARKERS.
Private Function MarkerList() As String
    MarkerList = "{""" & Replace(WF_MARKERS, ",", """,""") & """}"
End Function

' Expressao "a linha i e de total" sobre as CELULAS DE ORIGEM (referencias absolutas), nunca sobre a coluna IsTotal: a
' IsTotal da ultima linha depende de BD, e BD usa esta expressao (usar IsTotal aqui seria referencia circular).
' S = 1: a expressao da Fase 1; S > 1: algum valor da linha e um marcador.
Private Function MarkerExpr(data As clsChartData, i As Long) As String
    Dim ref As String
    If data.SeriesCount = 1 Then
        ref = SrcRef(data, i, 1, 1)
        MarkerExpr = "IF(ISERROR(" & ref & "),FALSE,IF(ISNUMBER(" & ref & "),FALSE,OR(LOWER(TRIM(" & ref & "&""""))=" & MarkerList() & ")))"
    Else
        ref = SrcRef(data, i, 1, data.SeriesCount)
        MarkerExpr = "SUMPRODUCT(--ISNUMBER(MATCH(LOWER(TRIM(" & ref & "&""""))," & MarkerList() & ",0)))>0"
    End If
End Function

Public Function Waterfall_BuildBlock(wb As Workbook, id As String, data As clsChartData, Optional kind As String = "waterfall") As Range
    Dim blk As Range, body As Range, n As Long, i As Long, hdr As Variant, c As Long, ns As Long, j As Long
    Dim valRow As String, delta As String, d As String, cpPrev As String, cnPrev As String, cp As String, cn As String
    Dim a As String, b As String, m1 As String, mn As String, tailRow As String
    n = data.CategoryCount
    ns = data.SeriesCount
    If ns > WF_MAX_SERIES Then Err.Raise vbObjectError + 6007, "Waterfall_BuildBlock", WF_ERR_TOO_MANY
    Set blk = CalcSheet_AllocateBlock(wb, id, kind, Mid$(ChartData_ExternalAddress(data.Source), 2), n + 1, WF_COL_BD + IIf(ns > 1, WF_MS_WIDTH * ns, 0))
    Set body = Block_Body(blk)
    hdr = Array("Cat", "Src", "IsTotal", "Delta", "Start", "End", "Lo", "Hi", "Base", "PiecePos", "PieceNeg", _
                "UpPos", "UpNeg", "DownPos", "DownNeg", "Total", "LabelVal", "LabelPos", "Pos", "BD")
    For c = 0 To WF_COL_BD - 1
        body.Cells(1, c + 1).Value = hdr(c)
    Next
    If ns > 1 Then
        hdr = Array("D", "CP", "CN", "A", "B", "M", "UP", "DN", "UPN", "DNN")
        For j = 1 To ns
            For c = 0 To WF_MS_WIDTH - 1
                body.Cells(1, WF_MS_BASE + WF_MS_WIDTH * (j - 1) + c).Value = hdr(c) & j
            Next
        Next
    End If
    valRow = SrcRef(data, n, 1, ns)                       ' valores da ultima linha (build-down)
    m1 = MarkerExpr(data, 1): mn = MarkerExpr(data, n)   ' marcador na primeira e na ultima linha (celulas de origem)
    For i = 1 To n
        body.Cells(i + 1, 1).Formula = ChartData_ExternalAddress(data.Source.Cells(i + 1, 1))
        body.Cells(i + 1, 2).Formula = ChartData_ExternalAddress(data.ValueCell(i, 1))
        body.Cells(i + 1, 5).FormulaR1C1 = IIf(i = 1, "=0", "=R[-1]C[1]")
        body.Cells(i + 1, WF_COL_POS).Formula = "=" & (n - i) & ".5"         ' independe do separador decimal da localidade
        body.Cells(i + 1, WF_COL_BD).FormulaR1C1 = "=AND(" & m1 & ",COUNT(" & valRow & ")>0,NOT(" & mn & "))"
    Next
    With body.Offset(1, 0).Resize(n, 1)
        If ns = 1 Then
            .Offset(0, 2).FormulaR1C1 = "=IF(ISERROR(RC[-1]),FALSE,IF(ISNUMBER(RC[-1]),FALSE,OR(LOWER(TRIM(RC[-1]&""""))=" & MarkerList() & ")))"   ' erro na origem = 0, nao total
            .Offset(0, 3).FormulaR1C1 = "=IF(RC[-1],0,IF(ISNUMBER(RC[-2]),RC[-2],0))"
            .Offset(0, 6).FormulaR1C1 = "=IF(RC[-4],MIN(0,RC[-1]),MIN(RC[-2],RC[-1]))"
            .Offset(0, 7).FormulaR1C1 = "=IF(RC[-5],MAX(0,RC[-2]),MAX(RC[-3],RC[-2]))"
            .Offset(0, 16).FormulaR1C1 = "=IF(RC[-14],RC[-11],RC[-13])"
            .Offset(0, 17).FormulaR1C1 = "=IF(RC[-1]>=0,RC[-10],RC[-11])"
        Else
            For i = 1 To n
                body.Cells(i + 1, 3).FormulaR1C1 = "=" & MarkerExpr(data, i)
            Next
            delta = ""
            For j = 1 To ns
                delta = delta & IIf(j > 1, ",", "") & "RC" & Waterfall_SeriesCol(j, "D")
            Next
            .Offset(0, 3).FormulaR1C1 = "=SUM(" & delta & ")"
            .Offset(0, 6).FormulaR1C1 = "=IF(RC3,MIN(0,RC6),RC5-RC" & Waterfall_SeriesCol(ns, "CN") & ")"
            .Offset(0, 7).FormulaR1C1 = "=IF(RC3,MAX(0,RC6),RC5+RC" & Waterfall_SeriesCol(ns, "CP") & ")"
            .Offset(0, 16).FormulaR1C1 = "=IF(RC3,RC16,NA())"
            .Offset(0, 17).FormulaR1C1 = "=IF(RC3,IF(RC16>=0,RC8,RC7),NA())"
        End If
        .Offset(0, 5).FormulaR1C1 = "=RC5+RC4"
        .Offset(0, 8).FormulaR1C1 = "=MAX(0,RC[-2])+MIN(0,RC[-1])"
        .Offset(0, 9).FormulaR1C1 = "=MAX(0,RC[-2])-MAX(0,RC[-3])"
        .Offset(0, 10).FormulaR1C1 = "=MIN(0,RC[-4])-MIN(0,RC[-3])"
        .Offset(0, 11).FormulaR1C1 = "=IF(AND(NOT(RC[-9]),RC[-8]>=0),RC[-2],0)"
        .Offset(0, 12).FormulaR1C1 = "=IF(AND(NOT(RC[-10]),RC[-9]>=0),RC[-2],0)"
        .Offset(0, 13).FormulaR1C1 = "=IF(AND(NOT(RC[-11]),RC[-10]<0),RC[-4],0)"
        .Offset(0, 14).FormulaR1C1 = "=IF(AND(NOT(RC[-12]),RC[-11]<0),RC[-4],0)"
        .Offset(0, 15).FormulaR1C1 = "=IF(RC[-13],RC[-10],0)"
    End With
    ' Build-down e ultima linha. Com uma unica linha BD e sempre falso (o marcador e a ausencia dele na mesma linha).
    If n > 1 Then
        tailRow = IIf(n = 2, "RC4", "R[" & (n - 2) & "]C4")                     ' com n = 2 a soma cobre so linhas de total (Delta 0)
        body.Cells(2, WF_COL_END).FormulaR1C1 = "=IF(RC20,AGGREGATE(9,6," & valRow & ")-SUM(R[1]C4:" & tailRow & "),RC5+RC4)"
        body.Cells(n + 1, WF_COL_END).FormulaR1C1 = "=IF(RC20,AGGREGATE(9,6," & valRow & "),RC5+RC4)"
    End If
    body.Cells(n + 1, 3).FormulaR1C1 = "=OR(" & mn & ",RC20)"
    If ns > 1 Then
        For j = 1 To ns
            d = "RC" & Waterfall_SeriesCol(j, "D"): cp = "RC" & Waterfall_SeriesCol(j, "CP"): cn = "RC" & Waterfall_SeriesCol(j, "CN")
            If j = 1 Then
                cpPrev = "0": cnPrev = "0"
            Else
                cpPrev = "RC" & Waterfall_SeriesCol(j - 1, "CP"): cnPrev = "RC" & Waterfall_SeriesCol(j - 1, "CN")
            End If
            a = "RC" & Waterfall_SeriesCol(j, "A"): b = "RC" & Waterfall_SeriesCol(j, "B")
            For i = 1 To n
                body.Cells(i + 1, Waterfall_SeriesCol(j, "D")).FormulaR1C1 = "=IF(RC3,0,IF(ISNUMBER(" & SrcRef(data, i, j, j) & ")," & SrcRef(data, i, j, j) & ",0))"
            Next
            With body.Offset(1, 0).Resize(n, 1)
                .Offset(0, Waterfall_SeriesCol(j, "CP") - 1).FormulaR1C1 = "=" & IIf(j > 1, cpPrev & "+", "") & "MAX(0," & d & ")"
                .Offset(0, Waterfall_SeriesCol(j, "CN") - 1).FormulaR1C1 = "=" & IIf(j > 1, cnPrev & "+", "") & "MAX(0,-" & d & ")"
                .Offset(0, Waterfall_SeriesCol(j, "A") - 1).FormulaR1C1 = "=IF(" & d & ">=0,RC5+" & cpPrev & ",RC5-" & cn & ")"
                .Offset(0, Waterfall_SeriesCol(j, "B") - 1).FormulaR1C1 = "=IF(" & d & ">=0,RC5+" & cp & ",RC5-" & cnPrev & ")"
                .Offset(0, Waterfall_SeriesCol(j, "M") - 1).FormulaR1C1 = "=(" & a & "+" & b & ")/2"
                .Offset(0, Waterfall_SeriesCol(j, "UP") - 1).FormulaR1C1 = "=IF(" & d & ">=0,MAX(0," & b & ")-MAX(0," & a & "),0)"
                .Offset(0, Waterfall_SeriesCol(j, "DN") - 1).FormulaR1C1 = "=IF(" & d & "<0,MAX(0," & b & ")-MAX(0," & a & "),0)"
                .Offset(0, Waterfall_SeriesCol(j, "UPN") - 1).FormulaR1C1 = "=IF(" & d & ">=0,MIN(0," & a & ")-MIN(0," & b & "),0)"
                .Offset(0, Waterfall_SeriesCol(j, "DNN") - 1).FormulaR1C1 = "=IF(" & d & "<0,MIN(0," & a & ")-MIN(0," & b & "),0)"
            End With
        Next
    End If
    Set Waterfall_BuildBlock = blk
End Function

' Cria uma serie de coluna empilhada da cascata (valores no bloco, categorias na origem). name pode ser texto ou uma
' referencia (="=Planilha!$B$1").
Private Function AddWfSeries(cht As Chart, name As String, valueCells As Range, cats As String) As Series
    Dim s As Series
    Set s = cht.SeriesCollection.NewSeries
    s.Name = name
    s.Values = Block_Ref(valueCells)
    s.XValues = cats
    Set AddWfSeries = s
End Function

' kind: "waterfall" (colunas empilhadas, o padrao) ou "barwaterfall" (barras empilhadas horizontais com as categorias
' invertidas por Style_ApplyBase; rotulos em auxiliares XY no grupo secundario; conectores verticais). O tipo vai para a
' meta do bloco (Registry_ChartType).
Public Function Waterfall_Insert(ws As Worksheet, data As clsChartData, Optional kind As String = "waterfall") As ChartObject
    Dim id As String, body As Range, n As Long, co As ChartObject, cht As Chart, anchor As Range, ns As Long, j As Long
    Dim en As Long, es As String, ed As String
    Dim names As Variant, cols As Variant, k As Long, cats As String, horiz As Boolean
    If Kinds_Family(kind) <> "waterfall" Then Err.Raise 5, "Waterfall_Insert", "Tipo de cascata desconhecido: " & kind      ' antes de alocar o bloco
    horiz = Kinds_IsHorizontal(kind)
    ns = data.SeriesCount
    If ns > WF_MAX_SERIES Then Err.Raise vbObjectError + 6007, "Waterfall_Insert", WF_ERR_TOO_MANY     ' antes de alocar o bloco
    On Error GoTo Fail
    id = Registry_NewId()
    Set body = Block_Body(Waterfall_BuildBlock(ws.Parent, id, data, kind))
    Util_TestHookAfterBlock
    n = data.CategoryCount
    cats = ChartData_ExternalAddress(data.CategoryCells())
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, Chart_FreeTop(ws, anchor.Left, anchor.Top, 460, 280), 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    If ns = 1 Then
        names = Array("tqq_base", "tqq_uppos", "tqq_upneg", "tqq_downpos", "tqq_downneg", "tqq_total")
        cols = Array(WF_COL_BASE, WF_COL_UPPOS, WF_COL_UPNEG, WF_COL_DOWNPOS, WF_COL_DOWNNEG, WF_COL_TOTAL)
        For k = 0 To 5
            AddWfSeries cht, CStr(names(k)), body.Cells(2, cols(k)).Resize(n, 1), cats
        Next
    Else
        ' Pilha positiva (de baixo para cima): base, DN_S..DN_1 (parte acima de zero das quedas), UP_1..UP_S (as series de
        ' dados, com o nome da coluna); pilha negativa (de cima para baixo): UPN_S..UPN_1, DNN_1..DNN_S; por fim os totais.
        AddWfSeries cht, "tqq_base", body.Cells(2, WF_COL_BASE).Resize(n, 1), cats
        For j = ns To 1 Step -1
            AddWfSeries cht, "tqq_dn_" & j, body.Cells(2, Waterfall_SeriesCol(j, "DN")).Resize(n, 1), cats
        Next
        For j = 1 To ns
            AddWfSeries cht, ChartData_ExternalAddress(data.SeriesNameCell(j)), body.Cells(2, Waterfall_SeriesCol(j, "UP")).Resize(n, 1), cats
        Next
        For j = ns To 1 Step -1
            AddWfSeries cht, "tqq_upn_" & j, body.Cells(2, Waterfall_SeriesCol(j, "UPN")).Resize(n, 1), cats
        Next
        For j = 1 To ns
            AddWfSeries cht, "tqq_dnn_" & j, body.Cells(2, Waterfall_SeriesCol(j, "DNN")).Resize(n, 1), cats
        Next
        AddWfSeries cht, "tqq_total", body.Cells(2, WF_COL_TOTAL).Resize(n, 1), cats
    End If
    cht.ChartType = IIf(horiz, xlBarStacked, xlColumnStacked)
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlZero
    cht.ChartGroups(1).GapWidth = 50
    Style_ApplyBase cht
    If ns > 1 Then
        cht.HasLegend = True
        cht.Legend.Position = xlLegendPositionRight
        cht.Legend.Font.Size = 9
        Registry_SetOption cht, "legend", "right"
    Else
        cht.HasLegend = False
    End If
    Waterfall_ApplyColors cht, "padrao"
    Registry_SetOption cht, "numfmt", "int"
    Registry_SetOption cht, "palette", "padrao"
    Registry_SetOption cht, "gap", "50"
    Registry_SetOption cht, "totals", "0"
    Registry_SetOption cht, "n", CStr(data.CategoryCount)
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

' Cor das series de dados e das partes auxiliares. Uma serie: como na Fase 1 (alta/queda/total). Varias: todas as partes
' da serie j (UP/DN/UPN/DNN) recebem a cor j da paleta (ciclo de 6) e os totais o cinza (indice 3); o j das auxiliares
' vem do sufixo _j do nome e o das series de dados (UP_j, com o nome da coluna) da ordem entre as nao auxiliares.
Public Sub Waterfall_ApplyColors(cht As Chart, paletteName As String)
    Dim pal As Variant, s As Series, multi As Boolean, nm As String, j As Long, k As Long, p As Long
    pal = Style_Palette(paletteName)
    multi = (Waterfall_SeriesCount(cht) > 1)
    For Each s In cht.SeriesCollection
        nm = s.Name
        If nm = "tqq_base" Then
            s.Format.Fill.Visible = msoFalse
            s.Format.Line.Visible = msoFalse
        ElseIf multi Then
            j = 0
            If nm = "tqq_total" Then
                Style_SetColor s.Format.Fill.ForeColor, pal(3): s.Format.Line.Visible = msoFalse
            ElseIf Left$(nm, 7) = "tqq_dn_" Or Left$(nm, 8) = "tqq_upn_" Or Left$(nm, 8) = "tqq_dnn_" Then
                p = InStrRev(nm, "_")
                j = CLng(Val(Mid$(nm, p + 1)))
            ElseIf Not Style_IsHelperSeries(s) Then
                k = k + 1
                j = k
            End If
            If j > 0 Then Style_SetColor s.Format.Fill.ForeColor, pal((j - 1) Mod 6): s.Format.Line.Visible = msoFalse
        Else
            Select Case nm
                Case "tqq_uppos", "tqq_upneg"
                    Style_SetColor s.Format.Fill.ForeColor, pal(1): s.Format.Line.Visible = msoFalse
                Case "tqq_downpos", "tqq_downneg"
                    Style_SetColor s.Format.Fill.ForeColor, pal(2): s.Format.Line.Visible = msoFalse
                Case "tqq_total"
                    Style_SetColor s.Format.Fill.ForeColor, pal(0): s.Format.Line.Visible = msoFalse
            End Select
        End If
    Next
    If multi Then RefreshSegmentLabelColors cht
End Sub

' Cor da fonte dos rotulos de segmento (tqq_lbl_j) por contraste com a cor atual da serie j (a paleta pode ter mudado).
Private Sub RefreshSegmentLabelColors(cht As Chart)
    Dim ns As Long, j As Long, ls As Series
    ns = Waterfall_SeriesCount(cht)
    For j = 1 To ns
        Set ls = Chart_FindSeries(cht, "tqq_lbl_" & j)
        If Not ls Is Nothing Then
            If ls.HasDataLabels Then ls.DataLabels.Font.Color = Style_ContrastFont(DataSeriesColor(cht, j))
        End If
    Next
End Sub

' Cor (RGB) atual da j-esima serie de dados (a nao auxiliar de ordem j).
Private Function DataSeriesColor(cht As Chart, j As Long) As Long
    Dim s As Series, k As Long
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            k = k + 1
            If k = j Then DataSeriesColor = CLng(s.Format.Fill.ForeColor.RGB): Exit Function
        End If
    Next
End Function

' Rotulos: S = 1 como na Fase 1 (tqq_lbl em LabelPos, ligado a LabelVal). S > 1: um auxiliar tqq_lbl_j por serie
' (em M_j, o meio do segmento, rotulo ligado a D_j, posicao central, fonte por contraste com a cor da serie) e tqq_lbl
' so para as linhas de total (LabelPos/LabelVal tem NA() nas demais). Em barwaterfall os auxiliares sao series XY no grupo
' secundario (Chart_AddHelperSeries com a coluna Pos como Y; a "posicao" vira o X) e os rotulos de total/delta saem a
' direita (valor >= 0, em Hi) ou a esquerda (valor < 0, em Lo).
Public Sub Waterfall_SetLabels(cht As Chart, turnOn As Boolean)
    Dim s As Series, body As Range, n As Long, fmt As String, ns As Long, j As Long, i As Long, horiz As Boolean, posCol As Range
    ns = Waterfall_SeriesCount(cht)
    If Not turnOn Then
        For i = cht.SeriesCollection.Count To 1 Step -1
            Set s = cht.SeriesCollection(i)
            If s.Name = "tqq_lbl" Or Left$(s.Name, 8) = "tqq_lbl_" Then s.Delete
        Next
        Registry_SetOption cht, "labels", "0"
        Exit Sub
    End If
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    horiz = Kinds_IsHorizontal(Registry_ChartType(cht))
    If horiz Then Set posCol = body.Cells(2, WF_COL_POS).Resize(n, 1)       ' Y das auxiliares XY: centro da faixa da categoria
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    If ns > 1 Then
        For j = 1 To ns
            Set s = Chart_FindSeries(cht, "tqq_lbl_" & j)
            If s Is Nothing Then Set s = Chart_AddHelperSeries(cht, "tqq_lbl_" & j, body.Cells(2, Waterfall_SeriesCol(j, "M")).Resize(n, 1), posCol)
            Labels_LinkPoints s, body.Cells(2, Waterfall_SeriesCol(j, "D")).Resize(n, 1), fmt, xlLabelPositionCenter, Style_ContrastFont(DataSeriesColor(cht, j)), horiz
        Next
    End If
    Set s = Chart_FindSeries(cht, "tqq_lbl")
    If s Is Nothing Then
        Set s = Chart_AddHelperSeries(cht, "tqq_lbl", body.Cells(2, WF_COL_LABELPOS).Resize(n, 1), posCol)
    End If
    If ns > 1 Then
        ' O texto do rotulo ligado a uma celula e o da celula: LabelVal tem #N/D fora das linhas de total e o rotulo mostraria
        ' "#N/D" (medido no Excel 2019: no zero do eixo, por baixo das barras). A coluna Total vale 0 nessas linhas e o
        ' formato numerico (zero vazio) apaga o texto; nas linhas de total ela vale o mesmo que LabelVal. A posicao do
        ' ponto continua em LabelPos (NA() = ponto nao plotado).
        Labels_LinkPoints s, body.Cells(2, WF_COL_TOTAL).Resize(n, 1), fmt, , , horiz
        Chart_HideHelperLegendEntries cht
    Else
        Labels_LinkPoints s, body.Cells(2, WF_COL_LABELVAL).Resize(n, 1), fmt, , , horiz
    End If
    Registry_SetOption cht, "labels", "1"
End Sub

Public Function Waterfall_EndValues(cht As Chart) As Double()
    Dim body As Range, n As Long
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    Waterfall_EndValues = Util_ColumnToDoubles(body.Cells(2, WF_COL_END).Resize(n, 1))
End Function
