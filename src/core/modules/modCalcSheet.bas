Attribute VB_Name = "modCalcSheet"
Option Explicit
' Planilha muito oculta "_ThinQQ". Cada gráfico tem um bloco: linha 1 = meta
' (A=id, B=tipo, C=endereço da fonte sem "=", D=opções "chave=valor;..."), demais linhas = corpo.
' O nome definido oculto "tqq_<id>" aponta para o bloco inteiro.

Public Function CalcSheet_Get(wb As Workbook, Optional createIfMissing As Boolean = True) As Worksheet
    Dim ws As Worksheet, active As Object, su As Boolean, en As Long, ed As String, created As Boolean
    On Error Resume Next
    Set ws = wb.Worksheets(CALC_SHEET)
    On Error GoTo 0
    If ws Is Nothing And createIfMissing Then
        On Error GoTo Fail
        Set active = wb.ActiveSheet
        su = Application.ScreenUpdating
        Application.ScreenUpdating = False
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        created = True
        ws.Name = CALC_SHEET
        ws.Range("A1").Value = "ThinQQ - planilha interna de cálculo. Não editar."
        ws.Visible = xlSheetVeryHidden
        If Not active Is Nothing Then active.Activate
        Application.ScreenUpdating = su
        On Error GoTo 0
    End If
    Set CalcSheet_Get = ws
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    If created Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Application.ScreenUpdating = su
    On Error GoTo 0
    Err.Raise vbObjectError + 6001, "CalcSheet_Get", "Não foi possível criar a planilha interna do ThinQQ (estrutura da pasta protegida?). Detalhe: " & ed & " (" & en & ")"
End Function

Private Function NameExists(wb As Workbook, nm As String) As Boolean
    Dim n As Name
    For Each n In wb.Names
        If StrComp(n.Name, nm, vbTextCompare) = 0 Then NameExists = True: Exit Function
        If Right$(n.Name, Len(nm) + 1) = "!" & nm Then NameExists = True: Exit Function
    Next
End Function

' Existe bloco (nome tqq_<id>) para este id?
Public Function CalcSheet_BlockExists(wb As Workbook, id As String) As Boolean
    CalcSheet_BlockExists = NameExists(wb, ID_PREFIX & id)
End Function

' Fórmula de texto ="p1"&"p2" com pedaços de até 200 caracteres (limite de 255 por literal).
Private Function TextFormula(s As String) As String
    Dim i As Long, out As String
    If Len(s) = 0 Then TextFormula = "=""""": Exit Function
    For i = 1 To Len(s) Step 200
        out = out & IIf(i > 1, "&", "") & """" & Mid$(s, i, 200) & """"
    Next
    TextFormula = "=" & out
End Function

' Primeira linha livre abaixo do último bloco registrado (deixa 1 linha em branco entre blocos).
Private Function NextFreeRow(wb As Workbook, ws As Worksheet) As Long
    Dim nm As Name, rng As Range, last As Long
    last = 1
    For Each nm In wb.Names
        If Left$(nm.Name, Len(ID_PREFIX)) = ID_PREFIX Then
            Set rng = Nothing
            On Error Resume Next
            Set rng = nm.RefersToRange
            On Error GoTo 0
            If Not rng Is Nothing Then
                If rng.Worksheet.Name = ws.Name Then
                    If rng.Row + rng.Rows.Count - 1 > last Then last = rng.Row + rng.Rows.Count - 1
                End If
            End If
        End If
    Next
    NextFreeRow = last + 2
End Function

Public Function CalcSheet_AllocateBlock(wb As Workbook, id As String, chartType As String, sourceAddr As String, bodyRows As Long, bodyCols As Long) As Range
    Dim ws As Worksheet, r As Long, blk As Range
    If bodyRows < 1 Then Err.Raise vbObjectError + 6002, "CalcSheet_AllocateBlock", "bodyRows deve ser >= 1."
    If InStr(sourceAddr, """") > 0 Then Err.Raise vbObjectError + 6003, "CalcSheet_AllocateBlock", "Endereço de origem não pode conter aspas duplas."
    If NameExists(wb, ID_PREFIX & id) Then Err.Raise vbObjectError + 6004, "CalcSheet_AllocateBlock", "Já existe um bloco com o id '" & id & "'."
    Set ws = CalcSheet_Get(wb)
    r = NextFreeRow(wb, ws)
    Set blk = ws.Cells(r, 1).Resize(bodyRows + 1, IIf(bodyCols > 4, bodyCols, 4))
    blk.Clear
    blk.Cells(1, 1).NumberFormat = "@": blk.Cells(1, 1).Value = id
    blk.Cells(1, 2).Value = chartType
    ' Endereço como fórmula de texto (="...") para preservar o apóstrofo inicial de 'Nome'!$A$1
    blk.Cells(1, 3).Formula = TextFormula(sourceAddr)
    blk.Cells(1, 4).NumberFormat = "@": blk.Cells(1, 4).Value = ""
    blk.Rows(1).Font.Bold = True
    wb.Names.Add Name:=ID_PREFIX & id, RefersTo:="='" & CALC_SHEET & "'!" & blk.Address(True, True), Visible:=False
    Set CalcSheet_AllocateBlock = blk
End Function

Public Function CalcSheet_FindBlock(wb As Workbook, id As String) As Range
    On Error Resume Next
    Set CalcSheet_FindBlock = wb.Names(ID_PREFIX & id).RefersToRange
    On Error GoTo 0
End Function

Public Sub CalcSheet_DeleteBlock(wb As Workbook, id As String)
    Dim blk As Range
    Set blk = CalcSheet_FindBlock(wb, id)
    If blk Is Nothing Then Exit Sub
    blk.Clear
    wb.Names(ID_PREFIX & id).Delete
End Sub

Public Function Block_Meta(blk As Range) As Range
    Set Block_Meta = blk.Rows(1)
End Function

Public Function Block_Body(blk As Range) As Range
    Set Block_Body = blk.Offset(1, 0).Resize(blk.Rows.Count - 1, blk.Columns.Count)
End Function

' Fórmulas (Series.Formula) de TODOS os gráficos da pasta: gráficos embutidos de cada planilha (inclusive dentro de
' grupos de formas) e planilhas de gráfico. Uma cópia de gráfico ThinQQ, ou um gráfico movido para planilha de
' gráfico, não tem mais o nome "tqq_<id>", mas continua apontando para o bloco: a vida do bloco é decidida por referência.
Private Function AllSeriesFormulas(wb As Workbook) As String
    Dim ws As Worksheet, sh As Object, gi As Object, co As Object, acc As String
    On Error Resume Next
    For Each ws In wb.Worksheets
        For Each sh In ws.Shapes
            If sh.Type = 6 Then                         ' msoGroup
                For Each gi In sh.GroupItems
                    If gi.HasChart Then acc = acc & ChartFormulas(gi.Chart)
                Next
            ElseIf sh.HasChart Then
                acc = acc & ChartFormulas(sh.Chart)
            End If
        Next
    Next
    For Each co In wb.Charts
        acc = acc & ChartFormulas(co)
    Next
    AllSeriesFormulas = acc
End Function

Private Function ChartFormulas(ch As Object) As String
    Dim s As Object, acc As String
    On Error Resume Next
    For Each s In ch.SeriesCollection
        acc = acc & vbLf & s.Formula
    Next
    ChartFormulas = acc
End Function

' Faixas da planilha _ThinQQ citadas no texto de fórmulas (referências de outras pastas, "[x]_ThinQQ", são ignoradas).
Private Function ReferencedCalcRanges(txt As String, calc As Worksheet) As Collection
    Dim refs As New Collection, p As Long, q As Long, addr As String, ch As String
    p = InStr(1, txt, CALC_SHEET, vbTextCompare)
    Do While p > 0
        q = p + Len(CALC_SHEET)
        If Mid$(txt, q, 1) = "'" Then q = q + 1
        If Mid$(txt, q, 1) = "!" And Mid$(txt, IIf(p > 1, p - 1, 1), 1) <> "]" Then
            q = q + 1
            addr = ""
            Do While q <= Len(txt)
                ch = Mid$(txt, q, 1)
                If ch Like "[A-Za-z0-9$:]" Then addr = addr & ch: q = q + 1 Else Exit Do
            Loop
            If Len(addr) > 0 Then
                On Error Resume Next
                refs.Add calc.Range(addr)
                On Error GoTo 0
            End If
        End If
        p = InStr(q, txt, CALC_SHEET, vbTextCompare)
    Loop
    Set ReferencedCalcRanges = refs
End Function

Private Function AnyIntersect(refs As Collection, blk As Range) As Boolean
    Dim r As Variant
    For Each r In refs
        If Not Application.Intersect(r, blk) Is Nothing Then AnyIntersect = True: Exit Function
    Next
End Function

' Remove blocos cujo gráfico não existe mais. Retorna quantos removeu.
' Só considera nomes "tqq_*" cujo destino é um bloco da planilha _ThinQQ (meta A = id).
' Um bloco está vivo se existe um gráfico nomeado "tqq_<id>" OU se alguma série de qualquer gráfico da pasta
' referencia uma célula do bloco (cópias e gráficos movidos para planilha de gráfico).
Public Function CalcSheet_Cleanup(wb As Workbook) As Long
    Dim i As Long, nm As Name, id As String, removed As Long, rng As Range
    Dim refs As Collection, haveRefs As Boolean
    For i = wb.Names.Count To 1 Step -1
        Set nm = wb.Names(i)
        If Left$(nm.Name, Len(ID_PREFIX)) = ID_PREFIX Then
            id = Mid$(nm.Name, Len(ID_PREFIX) + 1)
            Set rng = Nothing
            On Error Resume Next
            Set rng = nm.RefersToRange
            On Error GoTo 0
            If Not rng Is Nothing Then
                If rng.Worksheet.Name = CALC_SHEET Then
                    If CStr(rng.Cells(1, 1).Value) = id Then
                        If Registry_FindChartObject(wb, id) Is Nothing Then
                            If Not haveRefs Then
                                Set refs = ReferencedCalcRanges(AllSeriesFormulas(wb), rng.Worksheet)
                                haveRefs = True
                            End If
                            If Not AnyIntersect(refs, rng) Then
                                CalcSheet_DeleteBlock wb, id
                                removed = removed + 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    CalcSheet_Cleanup = removed
End Function
