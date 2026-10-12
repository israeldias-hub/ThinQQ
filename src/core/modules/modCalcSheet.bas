Attribute VB_Name = "modCalcSheet"
Option Explicit
' Planilha muito oculta "_ThinQQ". Cada gráfico tem um bloco: linha 1 = meta
' (A=id, B=tipo, C=endereço da fonte sem "=", D=opções "chave=valor;..."), demais linhas = corpo.
' O nome definido oculto "tqq_<id>" aponta para o bloco inteiro.
' A vida do bloco é decidida por referência: tanto as fórmulas de série (Series.Formula) quanto os rótulos
' ligados a células (DataLabel.Formula) de qualquer gráfico da pasta mantêm o bloco vivo. A limpeza falha
' fechada: se a leitura de algum gráfico falhar, nada é removido (CalcSheet_CleanupReadFailed).

Private mReadFailed As Boolean
Private mFailDetail As String       ' motivo da primeira falha de leitura (vai para o log)
Private mPointsScanned As Long      ' pontos cujo rótulo foi lido na passada atual
Private Const MAX_POINTS_PER_PASS As Long = 20000

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

' Referencia absoluta (para Series.Values/XValues) a celulas da planilha interna: ='_ThinQQ'!$C$5:$C$9
Public Function Block_Ref(cells As Range) As String
    Block_Ref = "='" & CALC_SHEET & "'!" & cells.Address(True, True)
End Function

' Texto com as fórmulas de série (Series.Formula) E dos rótulos ligados a este livro (DataLabel.Formula) de TODOS os
' gráficos da pasta: gráficos embutidos de cada planilha (inclusive dentro de grupos de formas) e planilhas de
' gráfico. Uma cópia de gráfico ThinQQ, ou um gráfico movido para planilha de gráfico, não tem mais o nome
' "tqq_<id>", mas continua apontando para o bloco (por série ou, num 100% sem totais, só pelos rótulos ligados): a
' vida do bloco é decidida por referência. labelCap = até que ponto ler rótulos em cada série (altura do maior
' bloco). Zera e preenche mReadFailed: qualquer falha de leitura (forma, grupo ou gráfico) o marca e a passada para
' na hora, sem ler os gráficos restantes.
Private Function AllChartReferences(wb As Workbook, labelCap As Long) As String
    Dim ws As Worksheet, sh As Object, gi As Object, co As Object, acc As String
    mReadFailed = False
    mFailDetail = ""
    mPointsScanned = 0
    On Error GoTo Fail
    For Each ws In wb.Worksheets
        For Each sh In ws.Shapes
            TestFailSite READFAIL_SHAPE
            If sh.Type = 6 Then                         ' msoGroup
                For Each gi In sh.GroupItems
                    TestFailSite READFAIL_GROUP
                    If gi.HasChart Then acc = acc & ChartReferences(gi.Chart, labelCap)
                    If mReadFailed Then GoTo Done
                Next
            ElseIf sh.HasChart Then
                acc = acc & ChartReferences(sh.Chart, labelCap)
                If mReadFailed Then GoTo Done
            End If
        Next
    Next
    For Each co In wb.Charts
        TestFailSite READFAIL_CHARTSHEET
        acc = acc & ChartReferences(co, labelCap)
        If mReadFailed Then GoTo Done
    Next
Done:
    AllChartReferences = acc
    Exit Function
Fail:
    MarkFailed "forma ou planilha de gráfico: " & Err.Number & " " & Err.Description
    AllChartReferences = acc
End Function

' Texto com as fórmulas de séries E de rótulos ligados de um gráfico. Qualquer erro de leitura marca mReadFailed.
' Limites da varredura de rótulos (o custo cresce com os pontos, e isto roda a cada Atualizar/Inserir enquanto
' existir uma cópia de gráfico ThinQQ):
'  - só os primeiros labelCap pontos de cada série: Labels_LinkPoints liga os pontos 1..n a UM bloco, logo nenhum
'    rótulo além da altura do maior bloco pode referenciar um bloco;
'  - série cuja própria fórmula já cita a planilha interna: seus rótulos ligam ao mesmo bloco, nada a ler;
'  - para no primeiro rótulo LIGADO à planilha interna (todos os rótulos de uma série ligam ao mesmo bloco); um rótulo de
'    texto livre que só cita "_ThinQQ" entra no texto (conservador) mas não encerra a série (IsCalcSheetLink);
'  - orçamento de MAX_POINTS_PER_PASS pontos por passada: estourou, falha fechada (nada é removido).
Private Function ChartReferences(ch As Object, labelCap As Long) As String
    Dim s As Object, i As Long, n As Long, acc As String, sf As String, lf As String, budget As Long
    Dim en As Long, ed As String
    On Error GoTo Fail
    budget = PointBudget()
    TestFailSite READFAIL_CHART
    For Each s In ch.SeriesCollection
        TestFailSite READFAIL_SERIES
        sf = s.Formula
        acc = acc & vbLf & sf
        If InStr(1, sf, CALC_SHEET, vbTextCompare) = 0 Then
            n = s.Points.Count
            If n > labelCap Then n = labelCap
            For i = 1 To n
                mPointsScanned = mPointsScanned + 1
                If mPointsScanned > budget Then
                    MarkFailed "limite de " & budget & " pontos lidos por limpeza excedido"
                    GoTo Done
                End If
                lf = LabelFormula(s.Points(i))
                If mReadFailed Then GoTo Done
                If InStr(1, lf, CALC_SHEET, vbTextCompare) > 0 Then
                    acc = acc & vbLf & lf
                    ' so um rotulo LIGADO a planilha interna encerra a serie; texto livre que apenas cita "_ThinQQ" continua a varredura
                    If IsCalcSheetLink(lf) Then Exit For
                End If
            Next
        End If
    Next
Done:
    ChartReferences = acc
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    ' Gráfico nativo do Excel 2016+ (cascata, funil, treemap...): o modelo de objetos não expõe Series.Formula
    ' (erro 445). O ThinQQ não os cria e uma cópia de gráfico ThinQQ nunca é um deles: não é falha de leitura.
    If en = 445 Then
        If IsNativeChartEx(ch) Then Exit Function
    End If
    MarkFailed "gráfico: " & en & " " & ed
    ChartReferences = acc
End Function

' Tipos de gráfico nativos do Excel 2016+ ("chartex"): treemap 117, histograma 118, cascata 119, sunburst 120,
' box-whisker 121, pareto 122, funil 123 e mapa 140.
Private Function IsNativeChartEx(ch As Object) As Boolean
    Dim t As Long
    On Error Resume Next
    t = ch.ChartType
    If Err.Number <> 0 Then Err.Clear: Exit Function
    IsNativeChartEx = (t >= 117 And t <= 123) Or t = 140
End Function

' Fórmula do rótulo do ponto: "=_ThinQQ!$D$7" se ligado a uma célula; "" se o ponto não tem rótulo.
' Comportamento observado no Excel 2019 pt-BR (Test_CalcSheet_LabelFormula_ObservedBehavior): ponto sem rótulo =>
' HasDataLabel False (sem erro) e .DataLabel levanta erro; rótulo automático => .Formula devolve o TEXTO exibido
' (ex. "10"); texto editado à mão => .Formula devolve o texto digitado; rótulo ligado => "=Dados!$B$2" (sempre em A1).
' Ou seja, .Formula nunca levanta erro num rótulo existente: qualquer erro aqui (HasDataLabel ou .Formula) é
' leitura falha e fecha a limpeza (não há mais o recurso de ler .Text). Texto livre que cite "_ThinQQ" só pode manter
' um bloco vivo a mais (conservador), nunca apagar um.
Private Function LabelFormula(p As Object) As String
    On Error GoTo Fail
    If BUILD_CONFIG = "Test" Then Util_LabelReadsForTest = Util_LabelReadsForTest + 1
    TestFailSite READFAIL_HASLABEL
    If Not p.HasDataLabel Then Exit Function
    TestFailSite READFAIL_FORMULA
    LabelFormula = p.DataLabel.Formula
    Exit Function
Fail:
    MarkFailed "rótulo: " & Err.Number & " " & Err.Description
End Function

' True quando o texto lido de DataLabel.Formula é uma ligação a uma célula da planilha interna ("=_ThinQQ!$B$2" ou "='_ThinQQ'!$B$2"):
' começa por "=" (texto livre nunca começa por "=") e cita a planilha seguida de "!".
Private Function IsCalcSheetLink(lf As String) As Boolean
    If Left$(lf, 1) <> "=" Then Exit Function
    IsCalcSheetLink = (InStr(1, lf, CALC_SHEET & "!", vbTextCompare) > 0 Or InStr(1, lf, CALC_SHEET & "'!", vbTextCompare) > 0)
End Function

Private Sub MarkFailed(detail As String)
    mReadFailed = True
    If Len(mFailDetail) = 0 Then mFailDetail = detail
End Sub

' Máximo de pontos cujo rótulo é lido numa passada (em Test, Util_LabelBudgetForTest > 0 o substitui).
Private Function PointBudget() As Long
    PointBudget = MAX_POINTS_PER_PASS
    If BUILD_CONFIG = "Test" Then
        If Util_LabelBudgetForTest > 0 Then PointBudget = Util_LabelBudgetForTest
    End If
End Function

' Gancho de teste: simula falha de leitura no ponto indicado (READFAIL_*). Inerte fora do build Test.
Private Sub TestFailSite(site As Long)
    If BUILD_CONFIG = "Test" Then
        If Util_FailChartReadSiteForTest = site Or (site = READFAIL_CHART And Util_FailChartReadForTest) Then
            Err.Raise vbObjectError + 9998, "ThinQQ", "Falha simulada de leitura (teste), ponto " & site
        End If
    End If
End Sub

' Altura (em linhas) do maior bloco registrado na planilha interna: limite de quantos pontos de uma série podem ter
' rótulo ligado a um bloco.
Private Function MaxBlockRows(wb As Workbook) As Long
    Dim nm As Name, rng As Range
    For Each nm In wb.Names
        If Left$(nm.Name, Len(ID_PREFIX)) = ID_PREFIX Then
            Set rng = Nothing
            On Error Resume Next
            Set rng = nm.RefersToRange
            On Error GoTo 0
            If Not rng Is Nothing Then
                If rng.Worksheet.Name = CALC_SHEET Then
                    If rng.Rows.Count > MaxBlockRows Then MaxBlockRows = rng.Rows.Count
                End If
            End If
        End If
    Next
End Function

' O último CalcSheet_Cleanup não pôde ler algum gráfico (ou excedeu o orçamento de leitura) e, por isso, não
' removeu nenhum bloco?
Public Function CalcSheet_CleanupReadFailed() As Boolean
    CalcSheet_CleanupReadFailed = mReadFailed
End Function

' Faixas da planilha _ThinQQ citadas no texto de fórmulas (referências de outras pastas, "[x]_ThinQQ", são ignoradas).
' Uma citação "_ThinQQ!..." que não pode ser interpretada como endereço marca mReadFailed (falha fechada) em vez de
' ser descartada em silêncio.
Private Function ReferencedCalcRanges(txt As String, calc As Worksheet) As Collection
    Dim refs As New Collection, p As Long, q As Long, addr As String, ch As String, rg As Range
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
            Set rg = Nothing
            If Len(addr) > 0 Then
                On Error Resume Next
                Set rg = calc.Range(addr)
                On Error GoTo 0
            End If
            If rg Is Nothing Then MarkFailed "referência à planilha interna não interpretável: '" & addr & "'" Else refs.Add rg
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
' Um bloco está vivo se existe um gráfico nomeado "tqq_<id>" OU se alguma série (Series.Formula) ou algum
' rótulo ligado (DataLabel.Formula) de qualquer gráfico da pasta referencia uma célula do bloco (cópias, gráficos
' movidos para planilha de gráfico e 100% sem totais, que só referencia o bloco pelos rótulos ligados).
' Falha fechada: se a leitura de qualquer gráfico falhar, não remove nada (devolve 0, registra em log e
' CalcSheet_CleanupReadFailed passa a ser True até a próxima limpeza).
Public Function CalcSheet_Cleanup(wb As Workbook) As Long
    Dim i As Long, nm As Name, id As String, removed As Long, rng As Range
    Dim refs As Collection, haveRefs As Boolean, txt As String
    mReadFailed = False
    mFailDetail = ""
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
                                txt = AllChartReferences(wb, MaxBlockRows(wb))
                                If Not mReadFailed Then Set refs = ReferencedCalcRanges(txt, rng.Worksheet)
                                If mReadFailed Then
                                    Errors_Log "CalcSheet_Cleanup: leitura de gráfico falhou; nenhum bloco removido [" & mFailDetail & "]"
                                    Exit Function
                                End If
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

' Coluna do corpo (1..Columns.Count) cujo cabecalho (linha 1 do corpo) e exatamente header; 0 se nao existe. As familias usam esta
' busca em vez da posicao: a Fase 3 acrescenta colunas A DIREITA do bloco sob demanda (Block_Extend), e "as duas ultimas colunas" deixam
' de ser Pos/VLine. header "" devolve 0 (casaria a primeira coluna em branco); um cabecalho com erro (#REF!) nao casa nada e nao interrompe
' a busca (CStr de um erro levantaria 13).
Public Function Block_FindCol(blk As Range, header As String) As Long
    Dim j As Long, hdr As Range, v As Variant
    If blk Is Nothing Then Exit Function
    If Len(header) = 0 Then Exit Function
    If blk.Rows.Count < 2 Then Exit Function
    Set hdr = blk.Rows(2)
    For j = 1 To hdr.Cells.Count
        v = hdr.Cells(1, j).Value
        If Not IsError(v) Then
            If CStr(v) = header Then Block_FindCol = j: Exit Function
        End If
    Next
End Function

' Acrescenta extraCols colunas vazias a direita do bloco tqq_<id> (os blocos se empilham a partir da coluna A: a area a direita e livre) e
' reescreve o nome para cobri-las. Devolve o bloco novo. Recusa (6014) se alguma celula a direita estiver ocupada: nunca sobrescreve.
Public Function Block_Extend(wb As Workbook, id As String, extraCols As Long) As Range
    Dim blk As Range, ext As Range, newBlk As Range
    Set blk = CalcSheet_FindBlock(wb, id)
    If blk Is Nothing Then Err.Raise vbObjectError + 6015, "Block_Extend", "Bloco inexistente: " & id
    If extraCols < 1 Then Set Block_Extend = blk: Exit Function
    Set ext = blk.Offset(0, blk.Columns.Count).Resize(blk.Rows.Count, extraCols)
    If Application.WorksheetFunction.CountA(ext) > 0 Then Err.Raise vbObjectError + 6014, "Block_Extend", "Há dados à direita do bloco " & id & "; o bloco não pode ser estendido."
    ext.Clear
    Set newBlk = blk.Resize(blk.Rows.Count, blk.Columns.Count + extraCols)
    wb.Names(ID_PREFIX & id).RefersTo = "='" & CALC_SHEET & "'!" & newBlk.Address(True, True)
    Set Block_Extend = newBlk
End Function

' Tabela de origem do grafico (endereco gravado na meta, coluna C, sem o "="), resolvida na pasta do grafico; Nothing se nao resolver.
Public Function Block_SourceRange(cht As Chart) As Range
    Dim blk As Range, a As String, bang As Long, sh As String
    On Error GoTo Fail
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    a = CStr(blk.Cells(1, 3).Value)
    bang = InStrRev(a, "!")
    If bang < 2 Then Exit Function
    sh = Left$(a, bang - 1)
    If Left$(sh, 1) = "'" And Right$(sh, 1) = "'" Then sh = Replace(Mid$(sh, 2, Len(sh) - 2), "''", "'")
    Set Block_SourceRange = Registry_Workbook(cht).Worksheets(sh).Range(Mid$(a, bang + 1))
    Exit Function
Fail:
    Set Block_SourceRange = Nothing
End Function