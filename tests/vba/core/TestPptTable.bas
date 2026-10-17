Attribute VB_Name = "TestPptTable"
Option Explicit
' 1.10.0: tabela do PowerPoint vinculada a intervalo (spec 2026-10-15 secao 5.4). Partes puras.

Private Function M2(a As Variant, rows As Long, cols As Long) As Variant
    Dim m() As String, r As Long, c As Long
    ReDim m(1 To rows, 1 To cols)
    For r = 1 To rows
        For c = 1 To cols
            m(r, c) = CStr(a((r - 1) * cols + c - 1))
        Next
    Next
    M2 = m
End Function

Public Sub Test_PptTable_SerializeParseRoundTrip()
    Dim m As Variant, p As Variant
    m = M2(Array("Ano", "Sul", "T1", "1.234,5"), 2, 2)
    p = PptTable_Parse(PptTable_Serialize(m))
    AssertEqual 2, UBound(p, 1): AssertEqual 2, UBound(p, 2)
    AssertEqual "1.234,5", p(2, 2)
    AssertTrue IsEmpty(PptTable_Parse("")), "vazio"
    m = M2(Array("", ""), 1, 2)
    p = PptTable_Parse(PptTable_Serialize(m))
    AssertEqual 2, UBound(p, 2), "celulas vazias contam"
End Sub

' Adendo do controlador: a etiqueta TQQ_TBLAST so tem caracteres imprimiveis e LF (o .pptx salvo perde Chr(30)/Chr(31) e controles, troca Tab
' por espaco e CR por LF); linhas por LF, celulas por "|", escape reversivel (como TQQ_TPRE/TSUF/TLAST): ida e volta sem perda.
Public Sub Test_PptTable_Serialize_EscapesSeparators()
    Dim m As Variant, p As Variant, s As String, i As Long, r As Long, c As Long, ch As Long
    m = M2(Array("a|b", "c" & vbLf & "d", "\x7C\", "e" & vbCr & vbTab & Chr$(11) & "f", Chr$(30) & Chr$(31) & "|", "", "", "|"), 4, 2)
    s = PptTable_Serialize(m)
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        AssertTrue ch = 10 Or (ch >= 32 And ch <> 127), "caractere de controle na etiqueta (pos " & i & ")"
    Next
    AssertEqual 3, UBound(Split(s, vbLf)), "4 linhas: 3 LF"
    AssertEqual 1, UBound(Split(Split(s, vbLf)(0), "|")), "2 celulas na 1a linha: um | so"
    p = PptTable_Parse(s)
    AssertEqual 4, UBound(p, 1): AssertEqual 2, UBound(p, 2)
    For r = 1 To 4
        For c = 1 To 2
            AssertEqual m(r, c), p(r, c), "celula " & r & "," & c
        Next
    Next
End Sub

' A etiqueta vem da apresentacao (de quem a fez): lixo nunca levanta erro; maior que os limites = sem historico (Empty).
Public Sub Test_PptTable_Parse_HostileNeverRaises()
    Dim p As Variant
    p = PptTable_Parse("a|b" & vbLf & "c")
    AssertEqual 2, UBound(p, 1): AssertEqual 2, UBound(p, 2)
    AssertEqual "", p(2, 2), "linha curta: celula vazia"
    p = PptTable_Parse("\x|\" & vbLf & "\q")
    AssertEqual "\x", p(1, 1)
    AssertTrue IsEmpty(PptTable_Parse(String$(PPTTABLE_MAX_ROWS + 1, vbLf))), "linhas demais"
    AssertTrue IsEmpty(PptTable_Parse(String$(PPTTABLE_MAX_COLS + 1, "|"))), "colunas demais"
    AssertTrue IsEmpty(PptTable_Parse(String$(PPTTABLE_MAX_TAG + 1, "x"))), "etiqueta grande demais"
End Sub

Public Sub Test_PptTable_Diff()
    Dim a As Variant, b As Variant, d As Collection
    a = M2(Array("x", "1", "y", "2"), 2, 2)
    b = M2(Array("x", "1", "y", "3"), 2, 2)
    Set d = PptTable_Diff(a, b)
    AssertEqual 1, d.Count: AssertEqual 2, d(1)(0): AssertEqual 2, d(1)(1)
    b = M2(Array("x", "1", "y", "2", "z", "5"), 3, 2)
    Set d = PptTable_Diff(a, b)
    AssertEqual 2, d.Count, "linha nova: as duas celulas"
    Set d = PptTable_Diff(Empty, a)
    AssertEqual 4, d.Count, "sem historico: tudo"
    Set d = PptTable_Diff(a, a)
    AssertEqual 0, d.Count
End Sub

Public Sub Test_PptTable_TextMatrixAndRefusals()
    Dim wb As Workbook, ws As Worksheet, m As Variant
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B2").Value = 1234.5: ws.Range("A1:B2").NumberFormat = "#,##0.0"
    m = PptTable_TextMatrix(ws.Range("A1:B2"))
    AssertEqual ws.Range("A1").Text, m(1, 1)
    AssertEqual "", PptTable_RefusalText(ws.Range("A1:B2"))
    ws.Range("D1:E1").Merge
    AssertTrue Len(PptTable_RefusalText(ws.Range("D1:E2"))) > 0, "mescladas recusadas"
    AssertTrue Len(PptTable_RefusalText(ws.Range("A1:A80"))) > 0, "mais de 75 linhas"
    AssertTrue Len(PptTable_RefusalText(ws.Range("A1:Z1"))) > 0, "mais de 25 colunas"
    AssertTrue Len(PptTable_RefusalText(ws.Range("A:A"))) > 0, "coluna inteira"
    AssertTrue Len(PptTable_RefusalText(ws.Range("A1:A2,C1:C2"))) > 0, "varias areas"
    ws.Range("D1:E1").UnMerge
    AssertEqual "", PptTable_RefusalText(ws.Range("A1:Y75")), "75 x 25 cabe"
    CloseTemp wb
End Sub

' Texto da celula para a tabela: quebra de linha da celula (Alt+Enter) vira a quebra de linha do PowerPoint (Chr(11)); outros controles viram
' espaco; cortado em PPTTEXT_MAX_LEN sem partir par substituto.
Public Sub Test_PptTable_SafeCell()
    Dim s As String, emo As String
    AssertEqual "a" & Chr$(11) & "b" & Chr$(11) & "c d", PptTable_SafeCell("a" & vbCrLf & "b" & vbLf & "c" & vbTab & "d")
    AssertEqual "x y", PptTable_SafeCell("x" & Chr$(30) & "y")
    AssertEqual PPTTEXT_MAX_LEN, Len(PptTable_SafeCell(String$(300, "z")))
    emo = ChrW$(&HD83D) & ChrW$(&HDE00)
    s = PptTable_SafeCell(String$(PPTTEXT_MAX_LEN - 1, "a") & emo)
    AssertEqual PPTTEXT_MAX_LEN - 1, Len(s), "nao parte o emoji"
End Sub

' ---- F2: seguranca e roteamento com dubles (sem PowerPoint) ----

Private Function SavedBook(fileName As String) As Workbook
    Dim wb As Workbook
    Set wb = NewTempWorkbook()
    wb.Worksheets(1).Name = "Dados"
    wb.SaveAs TempDir() & "\" & fileName, xlOpenXMLWorkbook
    Set SavedBook = wb
End Function

' Forma "tabela vinculada" de um deck hostil (duble sem Table): etiquetas forjadas por quem fez o deck.
Private Function HostileTable(nm As String, ref As String, book As String) As TestPptTextFake
    Dim t As TestPptTextFake
    Set t = New TestPptTextFake
    t.HasTextFrame = 0
    t.HasTable = 0                                       ' a conferencia HasTable (nao o erro 438) recusa a forma
    Set t.Tags = New TestPptTags
    t.Tags.Preset PPTTABLE_TAG, "1"
    t.Tags.Preset PPTSYNC_TAG_BOOK, book
    t.Tags.Preset PPTLINK_TAG_NAME, nm
    t.Tags.Preset PPTLINK_TAG_REF, ref
    Set HostileTable = t
End Function

' Seguranca (regra da revisao final E, I1, para tabelas): um deck hostil so faz a sincronizacao LER um nome TQQ_T_ criado pela acao de
' inserir; nome do usuario, nome TQQ_X_ (do texto) e endereco de reserva numa pasta sem a marca sao recusados sem tocar a forma. Com a marca,
' a reserva resolve, mas uma forma que nao e tabela falha sem gravar nada.
Public Sub Test_PptTable_Sync_HostileTagsRefused_Fake()
    Dim wb As Workbook, ws As Worksheet, t As TestPptTextFake
    Set wb = SavedBook("ppttable-hostil.xlsx"): Set ws = wb.Worksheets(1)
    ws.Range("A1:B2").Value = "segredo"
    wb.Names.Add Name:="Total_Salarios", RefersTo:="='Dados'!$A$1:$B$2"
    AssertEqual "table", PptSync_ItemKind(HostileTable("", "", wb.FullName))
    Set t = HostileTable("Total_Salarios", "", wb.FullName)
    AssertEqual PPTSYNC_R_NOCHART, PptTable_Sync(t, wb, True), "nome do usuario"
    Set t = HostileTable("", "Dados|$A$1:$B$2", wb.FullName)
    AssertEqual PPTSYNC_R_NOCHART, PptTable_Sync(t, wb, True), "reserva numa pasta sem vinculos do ThinQQ"
    Set t = HostileTable(PptLink_DefineName(ws.Range("A1:B2"), "TQQ_X_"), "", wb.FullName)
    AssertEqual PPTSYNC_R_NOCHART, PptTable_Sync(t, wb, True), "nome TQQ_X_ (texto vinculado) nao vale para tabela"
    AssertTrue PptLink_HasMark(wb), "premissa: a marca"
    Set t = HostileTable("Total_Salarios", "", wb.FullName)
    AssertEqual PPTSYNC_R_NOCHART, PptTable_Sync(t, wb, True), "nome do usuario, mesmo com a marca"
    Set t = HostileTable("", "Dados|$A$1:$B$2", wb.FullName)
    AssertEqual PPTSYNC_R_FAIL, PptTable_Sync(t, wb, True), "com a marca a reserva resolve; a forma nao e tabela: falha"
    AssertEqual 0, t.TextWrites, "nada gravado"
    AssertEqual "", CStr(t.Tags.Item(PPTTABLE_TAG_LAST)), "etiqueta intacta"
    AssertTrue InStr(PptSync_SummaryText(0, 0, "", 0, 0, 0, 0, 0, 1), "Intervalo da tabela vinculada não encontrado") > 0, "resumo proprio"
    CloseTemp wb
End Sub

' Editar dados e os comandos do grafico numa tabela vinculada: recusa propria, e a pasta do usuario (fechada) nunca e aberta.
Public Sub Test_PptTable_EditDataAndCommand_RefuseTable()
    Dim wb As Workbook, p As String, t As TestPptTextFake, sl As TestPptFake, slides As TestPptFake, pres As TestPptFake, e As String, n0 As Long
    Set wb = SavedBook("ppttable-editdata.xlsx")
    p = wb.FullName
    CloseTemp wb
    Set t = HostileTable("TQQ_T_x", "", p)
    Set sl = New TestPptFake: Set sl.Shapes = New Collection: sl.Shapes.Add t, "tqq_tab"
    Set slides = New TestPptFake: slides.SlideIdValue = 256: Set slides.Slide = sl
    Set pres = New TestPptFake: Set pres.Slides = slides
    e = PptIdx_Entry(256, "tqq_tab", p)
    n0 = Application.Workbooks.Count
    AssertEqual "!" & PPTTABLE_MSG_EDITDATA, Core_PptEditData(pres, e)
    AssertEqual PPTTABLE_MSG_NOTCHART, Core_PptCommand(pres, e, "clear", "")
    AssertEqual n0, Application.Workbooks.Count, "nenhuma pasta aberta"
End Sub

' ---- Revisao final da Parte F ----

' I1: ida e volta de uma tabela de UMA coluna com a 1a celula vazia e da 1 x 1 vazia (antes viravam "sem historico": toda atualizacao sem
' forcar reescrevia as celulas editadas no PowerPoint). A etiqueta da 1 x 1 vazia nao e "" (o mesmo que sem etiqueta).
Public Sub Test_PptTable_RoundTrip_OneColumnAndEmpty()
    Dim m As Variant, p As Variant, s As String
    m = M2(Array("", "x"), 2, 1)
    p = PptTable_Parse(PptTable_Serialize(m))
    AssertFalse IsEmpty(p), "1 coluna, 1a celula vazia"
    AssertEqual 2, UBound(p, 1): AssertEqual 1, UBound(p, 2)
    AssertEqual "", p(1, 1): AssertEqual "x", p(2, 1)
    AssertEqual 0, PptTable_Diff(p, m).Count, "nada a escrever"
    m = M2(Array(""), 1, 1)
    s = PptTable_Serialize(m)
    AssertTrue Len(s) > 0, "1 x 1 vazia: etiqueta nao vazia"
    p = PptTable_Parse(s)
    AssertFalse IsEmpty(p), "1 x 1 vazia"
    AssertEqual 1, UBound(p, 1): AssertEqual 1, UBound(p, 2): AssertEqual "", p(1, 1)
    AssertEqual 0, PptTable_Diff(p, m).Count
    m = M2(Array("", "", ""), 3, 1)
    p = PptTable_Parse(PptTable_Serialize(m))
    AssertEqual 3, UBound(p, 1): AssertEqual 1, UBound(p, 2)
    m = M2(Array("\-"), 1, 1)
    p = PptTable_Parse(PptTable_Serialize(m))
    AssertEqual "\-", p(1, 1), "o texto \- nao se confunde com o marcador"
    AssertTrue IsEmpty(PptTable_Parse("")), "sem etiqueta: sem historico"
End Sub

' I2: etiqueta hostil grande nunca congela o Excel: celula codificada maior que o maximo legitimo (4 x 255) = Empty; o resto decodifica em
' tempo linear (199.999 caracteres bem abaixo de 1 s).
Public Sub Test_PptTable_Parse_HugeHostileTagIsCheap()
    Dim t0 As Single, p As Variant, s As String, row As String, i As Long, d As String
    t0 = Timer
    AssertTrue IsEmpty(PptTable_Parse(String$(PPTTABLE_MAX_TAG - 1, "a"))), "uma celula de 199.999 caracteres"
    AssertTrue IsEmpty(PptTable_Parse(String$(5000, "a"))), "uma celula de 5000 caracteres"
    AssertTrue Timer - t0 < 1, "recusa rapida: " & Format$(Timer - t0, "0.00") & " s"
    p = PptTable_Parse(String$(PPTTABLE_MAX_CELLENC, "b"))
    AssertEqual PPTTABLE_MAX_CELLENC, Len(p(1, 1)), "celula de 1020 caracteres decodifica"
    row = String$(PPTTABLE_MAX_CELLENC / 2, "\")                ' "\\" x 510: o pior caso legitimo de escapes
    row = Replace(row, "\", "\\")
    s = row
    For i = 2 To PPTTABLE_MAX_COLS
        s = s & "|" & row
    Next
    d = s
    For i = 2 To 7
        d = d & vbLf & s
    Next
    AssertTrue Len(d) < PPTTABLE_MAX_TAG, "premissa: abaixo do teto total (" & Len(d) & ")"
    t0 = Timer
    p = PptTable_Parse(d)
    AssertTrue Timer - t0 < 1, "175 celulas cheias de escapes: " & Format$(Timer - t0, "0.00") & " s"
    AssertEqual 7, UBound(p, 1): AssertEqual 25, UBound(p, 2)
    AssertEqual String$(510, "\"), p(7, 25)
    t0 = Timer
    s = PptText_TagDec(Replace(String$(66666, "a"), "a", "\x41"))
    AssertTrue Timer - t0 < 1, "TagDec de 266.664 caracteres: " & Format$(Timer - t0, "0.00") & " s"
    AssertEqual String$(66666, "A"), s, "TagDec linear da o mesmo resultado"
End Sub