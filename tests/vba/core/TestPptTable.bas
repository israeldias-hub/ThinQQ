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
