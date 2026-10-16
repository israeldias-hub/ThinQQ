Attribute VB_Name = "TestPptText"
Option Explicit
' 1.9.0: texto vinculado a celula (spec 2026-10-15 secao 5.3). Partes puras, o modelo de nome definido e o vinculo/sincronizacao com um duble
' de caixa de texto (TestPptTextFake; sem PowerPoint).

Public Sub Test_PptText_Locate_Cases()
    Dim st As Long, ln As Long
    AssertTrue PptText_Locate("Receita de 123 mil", "Receita de ", " mil", "123", st, ln)
    AssertEqual 12, st: AssertEqual 3, ln
    AssertTrue PptText_Locate("Receita de 99999 mil", "Receita de ", " mil", "123", st, ln), "usuario editou so o campo"
    AssertEqual 12, st: AssertEqual 5, ln
    AssertFalse PptText_Locate("Lucro de 123 mil", "Receita de ", " mil", "123", st, ln), "editou o prefixo: nao tocar"
    AssertFalse PptText_Locate("x", "abc", "def", "", st, ln), "texto menor que prefixo+sufixo"
    AssertTrue PptText_Locate("123", "", "", "123", st, ln), "forma inteira"
    AssertEqual 1, st: AssertEqual 3, ln
    AssertTrue PptText_Locate("ab", "ab", "", "", st, ln), "campo vazio no fim"
    AssertEqual 3, st: AssertEqual 0, ln
End Sub

Public Sub Test_PptLink_NameFollowsRowsAndRef()
    Dim wb As Workbook, ws As Worksheet, nm As String, r As Range, wide As String, wb2 As Workbook, ws2 As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("E1").Value = 1234.5: ws.Range("E1").NumberFormat = "#,##0.0"
    ws.Columns(5).ColumnWidth = 20
    nm = PptLink_DefineName(ws.Range("E1"), "TQQ_X_")
    AssertTrue Left$(nm, 6) = "TQQ_X_"
    AssertFalse wb.Names(nm).Visible, "nome oculto"
    ws.Rows(1).Insert
    Set r = PptLink_ResolveRange(wb, nm, "")
    AssertEqual "$E$2", r.Address, "o nome acompanha a linha inserida"
    AssertEqual "Dados|$E$2", PptLink_RefOf(ws.Range("E2")), "reserva: planilha|endereco (sem Chr(31): nao sobrevive ao .pptx salvo)"
    Set r = PptLink_ResolveRange(wb, "TQQ_X_naoexiste", PptLink_RefOf(ws.Range("E2")))
    AssertEqual "$E$2", r.Address, "reserva pelo endereco"
    AssertTrue PptLink_ResolveRange(wb, "TQQ_X_naoexiste", "") Is Nothing
    AssertTrue PptLink_ResolveRange(wb, "", "|") Is Nothing, "reserva sem planilha"
    AssertTrue PptLink_ResolveRange(wb, "", "Dados|") Is Nothing, "reserva sem endereco"
    AssertTrue PptLink_ResolveRange(wb, "", "Nada|$A$1") Is Nothing, "planilha inexistente"
    AssertTrue PptLink_ResolveRange(wb, "", "Dados|$$!!") Is Nothing, "endereco invalido"
    wide = ws.Range("E2").Text
    AssertEqual wide, PptText_CellText(ws.Range("E2")), "o texto exibido (formato do Excel)"
    ws.Columns(5).ColumnWidth = 1
    AssertTrue InStr(PptText_CellText(ws.Range("E2")), "#") = 0, "coluna estreita nao vira ####"
    AssertEqual wide, PptText_CellText(ws.Range("E2")), "coluna estreita: o mesmo texto da coluna larga (" & ws.Range("E2").Text & ")"
    ws.Range("G5").Value = 45000: ws.Range("G5").NumberFormat = "dd/mm/yyyy": ws.Columns(7).ColumnWidth = 20
    wide = ws.Range("G5").Text
    ws.Columns(7).ColumnWidth = 1
    AssertEqual wide, PptText_CellText(ws.Range("G5")), "data em coluna estreita"
    ws.Range("H1").Value = 0.256: ws.Range("H1").NumberFormat = "0.0%": ws.Columns(8).ColumnWidth = 20
    wide = ws.Range("H1").Text
    ws.Columns(8).ColumnWidth = 1
    AssertEqual wide, PptText_CellText(ws.Range("H1")), "porcentagem em coluna estreita"
    ws.Range("I1").Value = -9876.25: ws.Range("I1").NumberFormat = "[Blue]#,##0.00;[Red]-#,##0.00": ws.Columns(9).ColumnWidth = 20
    wide = ws.Range("I1").Text
    ws.Columns(9).ColumnWidth = 1
    AssertEqual wide, PptText_CellText(ws.Range("I1")), "formato com cor em coluna estreita"
    AssertEqual "1.5", PptText_FormatValue(CVar("1.5"), "#,##0.0"), "texto fica como esta"
    AssertEqual "", PptText_FormatValue(CVar(1), "[Qualquer"), "formato invalido: vazio (quem chama usa o valor cru)"
    ' planilha com "|" no nome: a reserva divide pelo ULTIMO "|"
    Set ws2 = wb.Worksheets.Add: ws2.Name = "a|b"
    ws2.Range("C3").Value = "x"
    Set r = PptLink_ResolveRange(wb, "", PptLink_RefOf(ws2.Range("C3")))
    AssertTrue Not r Is Nothing, "planilha com | no nome"
    AssertEqual "x", CStr(r.Value)
    ' nome que aponta para OUTRA pasta nunca e seguido
    Set wb2 = NewTempWorkbook()
    wb.Names.Add Name:="TQQ_X_fora", RefersTo:="=[" & wb2.Name & "]" & wb2.Worksheets(1).Name & "!$A$1", Visible:=False
    AssertTrue PptLink_ResolveRange(wb, "TQQ_X_fora", "") Is Nothing, "nome para outra pasta"
    CloseTemp wb2
    CloseTemp wb
End Sub

Public Sub Test_PptAuto_IsCandidate_ByLinkName()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    wb.SaveAs TempDir() & "\ppttext-cand.xlsx", xlOpenXMLWorkbook
    AssertFalse PptAuto_IsCandidate(wb), "sem _ThinQQ e sem nomes"
    PptLink_DefineName ws.Range("A1"), "TQQ_X_"
    AssertTrue PptAuto_IsCandidate(wb), "texto vinculado"
    CloseTemp wb
End Sub

' O que vai para as etiquetas do PowerPoint so tem caracteres imprimiveis (o .pptx salvo perde Chr(30)/Chr(31)/controles, troca Tab por
' espaco e CR por LF) e volta igual.
Public Sub Test_PptText_TagEncDec_RoundTrip()
    Dim s As String, e As String, i As Long
    s = "a\b" & vbCr & "c" & vbLf & "d" & vbTab & "e" & Chr$(11) & "f" & Chr$(30) & Chr$(31) & Chr$(0) & Chr$(127) & "|\x41\n"
    e = PptText_TagEnc(s)
    For i = 1 To Len(e)
        AssertTrue AscW(Mid$(e, i, 1)) >= 32 And AscW(Mid$(e, i, 1)) <> 127, "caractere de controle na etiqueta (pos " & i & ")"
    Next
    AssertEqual s, PptText_TagDec(e), "ida e volta"
    AssertEqual "", PptText_TagEnc("")
    AssertEqual "", PptText_TagDec("")
    AssertEqual "Receita de ", PptText_TagDec(PptText_TagEnc("Receita de "))
    AssertEqual "\", PptText_TagDec("\"), "barra solta no fim"
    AssertEqual "\q", PptText_TagDec("\q"), "sequencia desconhecida fica"
    AssertEqual "\xZ", PptText_TagDec("\xZ"), "\x truncado fica"
    AssertEqual "\x4", PptText_TagDec("\x4"), "\x com um digito fica"
End Sub

Public Sub Test_PptText_SafeValue_ControlAndCap()
    Dim cut As Boolean, s As String
    AssertEqual "a b c d e", PptText_SafeValue("a" & vbCrLf & "b" & vbLf & "c" & vbTab & "d" & Chr$(7) & "e", cut)
    AssertFalse cut
    s = PptText_SafeValue(String$(300, "x"), cut)
    AssertTrue cut, "cortado"
    AssertEqual PPTTEXT_MAX_LEN, Len(s)
    s = PptText_SafeValue(String$(255, "y"), cut)
    AssertFalse cut, "255 cabe"
    AssertEqual 255, Len(s)
End Sub

' Resultado de formula e erro: o texto exibido.
Public Sub Test_PptText_CellText_FormulaAndError()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1").Formula = "=2+3"
    ws.Range("A2").Formula = "=1/0"
    ws.Range("A3").Value = "texto"
    AssertEqual "5", PptText_CellText(ws.Range("A1"))
    AssertEqual ws.Range("A2").Text, PptText_CellText(ws.Range("A2")), "erro: o texto do erro"
    AssertTrue Left$(PptText_CellText(ws.Range("A2")), 1) = "#", "erro comeca com #"
    AssertEqual "texto", PptText_CellText(ws.Range("A3:B4")), "intervalo: a 1a celula"
    AssertEqual "", PptText_CellText(ws.Range("C9")), "vazia"
    CloseTemp wb
End Sub

Private Function TextShape(txt As String) As TestPptTextFake
    Dim t As TestPptTextFake
    Set t = New TestPptTextFake
    t.HasTextFrame = -1
    Set t.Tags = New TestPptTags
    t.Preset txt
    Set TextShape = t
End Function

Private Function SavedBook(fileName As String) As Workbook
    Dim wb As Workbook
    Set wb = NewTempWorkbook()
    wb.Worksheets(1).Name = "Dados"
    wb.SaveAs TempDir() & "\" & fileName, xlOpenXMLWorkbook
    Set SavedBook = wb
End Function

' Recusas do vincular: nada muda (nem o texto, nem as etiquetas, nem os nomes da pasta).
Public Sub Test_PptText_Link_RefusalsChangeNothing()
    Dim wb As Workbook, t As TestPptTextFake, n0 As Long, p As String
    Set t = TextShape("Total: X")
    Set wb = NewTempWorkbook()
    n0 = wb.Names.Count
    AssertEqual PPTTEXT_MSG_UNSAVED, PptText_Link(t, 8, 1, wb.Worksheets(1).Range("A1")), "pasta nao salva"
    AssertEqual "Total: X", t.Text
    AssertEqual n0, wb.Names.Count
    AssertEqual 0, t.Tags.Writes
    p = TempDir() & "\ppttext.csv"
    Application.DisplayAlerts = False
    wb.SaveAs p, 6                                                  ' xlCSV
    Application.DisplayAlerts = True
    AssertEqual PPTTEXT_MSG_NOTBOOK, PptText_Link(t, 8, 1, wb.Worksheets(1).Range("A1")), "csv nao e pasta do Excel"
    AssertEqual "Total: X", t.Text
    CloseTemp wb
    Set wb = SavedBook("ppttext-ref.xlsx")
    n0 = wb.Names.Count
    t.HasTextFrame = 0
    AssertEqual PPTTEXT_MSG_NOTEXT, PptText_Link(t, 8, 1, wb.Worksheets(1).Range("A1")), "sem texto"
    AssertEqual PPTTEXT_MSG_NOTEXT, PptText_Link(Nothing, 8, 1, wb.Worksheets(1).Range("A1")), "sem forma"
    AssertEqual n0, wb.Names.Count, "nenhum nome criado"
    AssertEqual 0, t.Tags.Writes
    AssertTrue wb.Saved, "pasta intacta"
    CloseTemp wb
End Sub

' Falha no PowerPoint depois do nome definido: o nome sai, o texto volta, a forma nao vira "texto vinculado" e o erro sobe.
Public Sub Test_PptText_Link_FailureUndone()
    Dim wb As Workbook, t As TestPptTextFake, n0 As Long, en As Long
    Set wb = SavedBook("ppttext-undo.xlsx")
    wb.Worksheets(1).Range("A1").Value = 10
    Set t = TextShape("Total: X")
    n0 = wb.Names.Count
    t.FailTextWrite = True
    On Error Resume Next
    PptText_Link t, 8, 1, wb.Worksheets(1).Range("A1")
    en = Err.Number
    On Error GoTo 0
    AssertEqual vbObjectError + 9993, en, "o erro do PowerPoint sobe"
    AssertEqual n0, wb.Names.Count, "nome desfeito"
    AssertEqual "Total: X", t.Text, "texto intacto"
    AssertEqual "", t.Tags.Item(PPTTEXT_TAG), "nao vinculado"
    CloseTemp wb
End Sub

' Vincular e sincronizar com o duble: forma inteira, parte e cursor; sem mudanca nao reescreve; a pasta so e LIDA pela sincronizacao (nem nome
' novo nem "nao salva"); edicao fora do campo nao e tocada; nome apagado cai na reserva; celula perdida = R_NOCHART.
Public Sub Test_PptText_LinkAndSync_Fake()
    Dim wb As Workbook, ws As Worksheet, t As TestPptTextFake, nNames As Long, w0 As Long, pre As String, raw As String, i As Long
    Set wb = SavedBook("ppttext-fake.xlsx"): Set ws = wb.Worksheets(1)
    ws.Range("B2").Value = 123
    ' parte do texto, com quebra de paragrafo (CR), quebra de linha (Chr(11)) e Tab no prefixo
    pre = "Linha 1" & vbCr & "Receita" & Chr$(11) & "de" & vbTab
    Set t = TextShape(pre & "X mil")
    AssertEqual "", PptText_Link(t, Len(pre) + 1, 1, ws.Range("B2"))
    AssertEqual pre & "123 mil", t.Text
    AssertEqual "text", PptSync_ItemKind(t)
    AssertEqual "1", t.Tags.Item(PPTTEXT_TAG)
    AssertEqual wb.FullName, t.Tags.Item(PPTSYNC_TAG_BOOK)
    AssertEqual "Dados|$B$2", t.Tags.Item(PPTLINK_TAG_REF)
    raw = t.Tags.Item(PPTTEXT_TAG_PRE)
    For i = 1 To Len(raw)
        AssertTrue AscW(Mid$(raw, i, 1)) >= 32, "etiqueta do prefixo sem caracteres de controle"
    Next
    AssertEqual pre, PptText_TagDec(raw)
    AssertEqual " mil", PptText_TagDec(t.Tags.Item(PPTTEXT_TAG_SUF))
    AssertFalse wb.Saved, "vincular criou o nome (a pasta do usuario fica nao salva: decisao dele)"
    wb.Save
    nNames = wb.Names.Count
    ' sem mudanca: R_SAME, nada reescrito
    w0 = t.TextWrites
    AssertEqual PPTSYNC_R_SAME, PptText_Sync(t, wb, False)
    AssertEqual w0, t.TextWrites, "sem mudanca nao reescreve"
    ' forcado sem mudanca: R_OK, texto nao reescrito
    AssertEqual PPTSYNC_R_OK, PptText_Sync(t, wb, True)
    AssertEqual w0, t.TextWrites, "forcado sem mudanca nao reescreve o texto"
    ' mudou: so o campo
    ws.Range("B2").Value = 4567
    AssertEqual PPTSYNC_R_OK, PptText_Sync(t, wb, False)
    AssertEqual pre & "4567 mil", t.Text
    ' usuario editou so o campo: trocado de novo
    t.Preset pre & "???? mil"
    AssertEqual PPTSYNC_R_OK, PptText_Sync(t, wb, False)
    AssertEqual pre & "4567 mil", t.Text
    ' usuario editou fora do campo: nada tocado
    t.Preset "Outro texto"
    ws.Range("B2").Value = 1
    w0 = t.TextWrites
    AssertEqual PPTSYNC_R_EDITED, PptText_Sync(t, wb, True)
    AssertEqual "Outro texto", t.Text
    AssertEqual w0, t.TextWrites
    ' nome apagado: a reserva (endereco) acha a celula
    t.Preset pre & "4567 mil"
    wb.Names(t.Tags.Item(PPTLINK_TAG_NAME)).Delete
    wb.Save
    nNames = wb.Names.Count
    AssertEqual PPTSYNC_R_OK, PptText_Sync(t, wb, False)
    AssertEqual pre & "1 mil", t.Text, "pela reserva"
    ' celula perdida (nome e planilha da reserva): R_NOCHART, nada tocado
    t.Tags.Add PPTLINK_TAG_REF, "Sumiu|$B$2"
    AssertEqual PPTSYNC_R_NOCHART, PptText_Sync(t, wb, True)
    AssertEqual pre & "1 mil", t.Text
    AssertTrue wb.Saved, "a sincronizacao nunca grava na pasta"
    AssertEqual nNames, wb.Names.Count, "a sincronizacao nunca cria nomes"
    CloseTemp wb
End Sub

' Forma inteira e cursor (tamanho 0); valor vazio; valor longo cortado (a mensagem diz).
Public Sub Test_PptText_Link_WholeCursorEmptyAndCut()
    Dim wb As Workbook, ws As Worksheet, t As TestPptTextFake
    Set wb = SavedBook("ppttext-modos.xlsx"): Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = "abc"
    Set t = TextShape("tudo")
    AssertEqual "", PptText_Link(t, 1, 4, ws.Range("A1"))
    AssertEqual "abc", t.Text, "forma inteira"
    Set t = TextShape("Total: ")
    AssertEqual "", PptText_Link(t, 8, 0, ws.Range("A1"))
    AssertEqual "Total: abc", t.Text, "cursor no fim"
    Set t = TextShape("[] fim")
    AssertEqual "", PptText_Link(t, 2, 0, ws.Range("A1"))
    AssertEqual "[abc] fim", t.Text, "cursor no meio"
    ws.Range("A1").ClearContents
    AssertEqual PPTSYNC_R_OK, PptText_Sync(t, wb, False)
    AssertEqual "[] fim", t.Text, "celula vazia: campo vazio"
    ws.Range("A1").Value = "zz"
    AssertEqual PPTSYNC_R_OK, PptText_Sync(t, wb, False)
    AssertEqual "[zz] fim", t.Text, "campo vazio volta a ter valor"
    ws.Range("A2").Value = String$(300, "k") & vbLf & "x"
    Set t = TextShape("v=X")
    AssertEqual PPTTEXT_MSG_CUT, PptText_Link(t, 3, 1, ws.Range("A2"))
    AssertEqual "v=" & String$(255, "k"), t.Text, "cortado em 255"
    CloseTemp wb
End Sub

' Editar dados e os comandos do grafico numa forma de texto vinculado: recusa propria, e a pasta do usuario (fechada) nunca e aberta.
Public Sub Test_PptText_EditDataAndCommand_RefuseTextShape()
    Dim wb As Workbook, p As String, t As TestPptTextFake, sl As TestPptFake, slides As TestPptFake, pres As TestPptFake, e As String, n0 As Long
    Set wb = SavedBook("ppttext-editdata.xlsx")
    p = wb.FullName
    CloseTemp wb
    Set t = TextShape("v=1")
    t.Tags.Preset PPTTEXT_TAG, "1"
    t.Tags.Preset PPTSYNC_TAG_BOOK, p
    Set sl = New TestPptFake: Set sl.Shapes = New Collection: sl.Shapes.Add t, "tqq_t"
    Set slides = New TestPptFake: slides.SlideIdValue = 256: Set slides.Slide = sl
    Set pres = New TestPptFake: Set pres.Slides = slides
    e = PptIdx_Entry(256, "tqq_t", p)
    n0 = Application.Workbooks.Count
    AssertEqual "!" & PPTTEXT_MSG_EDITDATA, Core_PptEditData(pres, e)
    AssertEqual PPTTEXT_MSG_NOTCHART, Core_PptCommand(pres, e, "clear", "")
    AssertEqual n0, Application.Workbooks.Count, "nenhuma pasta aberta"
    AssertEqual "v=1", t.Text
End Sub

' Titulo ligado a uma celula (E2; sonda P6 medida aqui): acompanha a celula; Core_SafeRefresh o mantem; "Alterar dados" (grafico recriado)
' o religa, na celula de AGORA (linha inserida acima); Cancelar nao muda nada; remover tira so o titulo ligado.
Public Sub Test_LinkTitle_FollowsCellRefreshAndRebuild()
    Dim wb As Workbook, ws As Worksheet, c As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5): ws.Range("A3:C3").Value = Array("T2", 20, 6)
    ws.Range("E1").Value = "Receita por região"
    ws.Range("A1:C3").Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    AssertTrue Not c Is Nothing, "grafico inserido"
    AssertFalse c.HasTitle, "nasce sem titulo"
    Core_LinkTitle
    AssertFalse c.HasTitle, "build Test sem o gancho: nada (sem InputBox)"
    Core_TestSetTitleCell ws.Range("E1")
    Core_LinkTitle
    AssertTrue c.HasTitle
    AssertEqual "Receita por região", c.ChartTitle.Text
    AssertEqual "Dados|$E$1", Registry_GetOption(c, "titleref", ""), "formato de PptLink_RefOf"
    ws.Range("E1").Value = "Novo título"
    AssertEqual "Novo título", c.ChartTitle.Text, "acompanha a celula"
    Core_SafeRefresh c
    AssertTrue c.HasTitle, "P6: refresh mantem"
    AssertEqual "Novo título", c.ChartTitle.Text, "P6: refresh mantem o texto"
    ws.Rows(1).Insert                                   ' dados em A2:C4, titulo em E2 (a formula do titulo acompanha)
    AssertEqual "Novo título", c.ChartTitle.Text
    AssertTrue Core_RebuildFromRange(c, ws.Range("A2:C4")), "Alterar dados"
    Set c = Core_ActiveChart()                          ' o grafico foi recriado (o antigo nao existe mais)
    AssertTrue c.HasTitle, "Alterar dados religa o titulo"
    AssertEqual "Novo título", c.ChartTitle.Text
    AssertEqual "Dados|$E$2", Registry_GetOption(c, "titleref", ""), "a celula de agora (linha inserida), nao o endereco gravado"
    ws.Range("E2").Value = "Depois"
    AssertEqual "Depois", c.ChartTitle.Text, "o grafico novo acompanha a celula"
    Core_TestSetTitleCell Nothing, True
    Core_LinkTitle
    AssertEqual "Dados|$E$2", Registry_GetOption(c, "titleref", ""), "cancelar nao muda nada"
    AssertTrue c.HasTitle, "cancelar mantem o titulo"
    Core_TestSetTitleCell Nothing
    Core_LinkTitle
    AssertFalse c.HasTitle, "sem celula: remove o titulo ligado"
    AssertEqual "", Registry_GetOption(c, "titleref", "")
    AssertTrue Core_RebuildFromRange(c, ws.Range("A2:C4"))
    Set c = Core_ActiveChart()
    AssertFalse c.HasTitle, "removido: Alterar dados nao o recria"
    Watchers_Stop
    CloseTemp wb
End Sub

' Celula de outra pasta: aviso, nada muda. Planilha com ";" "=" "%" "|" "'" no nome (a opcao codifica). Titulo que o ThinQQ nao ligou
' (texto do usuario) nao e removido por "remover"; titulo ligado que o usuario apagou nao volta no "Alterar dados".
Public Sub Test_LinkTitle_OtherBookOddSheetAndUserTitle()
    Dim wb As Workbook, wb2 As Workbook, ws As Worksheet, ws2 As Worksheet, c As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5): ws.Range("A3:C3").Value = Array("T2", 20, 6)
    Set ws2 = wb.Worksheets.Add(After:=ws): ws2.Name = "a;b=c%3B|d'e"
    ws2.Range("B2").Value = "Título estranho"
    ws.Activate
    ws.Range("A1:C3").Select
    Core_Insert "clustered"
    Set c = Core_ActiveChart()
    Set wb2 = NewTempWorkbook()
    wb2.Worksheets(1).Range("A1").Value = "fora"
    wb.Activate: c.Parent.Activate                      ' a pasta nova roubou a janela: o grafico volta a ser o ativo
    AssertTrue Not Core_ActiveChart() Is Nothing, "grafico ativo"
    Core_ClearLastNotice
    Core_TestSetTitleCell wb2.Worksheets(1).Range("A1")
    Core_LinkTitle
    AssertEqual LINKTITLE_MSG_OTHERBOOK, Core_LastNotice(), "outra pasta: aviso"
    AssertFalse c.HasTitle, "outra pasta: nada muda"
    AssertEqual "", Registry_GetOption(c, "titleref", "")
    CloseTemp wb2
    Core_TestSetTitleCell ws2.Range("B2")
    Core_LinkTitle
    AssertEqual "Título estranho", c.ChartTitle.Text
    AssertEqual "a;b=c%3B|d'e|$B$2", LinkTitle_Dec(Registry_GetOption(c, "titleref", "")), "opcao codificada volta igual"
    AssertEqual "Título estranho", CStr(LinkTitle_Cell(c).Value)
    AssertTrue Core_RebuildFromRange(c, ws.Range("A1:C3"))
    Set c = Core_ActiveChart()
    AssertEqual "Título estranho", c.ChartTitle.Text, "religado na planilha de nome estranho"
    ' titulo ligado que o usuario apagou: Alterar dados nao o recria
    c.HasTitle = False
    AssertTrue LinkTitle_Cell(c) Is Nothing
    AssertTrue Core_RebuildFromRange(c, ws.Range("A1:C3"))
    Set c = Core_ActiveChart()
    AssertFalse c.HasTitle, "apagado pelo usuario: fica apagado"
    AssertEqual "", Registry_GetOption(c, "titleref", ""), "e a opcao copiada sai"
    ' titulo do usuario (nao ligado pelo ThinQQ): remover nao o toca
    c.HasTitle = True: c.ChartTitle.Text = "Meu título"
    Core_TestSetTitleCell Nothing
    Core_LinkTitle
    AssertTrue c.HasTitle, "titulo do usuario fica"
    AssertEqual "Meu título", c.ChartTitle.Text
    AssertEqual "a%25b%3Bc%3Dd", LinkTitle_Enc("a%b;c=d")
    AssertEqual "%3B;=", LinkTitle_Dec(LinkTitle_Enc("%3B;="))
    Watchers_Stop
    CloseTemp wb
End Sub

' Selecao do PowerPoint lida pelo botao do Excel (E3): texto/cursor = o trecho; UMA forma com texto = o texto inteiro; varias formas, forma
' sem texto ou outro tipo = nada.
Public Sub Test_PptText_SelectionParts()
    Dim sel As TestPptSelFake, t As TestPptTextFake, t2 As TestPptTextFake, shp As Object, st As Long, ln As Long
    Set t = TextShape("Receita: X mil")
    Set sel = New TestPptSelFake: Set sel.ShapeRange = New Collection: sel.ShapeRange.Add t
    sel.Start = 10: sel.Length = 1
    AssertTrue PptText_SelectionParts(3, sel, shp, st, ln), "texto selecionado"
    AssertTrue shp Is t
    AssertEqual 10, st: AssertEqual 1, ln
    sel.Length = 0
    AssertTrue PptText_SelectionParts(3, sel, shp, st, ln), "cursor"
    AssertEqual 0, ln
    Set shp = Nothing
    AssertTrue PptText_SelectionParts(2, sel, shp, st, ln), "a forma inteira"
    AssertEqual 1, st: AssertEqual Len("Receita: X mil"), ln
    Set t2 = TextShape("outra"): sel.ShapeRange.Add t2
    Set shp = Nothing
    AssertFalse PptText_SelectionParts(2, sel, shp, st, ln), "duas formas"
    AssertTrue shp Is Nothing
    Set sel.ShapeRange = New Collection: t.HasTextFrame = 0: sel.ShapeRange.Add t
    AssertFalse PptText_SelectionParts(2, sel, shp, st, ln), "forma sem texto"
    AssertFalse PptText_SelectionParts(1, sel, shp, st, ln), "slides selecionados"
    AssertFalse PptText_SelectionParts(0, sel, shp, st, ln), "nada selecionado"
    Set sel.ShapeRange = New Collection
    AssertFalse PptText_SelectionParts(3, sel, shp, st, ln), "selecao vazia: erro vira False"
End Sub

' Botao do Excel "Vincular celula ao texto selecionado": sem celula selecionada, sem PowerPoint (nunca abre um) e sem selecao legivel na
' janela do PowerPoint: avisos proprios, nada vinculado, operacao fechada.
Public Sub Test_Core_PptLinkTextFromPpt_Refusals()
    Dim wb As Workbook, ws As Worksheet, app As TestPptFake, win As TestPptFake, sh As Shape, n0 As Long
    Set wb = SavedBook("ppttext-fromppt.xlsx"): Set ws = wb.Worksheets(1)
    ws.Range("C4").Value = 77
    n0 = wb.Names.Count
    wb.Activate: ws.Activate
    Set sh = ws.Shapes.AddShape(1, 10, 10, 40, 40)
    sh.Select
    Core_ClearLastNotice
    Core_PptLinkTextFromPpt
    AssertEqual PPTTEXT_MSG_NOCELL, Core_LastNotice(), "sem celula selecionada"
    sh.Delete
    wb.Save
    ws.Range("C4").Select
    Ppt_ResetTestHooks
    Ppt_SetAppForTest Nothing
    Core_PptLinkTextFromPpt
    AssertEqual PPTTEXT_MSG_NOPPT, Core_LastNotice(), "sem PowerPoint aberto"
    Set app = New TestPptFake: Set win = New TestPptFake
    Set app.ActiveWindow = win                                   ' janela sem selecao (Selection Nothing: erro 91, como no PowerPoint)
    Ppt_SetAppForTest app
    Core_PptLinkTextFromPpt
    AssertEqual PPTTEXT_MSG_NOTEXT, Core_LastNotice(), "nada selecionado no PowerPoint"
    AssertEqual n0, wb.Names.Count, "nenhum nome criado"
    AssertTrue wb.Saved, "pasta intacta"
    AssertFalse Ppt_InOperation(), "operacao fechada"
    Ppt_ResetTestHooks
    CloseTemp wb
End Sub

' API 5: a celula ATIVA deste Excel; modo silencioso e operacao equilibrados; nenhuma caixa; nunca propaga.
Public Sub Test_Core_PptLinkText_ActiveCellQuietBalanced()
    Dim wb As Workbook, ws As Worksheet, t As TestPptTextFake, s As String
    Set wb = SavedBook("ppttext-api.xlsx"): Set ws = wb.Worksheets(1)
    ws.Range("C4").Value = 77
    wb.Activate: ws.Activate: ws.Range("C4").Select
    Core_ClearLastNotice
    Set t = TextShape("n=X")
    s = Core_PptLinkText(t, 3, 1)
    AssertEqual "Texto vinculado à célula ppttext-api.xlsx!Dados!C4.", s, "a pasta, a planilha e a celula (M9)"
    AssertEqual "n=77", t.Text
    AssertEqual "Dados|$C$4", t.Tags.Item(PPTLINK_TAG_REF)
    AssertFalse Core_IsQuiet(), "modo silencioso fechado"
    AssertFalse Ppt_InOperation(), "operacao fechada"
    AssertEqual PPTTEXT_MSG_NOTEXT, Core_PptLinkText(Nothing, 1, 0)
    Set t = TextShape("w")
    t.FailTextWrite = True
    s = Core_PptLinkText(t, 1, 1)
    AssertTrue InStr(s, "Não foi possível vincular o texto") = 1, s
    AssertFalse Core_IsQuiet(), "modo silencioso fechado na falha"
    AssertFalse Ppt_InOperation(), "operacao fechada na falha"
    AssertEqual 0, Core_ShownCountForTest(), "nenhuma caixa"
    CloseTemp wb
End Sub

' ---- Revisao final da Parte E ----

' Forma de texto "vinculado" de um deck hostil (duble): etiquetas forjadas por quem fez o deck.
Private Function HostileShape(nm As String, ref As String, book As String) As TestPptTextFake
    Dim t As TestPptTextFake
    Set t = TextShape("v=X")
    t.Tags.Preset PPTTEXT_TAG, "1"
    t.Tags.Preset PPTSYNC_TAG_BOOK, book
    t.Tags.Preset PPTLINK_TAG_NAME, nm
    t.Tags.Preset PPTLINK_TAG_REF, ref
    t.Tags.Preset PPTTEXT_TAG_PRE, PptText_TagEnc("v=")
    t.Tags.Preset PPTTEXT_TAG_LAST, "X"
    Set HostileShape = t
End Function

' I1 (seguranca): um deck hostil nao le celulas arbitrarias da pasta do usuario. Nome definido do usuario (nao TQQ_X_) nunca e seguido;
' endereco de reserva so numa pasta com a marca (o usuario vinculou algo nela). Nada e escrito no slide.
Public Sub Test_PptText_Sync_HostileTagsRefused_Fake()
    Dim wb As Workbook, ws As Worksheet, t As TestPptTextFake
    Set wb = SavedBook("ppttext-hostil-tags.xlsx"): Set ws = wb.Worksheets(1)
    ws.Range("B2").Value = "salario secreto"
    wb.Names.Add Name:="Total_Salarios", RefersTo:="='Dados'!$B$2"
    AssertFalse PptLink_HasMark(wb), "premissa: pasta sem a marca"
    Set t = HostileShape("Total_Salarios", "", wb.FullName)
    AssertEqual PPTSYNC_R_NOCHART, PptText_Sync(t, wb, True), "nome do usuario"
    AssertEqual "v=X", t.Text
    Set t = HostileShape("", "Dados|$B$2", wb.FullName)
    AssertEqual PPTSYNC_R_NOCHART, PptText_Sync(t, wb, True), "reserva numa pasta sem vinculos do ThinQQ"
    AssertEqual "v=X", t.Text
    Set t = HostileShape("TQQ_X_naoexiste", "Dados|$B$2", wb.FullName)
    AssertEqual PPTSYNC_R_NOCHART, PptText_Sync(t, wb, True), "nome TQQ_X_ inexistente e reserva sem a marca"
    AssertEqual "v=X", t.Text
    AssertTrue InStr(PptSync_SummaryText(0, 0, "", 0, 0, 0, 0, 1), "Célula do texto vinculado não encontrada") > 0, "M5: o resumo nao diz 'grafico'"
    ' depois que o usuario vincula algo nesta pasta (marca): a reserva vale; um nome que nao e TQQ_X_ continua recusado
    PptLink_DefineName ws.Range("C3"), "TQQ_X_"
    AssertTrue PptLink_HasMark(wb), "vincular cria a marca"
    Set t = HostileShape("Total_Salarios", "", wb.FullName)
    AssertEqual PPTSYNC_R_NOCHART, PptText_Sync(t, wb, True), "nome do usuario, mesmo com a marca"
    AssertEqual "v=X", t.Text
    Set t = HostileShape("", "Dados|$B$2", wb.FullName)
    AssertEqual PPTSYNC_R_OK, PptText_Sync(t, wb, True), "com a marca, a reserva vale"
    CloseTemp wb
End Sub

' I3 (custo): PptAuto_IsCandidate (a cada SheetChange/SheetCalculate) so faz buscas diretas: uma pasta com muitos nomes, um deles TQQ_X_
' posto sem a marca, NAO e candidata (prova que os nomes nao sao varridos); com a marca (vincular), e.
Public Sub Test_PptAuto_IsCandidate_NoNameLoop()
    Dim wb As Workbook, ws As Worksheet, i As Long, t0 As Single, k As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    wb.SaveAs TempDir() & "\ppttext-nomes.xlsx", xlOpenXMLWorkbook
    For i = 1 To 600
        wb.Names.Add Name:="Nome_" & i, RefersTo:="=" & ws.Name & "!$A$" & i, Visible:=(i Mod 2 = 0)
    Next
    wb.Names.Add Name:="TQQ_X_semmarca", RefersTo:="=" & ws.Name & "!$B$1", Visible:=False
    AssertFalse PptAuto_IsCandidate(wb), "sem a marca: os nomes nao sao varridos"
    t0 = Timer
    For k = 1 To 200
        PptAuto_IsCandidate wb
    Next
    AssertTrue Timer - t0 < 2, "200 chamadas numa pasta com 600 nomes: " & Format$(Timer - t0, "0.00") & " s"
    PptLink_DefineName ws.Range("C1"), "TQQ_X_"
    AssertTrue PptAuto_IsCandidate(wb), "com a marca"
    CloseTemp wb
End Sub

' M1/M2: vincular de novo a mesma forma (outra parte do texto, outra celula): o campo antigo vira texto comum e a sincronizacao so mexe no
' campo novo. Falha numa etiqueta ao vincular de novo: as etiquetas antigas voltam (nada misturado) e a forma continua vinculada como antes.
Public Sub Test_PptText_Relink_SameShape()
    Dim wb As Workbook, ws As Worksheet, t As TestPptTextFake, nm0 As String, ref0 As String, en As Long
    Set wb = SavedBook("ppttext-relink.xlsx"): Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = 10: ws.Range("B1").Value = "Sul"
    Set t = TextShape("X em Y")
    AssertEqual "", PptText_Link(t, 1, 1, ws.Range("A1"))
    AssertEqual "10 em Y", t.Text
    nm0 = t.Tags.Item(PPTLINK_TAG_NAME): ref0 = t.Tags.Item(PPTLINK_TAG_REF)
    ' falha na gravacao da reserva ao vincular de novo: tudo volta
    t.Tags.FailOnKey = PPTLINK_TAG_REF
    On Error Resume Next
    PptText_Link t, 7, 1, ws.Range("B1")
    en = Err.Number
    On Error GoTo 0
    AssertEqual vbObjectError + 9995, en, "o erro sobe"
    AssertEqual "10 em Y", t.Text, "texto de volta"
    AssertEqual nm0, t.Tags.Item(PPTLINK_TAG_NAME), "nome antigo de volta"
    AssertEqual ref0, t.Tags.Item(PPTLINK_TAG_REF), "reserva antiga de volta"
    AssertEqual "1", t.Tags.Item(PPTTEXT_TAG), "continua vinculada"
    AssertEqual "", PptText_TagDec(t.Tags.Item(PPTTEXT_TAG_PRE)), "prefixo antigo de volta"
    ' novo vinculo de verdade: o Y vira a celula B1; o 10 vira texto comum
    AssertEqual "", PptText_Link(t, 7, 1, ws.Range("B1"))
    AssertEqual "10 em Sul", t.Text
    ws.Range("A1").Value = 99: ws.Range("B1").Value = "Norte"
    AssertEqual PPTSYNC_R_OK, PptText_Sync(t, wb, False)
    AssertEqual "10 em Norte", t.Text, "o campo antigo virou texto comum"
    CloseTemp wb
End Sub

' M3: o corte em 255 nunca parte um par substituto (emoji); M4: forma dentro de um grupo e recusada (a sincronizacao nao desce em grupos).
Public Sub Test_PptText_SurrogateCutAndGroup()
    Dim cut As Boolean, s As String, emo As String, wb As Workbook, t As TestPptTextFake, sel As TestPptSelFake, shp As Object, st As Long, ln As Long
    emo = ChrW$(&HD83D) & ChrW$(&HDE00)
    s = PptText_SafeValue(String$(254, "a") & emo & "b", cut)
    AssertTrue cut
    AssertEqual 254, Len(s), "o emoji que nao cabe inteiro sai"
    s = PptText_SafeValue(String$(253, "a") & emo & "b", cut)
    AssertEqual 255, Len(s), "o emoji que cabe fica inteiro"
    AssertEqual emo, Right$(s, 2)
    Set wb = SavedBook("ppttext-grupo.xlsx")
    Set t = TextShape("v=X"): t.Child = -1
    AssertEqual PPTTEXT_MSG_GROUP, PptText_Link(t, 3, 1, wb.Worksheets(1).Range("A1")), "filho de grupo"
    AssertEqual "v=X", t.Text
    AssertEqual 0, t.Tags.Writes
    Set sel = New TestPptSelFake: Set sel.ShapeRange = New Collection: sel.ShapeRange.Add TextShape("grupo")
    sel.HasChildShapeRange = True: Set sel.ChildShapeRange = New Collection: sel.ChildShapeRange.Add t
    sel.Start = 3: sel.Length = 1
    AssertTrue PptText_SelectionParts(3, sel, shp, st, ln)
    AssertTrue shp Is t, "selecao num grupo: o filho (PptText_Link o recusa)"
    CloseTemp wb
End Sub

' M7: celula do titulo apagada (#REF!): o titulo ligado sai no Alterar dados (nao religa a celula que subiu). M8: falha ao ligar o titulo
' (gancho) devolve o grafico como estava.
Public Sub Test_LinkTitle_RefDeletedAndFailureUndone()
    Dim wb As Workbook, ws As Worksheet, c As Chart, en As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5): ws.Range("A3:C3").Value = Array("T2", 20, 6)
    ws.Range("E5").Value = "Titulo": ws.Range("E6").Value = "Outro texto"
    ws.Range("A1:C3").Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    ' M8: falha depois de ligar HasTitle: volta sem titulo e sem opcao
    Util_FailThinSyncForTest = "linktitle-formula:1"
    On Error Resume Next
    LinkTitle_Apply c, ws.Range("E5")
    en = Err.Number
    On Error GoTo 0
    Util_FailThinSyncForTest = ""
    AssertEqual vbObjectError + 9994, en, "o erro sobe"
    AssertFalse c.HasTitle, "sem o titulo automatico"
    AssertEqual "", Registry_GetOption(c, "titleref", "")
    ' M7
    Core_TestSetTitleCell ws.Range("E5")
    Core_LinkTitle
    AssertEqual "Titulo", c.ChartTitle.Text
    ws.Rows(5).Delete
    AssertTrue LinkTitle_Cell(c) Is Nothing, "formula com #REF!: sem celula (" & c.ChartTitle.Formula & ")"
    AssertTrue Core_RebuildFromRange(c, ws.Range("A1:C3"))
    Set c = Core_ActiveChart()
    AssertFalse c.HasTitle, "o titulo ligado sai (nao vira 'Outro texto')"
    AssertEqual "", Registry_GetOption(c, "titleref", "")
    Watchers_Stop
    CloseTemp wb
End Sub
