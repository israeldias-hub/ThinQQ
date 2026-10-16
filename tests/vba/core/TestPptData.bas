Attribute VB_Name = "TestPptData"
Option Explicit
' Parte G (1.8.0): pasta de dados dos graficos criados pelo PowerPoint (sem PowerPoint: TestPptDataFake faz o slide e a apresentacao nos
' caminhos que terminam antes da colagem).

Private Function SheetExists(wb As Workbook, nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = wb.Sheets(nm)
    SheetExists = Not s Is Nothing
    Err.Clear                                                       ' o Cleanup de quem chama le Err.Number ao cair nele
End Function

' Pasta aberta neste Excel com este caminho (Nothing se nao).
Private Function OpenBook(path As String) As Workbook
    Dim w As Workbook
    For Each w In Application.Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then Set OpenBook = w: Exit Function
    Next
End Function

' Fecha sem salvar (se aberta) e apaga o arquivo (se existe). Nunca propaga erro.
Private Sub DropBook(path As String)
    Dim w As Workbook
    On Error Resume Next
    Set w = OpenBook(path)
    If Not w Is Nothing Then w.Close SaveChanges:=False
    If Len(path) > 0 Then
        If Len(Dir(path)) > 0 Then Kill path
    End If
    Err.Clear
End Sub

' Slide falso cuja apresentacao "esta salva" em presPath.
Private Function FakeSlide(presPath As String) As TestPptDataFake
    Dim pres As New TestPptDataFake, sl As New TestPptDataFake
    pres.FullName = presPath
    pres.Path = Left$(presPath, InStrRev(presPath, "\") - 1)
    pres.Name = Mid$(presPath, InStrRev(presPath, "\") + 1)
    Set pres.Tags = CreateObject("Scripting.Dictionary")
    Set sl.Parent = pres
    Set FakeSlide = sl
End Function

' Pasta comum (sem a marca TQQ_DATABOOK) salva em path e fechada.
Private Sub PlainBook(path As String)
    Dim w As Workbook
    Set w = Application.Workbooks.Add(xlWBATWorksheet)
    w.Worksheets(1).Range("A1").Value = "do usuario"
    w.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
    w.Close SaveChanges:=False
End Sub

Public Sub Test_PptData_PathFor()
    AssertEqual "C:\a\Vendas - dados ThinQQ.xlsx", PptData_PathFor("C:\a\Vendas.pptx")
    AssertEqual "C:\a\b.c - dados ThinQQ.xlsx", PptData_PathFor("C:\a\b.c.pptx"), "so a ultima extensao sai"
    AssertEqual "\\srv\x\R - dados ThinQQ.xlsx", PptData_PathFor("\\srv\x\R.pptx")
    AssertEqual "", PptData_PathFor("https://empresa.sharepoint.com/x/R.pptx"), "endereco da Web: recusado"
    AssertEqual "", PptData_PathFor("Apresentação1"), "nunca salva: sem pasta"
End Sub

Public Sub Test_PptData_NextSheetName()
    AssertEqual "Gráfico 1", PptData_NextSheetName("|Leia-me|")
    AssertEqual "Gráfico 3", PptData_NextSheetName("|Leia-me|Gráfico 1|Gráfico 2|")
    AssertEqual "Gráfico 2", PptData_NextSheetName("|Gráfico 1|Gráfico 3|"), "primeiro livre"
End Sub

' Tipo desconhecido e recusado antes de qualquer acesso ao slide (nem PowerPoint e preciso); nunca propaga erro.
Public Sub Test_PptInsertChart_UnknownKind_NoPowerPoint()
    AssertEqual "Tipo de gráfico desconhecido: xyz", Core_PptInsertChart(Nothing, "xyz")
End Sub

Public Sub Test_PptData_CreateReopenAndNewChart()
    Dim p As String, wb As Workbook, co As ChartObject, rng As Range, en As Long, ed As String, wasActive As Boolean
    wasActive = Watchers_IsActive()
    p = TempDir() & "\Teste G - dados ThinQQ.xlsx"
    On Error GoTo Cleanup
    Set wb = PptData_OpenOrCreate(p, "Teste G.pptx")
    AssertTrue PptData_IsDataBook(wb), "marcada"
    AssertEqual "Leia-me", wb.Worksheets(1).Name
    AssertTrue Len(Dir(p)) > 0, "salva no caminho"
    Set co = PptData_NewChart(wb, "stacked", rng)
    AssertEqual "Gráfico 1", co.Parent.Name
    AssertEqual "$A$1:$D$5", rng.Address, "tabela de exemplo (Ano x 3 regioes, 4 anos)"
    AssertEqual "stacked", Registry_ChartType(co.Chart)
    Set co = PptData_NewChart(wb, "pie", rng)
    AssertEqual "Gráfico 2", co.Parent.Name
    wb.Save
    AssertTrue PptData_OpenOrCreate(p, "Teste G.pptx") Is wb, "ja aberta: a mesma"
    Watchers_Stop
    wb.Close SaveChanges:=False
    Set wb = PptData_OpenForEdit(p)
    AssertFalse wb.ReadOnly
    AssertTrue SheetExists(wb, "Gráfico 1") And SheetExists(wb, "Gráfico 2"), "reaberta com os dois graficos"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set co = Nothing: Set rng = Nothing: Set wb = Nothing
    Watchers_Stop
    DropBook p
    If wasActive Then Watchers_Start
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptData", ed
End Sub

' I1 (revisao da G2): um aviso do caminho comum (Core_Insert com tabela invalida) durante a chamada do PowerPoint NAO abre caixa no Excel
' (o PowerPoint esta parado no xl.Run): so registra, e o texto volta para quem chamou. O modo silencioso fecha em todo caminho.
Public Sub Test_PptInsertChart_NotifyInsideCall_NoMessageBox()
    Dim td As String, p As String, msg As String, wb As Workbook, en As Long, ed As String
    td = TempDir()
    p = td & "\Quieto - dados ThinQQ.xlsx"
    On Error GoTo Cleanup
    Core_ClearLastNotice
    PptData_SetTableForTest Array(Array("Ano", "Sul"))           ' so o cabecalho: o Core_Insert avisa (minimo 2 x 2)
    msg = Core_PptInsertChart(FakeSlide(td & "\Quieto.pptx"), "stacked")
    AssertTrue Core_NoticeCount() >= 1, "o Core_Insert avisou"
    AssertEqual 0, Core_ShownCountForTest(), "nenhuma caixa durante a chamada do PowerPoint"
    AssertEqual "Não foi possível criar o gráfico na pasta de dados: " & Core_LastNotice(), msg, "o aviso volta no texto"
    AssertFalse Core_IsQuiet(), "modo silencioso desligado depois da chamada"
    AssertFalse Ppt_InOperation(), "operacao do PowerPoint fechada"
    Set wb = OpenBook(p)
    AssertTrue Not wb Is Nothing, "pasta de dados criada e aberta"
    AssertFalse SheetExists(wb, "Gráfico 1"), "planilha do grafico que nao nasceu apagada"
    Core_ClearLastNotice
    Core_Notify "fora da chamada"
    AssertEqual 1, Core_ShownCountForTest(), "fora da chamada o aviso volta a abrir caixa"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptData_ResetTestHooks
    Core_ClearLastNotice
    Set wb = Nothing
    Watchers_Stop
    DropBook p
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptData", ed
End Sub

Public Sub Test_Core_Quiet_NestsAndBalances()
    Core_ClearLastNotice
    Core_QuietBegin
    Core_QuietBegin
    Core_QuietEnd
    AssertTrue Core_IsQuiet(), "aninhado: ainda silencioso"
    Core_Notify "x"
    AssertFalse Core_Confirm("y?"), "silencioso: pergunta responde Nao"
    AssertEqual 0, Core_ShownCountForTest()
    Core_QuietEnd
    Core_QuietEnd                                                   ' End a mais nao fica negativo
    AssertFalse Core_IsQuiet()
    Core_ClearLastNotice
End Sub

' M7: arquivo com o nome derivado que NAO e pasta de dados (sem a marca): recusado, nada alterado, nada fica aberto.
Public Sub Test_PptInsertChart_ExistingUnmarkedFile_Refused()
    Dim td As String, p As String, msg As String, t0 As Date, en As Long, ed As String
    td = TempDir()
    p = td & "\Alheio - dados ThinQQ.xlsx"
    On Error GoTo Cleanup
    PlainBook p
    t0 = FileDateTime(p)
    msg = Core_PptInsertChart(FakeSlide(td & "\Alheio.pptx"), "stacked")
    AssertEqual Replace(PPTDATA_MSG_NOTMARKED, "%1", p), msg
    AssertTrue OpenBook(p) Is Nothing, "o arquivo do usuario nao ficou aberto"
    AssertTrue FileDateTime(p) = t0, "o arquivo do usuario nao foi gravado"
    AssertFalse Core_IsQuiet()
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    DropBook p
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptData", ed
End Sub

' I2: outra pasta com o MESMO NOME aberta de outro lugar: mensagem propria (antes era "... (1004)"), na criacao e na abertura.
Public Sub Test_PptData_SameNameOpenElsewhere_Refused()
    Dim tdA As String, tdB As String, pA As String, pB As String, msg As String, w As Workbook, en As Long, ed As String, got As String
    tdA = TempDir(): tdB = TempDir()
    pA = tdA & "\Dup - dados ThinQQ.xlsx": pB = tdB & "\Dup - dados ThinQQ.xlsx"
    On Error GoTo Cleanup
    PlainBook pA
    Set w = Application.Workbooks.Open(pA)
    msg = Core_PptInsertChart(FakeSlide(tdB & "\Dup.pptx"), "stacked")
    AssertEqual Replace(PPTDATA_MSG_SAMENAME, "%1", pA), msg, "criacao"
    AssertEqual "", Dir(pB), "nada criado em B"
    w.Close SaveChanges:=False: Set w = Nothing
    PptData_OpenOrCreate(pB, "Dup.pptx").Close SaveChanges:=False
    Set w = Application.Workbooks.Open(pA)
    On Error Resume Next
    PptData_OpenForEdit pB
    got = Err.Description: en = Err.Number
    On Error GoTo Cleanup
    AssertEqual vbObjectError + 7130, en, "abertura: 7130"
    AssertEqual Replace(PPTDATA_MSG_SAMENAME, "%1", pA), got
    en = 0
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set w = Nothing
    DropBook pA
    DropBook pB
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptData", ed
End Sub

' G3: Salvar no PowerPoint (Core_PptDataSave) salva so a pasta de DADOS da apresentacao (etiqueta TQQ_DATABOOK) aberta e alterada; nunca
' abre nada, nunca salva a pasta do usuario nem a somente leitura; sem caixa e com o modo silencioso equilibrado.
Public Sub Test_PptDataSave_OnlyOpenChangedDataBook()
    Dim td As String, p As String, pu As String, sl As TestPptDataFake, wb As Workbook, wu As Workbook, en As Long, ed As String
    td = TempDir()
    p = td & "\Salva - dados ThinQQ.xlsx": pu = td & "\Usuario.xlsx"
    On Error GoTo Cleanup
    Set sl = FakeSlide(td & "\Salva.pptx")
    AssertEqual 0, Core_PptDataSave(sl.Parent), "sem etiqueta: nada"
    AssertEqual 0, Core_PptDataSave(Nothing), "nunca propaga erro"
    sl.Parent.Tags(PPTDATA_MARK) = p                                ' (a leitura acima ja criou a chave vazia no Dictionary)
    AssertEqual 0, Core_PptDataSave(sl.Parent), "pasta de dados fechada: nada (nunca abre)"
    AssertTrue OpenBook(p) Is Nothing
    Set wb = PptData_OpenOrCreate(p, "Salva.pptx")
    AssertEqual 0, Core_PptDataSave(sl.Parent), "aberta sem alteracao: nada"
    wb.Worksheets(1).Range("C1").Value = "x"
    Core_ClearLastNotice
    AssertEqual 1, Core_PptDataSave(sl.Parent)
    AssertTrue wb.Saved
    AssertFalse Core_IsQuiet()
    AssertEqual 0, Core_ShownCountForTest()
    PlainBook pu
    Set wu = Application.Workbooks.Open(pu)
    wu.Worksheets(1).Range("B1").Value = "y"
    sl.Parent.Tags(PPTDATA_MARK) = pu
    AssertEqual 0, Core_PptDataSave(sl.Parent), "pasta do usuario (sem a marca) nunca e salva"
    AssertFalse wu.Saved
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set wb = Nothing: Set wu = Nothing
    DropBook p
    DropBook pu
    Core_ClearLastNotice
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptData", ed
End Sub

' G7 (1.9.0): estado das opcoes que a faixa do PowerPoint le na etiqueta TQQ_OPTS, sem chamar o Excel. Um valor alterado no registro do
' grafico nunca leva separadores nem Chr(30)/Chr(31)/quebras de linha para a etiqueta.
Public Sub Test_PptSync_OptsText()
    Dim wb As Workbook, ws As Worksheet, c As Chart, s As String, en As Long, ed As String
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Samples_WriteTable(ws.Range("A1"), Samples_TableFor("stacked")).Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    s = PptSync_OptsText(c)
    AssertTrue InStr(s, "numfmt=") > 0 And InStr(s, "legend=") > 0, s
    AssertTrue InStr(s, "caps=") > 0 And InStr(s, "totals,") > 0 And InStr(s, "cagr,") > 0, "empilhadas: totais e CAGR (" & s & ")"
    Core_SetToggle "totals", True
    AssertTrue InStr(PptSync_OptsText(c), "totals=1") > 0
    Core_SetToggle "totals", False
    AssertTrue InStr(PptSync_OptsText(c), "totals=0") > 0
    Registry_SetOption c, "numfmt", "pct0,x" & Chr$(31) & vbLf & "1" & vbTab & "y"      ' (o registro ja recusa ';' e '=')
    s = PptSync_OptsText(c)
    AssertTrue InStr(s, "numfmt=pct0x1y;legend=") > 0, "so letras e digitos: " & s
    AssertEqual 0, InStr(s, Chr$(31)) + InStr(s, Chr$(30)) + InStr(s, vbLf) + InStr(s, vbCr) + InStr(s, vbTab)
    AssertEqual "abc123", PptSync_OptValue("a,b;c=1" & Chr$(30) & "2 3")
    AssertEqual 20, Len(PptSync_OptValue(String$(50, "a"))), "no maximo 20 caracteres"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set c = Nothing: Set ws = Nothing
    Watchers_Stop
    CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptData", ed
End Sub

' G7: a API 4 nunca propaga erro e recusa um comando desconhecido antes de procurar as formas.
Public Sub Test_PptCommand_NeverRaises_NoPowerPoint()
    AssertEqual PPTCMD_MSG_UNKNOWN & "xyz", Core_PptCommand(Nothing, "", "xyz", "")
    AssertTrue Len(Core_PptCommand(Nothing, "256|a|", "clear", "")) > 0, "apresentacao Nothing: mensagem, sem erro"
    AssertEqual "", Core_PptCategories(Nothing, "256|a|")
    AssertFalse Core_IsQuiet(), "modo silencioso fechado"
    AssertFalse Ppt_InOperation(), "operacao do PowerPoint fechada"
End Sub

' M1: pasta de dados ja aberta SOMENTE LEITURA neste Excel: PPTDATA_MSG_RO (antes falhava no Save com "... (1004)").
Public Sub Test_PptInsertChart_DataBookOpenReadOnly_Refused()
    Dim td As String, p As String, msg As String, w As Workbook, en As Long, ed As String
    td = TempDir()
    p = td & "\SoLeitura - dados ThinQQ.xlsx"
    On Error GoTo Cleanup
    PptData_OpenOrCreate(p, "SoLeitura.pptx").Close SaveChanges:=False
    Set w = Application.Workbooks.Open(Filename:=p, ReadOnly:=True)
    msg = Core_PptInsertChart(FakeSlide(td & "\SoLeitura.pptx"), "stacked")
    AssertEqual PPTDATA_MSG_RO & p, msg
    AssertFalse SheetExists(w, "Gráfico 1"), "nada criado"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set w = Nothing
    DropBook p
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPptData", ed
End Sub
