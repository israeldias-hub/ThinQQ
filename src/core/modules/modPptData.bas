Attribute VB_Name = "modPptData"
Option Explicit
' Graficos criados PELO POWERPOINT (1.8.0; "Insert > Elements" do think-cell; spec 2026-10-15 secao 14). Pasta de dados ao lado da
' apresentacao ("<nome> - dados ThinQQ.xlsx", sem macros, marcada pelo nome oculto TQQ_DATABOOK), uma planilha por grafico ("Gráfico N":
' tabela em A1 e o grafico ThinQQ ao lado). O grafico vai ao slide VINCULADO (o motor da 1.4.0) e nasce automatico (1.6.0). Aberta para
' edicao sempre com as macros desligadas (AutomationSecurity = 3). Sem processo filho.
Public Const PPTDATA_SUFFIX As String = " - dados ThinQQ.xlsx"
Public Const PPTDATA_MARK As String = "TQQ_DATABOOK"
Public Const PPTDATA_TAG As String = "TQQ_DATA"
Public Const PPTDATA_SHEET_PREFIX As String = "Gráfico "
Public Const PPTDATA_MSG_UNSAVED As String = "Salve a apresentação primeiro: os dados do gráfico ficam numa pasta do Excel ao lado dela (""<nome> - dados ThinQQ.xlsx"")."
Public Const PPTDATA_MSG_WEB As String = "A apresentação está aberta por um endereço da Web (OneDrive/SharePoint). Salve-a numa pasta do computador (ou na pasta sincronizada do OneDrive) para inserir gráficos do ThinQQ."
Public Const PPTDATA_MSG_RO As String = "A pasta de dados está aberta por outra pessoa ou é somente leitura: "
Public Const PPTDATA_MSG_SAMENAME As String = "Feche no Excel a pasta '%1': ela tem o mesmo nome da pasta de dados desta apresentação."
Public Const PPTDATA_MSG_NOTMARKED As String = "Já existe o arquivo '%1' e ele não é uma pasta de dados do ThinQQ (nada foi alterado nele). Renomeie ou mova esse arquivo para inserir gráficos nesta apresentação."
Private mTestTable As Variant           ' gancho de teste: tabela usada no lugar da de exemplo (so no build Test; Empty = a de exemplo)
Private mTestFail As String             ' gancho de teste: ponto de Core_PptInsertChart que falha uma vez (so no build Test)

' Gancho de teste: a proxima PptData_NewChart usa esta tabela (uso unico). Inerte fora do build Test.
Public Sub PptData_SetTableForTest(tbl As Variant)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestTable = tbl
End Sub

' Gancho de teste: o proximo PptData_Hook com este ponto levanta erro (uso unico). Inerte fora do build Test.
Public Sub PptData_SetFailForTest(point As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestFail = point
End Sub

Public Sub PptData_ResetTestHooks()
    mTestTable = Empty
    mTestFail = ""
End Sub

Public Sub PptData_Hook(point As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    If Len(mTestFail) = 0 Or mTestFail <> point Then Exit Sub
    mTestFail = ""
    Err.Raise vbObjectError + 9993, "PptData", "Falha simulada (teste) " & point
End Sub

' Caminho da pasta de dados de uma apresentacao (FullName): mesma pasta, nome sem a ultima extensao + PPTDATA_SUFFIX. "" para endereco
' http(s) ou apresentacao sem pasta (nunca salva). Pura.
Public Function PptData_PathFor(presFullName As String) As String
    Dim p As Long, k As Long, base As String, f As String
    f = Trim$(presFullName)
    If LCase$(Left$(f, 7)) = "http://" Or LCase$(Left$(f, 8)) = "https://" Then Exit Function
    p = InStrRev(f, "\")
    If p < 2 Then Exit Function
    base = Mid$(f, p + 1)
    k = InStrRev(base, ".")
    If k > 1 Then base = Left$(base, k - 1)
    PptData_PathFor = Left$(f, p - 1) & "\" & base & PPTDATA_SUFFIX
End Function

' Primeiro "Gráfico N" livre; names = nomes das planilhas como "|a|b|". Pura.
Public Function PptData_NextSheetName(names As String) As String
    Dim i As Long
    For i = 1 To 999
        If InStr(1, names, "|" & PPTDATA_SHEET_PREFIX & i & "|", vbTextCompare) = 0 Then PptData_NextSheetName = PPTDATA_SHEET_PREFIX & i: Exit Function
    Next
End Function

Private Function PptData_SheetNames(wb As Workbook) As String
    Dim ws As Object, s As String
    s = "|"
    For Each ws In wb.Sheets
        s = s & ws.Name & "|"
    Next
    PptData_SheetNames = s
End Function

' A pasta tem a marca TQQ_DATABOOK. Nunca propaga erro.
Public Function PptData_IsDataBook(wb As Workbook) As Boolean
    Dim n As Name
    On Error Resume Next
    Set n = wb.Names(PPTDATA_MARK)
    PptData_IsDataBook = Not n Is Nothing
    Err.Clear
End Function

' Pasta ja aberta neste Excel com o caminho dado (Nothing se nao). Outra pasta aberta com o MESMO NOME de arquivo mas de outro lugar (o Excel
' nao abre duas pastas de mesmo nome: duas versoes do deck, pasta copiada, unidade mapeada x UNC) = erro 7130 com o texto para o usuario.
' Propaga erro.
Private Function PptData_FindOpen(path As String) As Workbook
    Dim w As Workbook, fileName As String
    fileName = Mid$(path, InStrRev(path, "\") + 1)
    For Each w In Application.Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then
            Set PptData_FindOpen = w
            Exit Function
        End If
        If StrComp(w.Name, fileName, vbTextCompare) = 0 Then Err.Raise vbObjectError + 7130, "PptData", Replace(PPTDATA_MSG_SAMENAME, "%1", w.FullName)
    Next
End Function

' Abre para EDICAO (macros desligadas e eventos desligados durante o Open, sem atualizar vinculos, sem MRU, sem a caixa de senha) ou devolve a
' ja aberta neste Excel. Somente leitura (aberta por outra pessoa, ou ja aberta somente leitura neste Excel) = erro 7130 (a que este codigo
' abriu e fechada antes); outra pasta de mesmo nome aberta = 7130 (PptData_FindOpen). Propaga erro.
Public Function PptData_OpenForEdit(path As String) As Workbook
    Dim w As Workbook, prevSec As Long, prevAlerts As Boolean, prevEv As Boolean, en As Long, ed As String
    Set w = PptData_FindOpen(path)
    If Not w Is Nothing Then
        If w.ReadOnly Then Err.Raise vbObjectError + 7130, "PptData", PPTDATA_MSG_RO & path
        Set PptData_OpenForEdit = w
        Exit Function
    End If
    prevAlerts = Application.DisplayAlerts: prevEv = Application.EnableEvents: prevSec = Application.AutomationSecurity
    Application.DisplayAlerts = False: Application.EnableEvents = False
    Application.AutomationSecurity = 3                          ' msoAutomationSecurityForceDisable
    On Error Resume Next
    Set w = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=False, Password:="x", IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    en = Err.Number: ed = Err.Description
    Err.Clear
    If en = 0 And Not w Is Nothing Then
        If w.ReadOnly Then w.Close SaveChanges:=False: en = vbObjectError + 7130: ed = PPTDATA_MSG_RO & path
    End If
    On Error GoTo 0
    Application.AutomationSecurity = prevSec
    Application.DisplayAlerts = prevAlerts: Application.EnableEvents = prevEv
    If en <> 0 Then Err.Raise en, "PptData_OpenForEdit", ed
    Set PptData_OpenForEdit = w
End Function

' A pasta de dados da apresentacao: aberta para edicao se existe; senao criada ("Leia-me" + marca) e salva como .xlsx. Propaga erro; a pasta
' nova que nao chegou a ser salva e fechada (nenhuma "Pasta1" sobra no Excel). Arquivo com o nome derivado mas SEM a marca TQQ_DATABOOK (um
' arquivo do usuario que coincide no nome) = erro 7130 sem alterar nada (fechado sem salvar se foi este codigo que o abriu).
Public Function PptData_OpenOrCreate(path As String, presName As String) As Workbook
    Dim w As Workbook, prevAlerts As Boolean, en As Long, ed As String, wasOpen As Boolean
    If CBool(CreateObject("Scripting.FileSystemObject").FileExists(path)) Then
        wasOpen = Not PptData_FindOpen(path) Is Nothing
        Set w = PptData_OpenForEdit(path)
        If Not PptData_IsDataBook(w) Then
            If Not wasOpen Then
                prevAlerts = Application.DisplayAlerts: Application.DisplayAlerts = False
                On Error Resume Next
                w.Close SaveChanges:=False
                On Error GoTo 0
                Application.DisplayAlerts = prevAlerts
            End If
            Err.Raise vbObjectError + 7130, "PptData", Replace(PPTDATA_MSG_NOTMARKED, "%1", path)
        End If
        Set PptData_OpenOrCreate = w
        Exit Function
    End If
    PptData_FindOpen path                                       ' outra pasta de mesmo nome aberta: o SaveAs falharia com 1004
    prevAlerts = Application.DisplayAlerts
    On Error GoTo Fail
    Set w = Application.Workbooks.Add(xlWBATWorksheet)
    w.Worksheets(1).Name = "Leia-me"
    w.Worksheets(1).Range("A1").Value = "Dados dos gráficos do ThinQQ da apresentação '" & presName & "'."
    w.Worksheets(1).Range("A2").Value = "Cada planilha 'Gráfico N' alimenta um gráfico do slide. Envie este arquivo junto com a apresentação (PowerPoint: ThinQQ > Empacotar para enviar)."
    w.Names.Add Name:=PPTDATA_MARK, RefersTo:="=""1""", Visible:=False
    Application.DisplayAlerts = False
    w.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook, AddToMru:=False
    Application.DisplayAlerts = prevAlerts
    Set PptData_OpenOrCreate = w
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Application.DisplayAlerts = prevAlerts
    If Not w Is Nothing Then w.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise en, "PptData_OpenOrCreate", ed
End Function

' Planilha nova "Gráfico N" com a tabela de exemplo do tipo e o grafico ThinQQ inserido pelo MESMO caminho do usuario (Core_Insert: formato
' inicial, avisos, observadores). rng volta com o intervalo dos dados. A planilha criada e apagada em qualquer falha (o grafico nao nasceu
' ou outro erro no meio). Propaga erro (7131 tipo sem exemplo; 7132 o Core_Insert nao criou o grafico).
Public Function PptData_NewChart(wb As Workbook, kind As String, ByRef rng As Range) As ChartObject
    Dim ws As Worksheet, tbl As Variant, c As Chart, prevAlerts As Boolean, en As Long, ed As String
    tbl = Samples_TableFor(kind)
    If BUILD_CONFIG = "Test" And Not IsEmpty(mTestTable) Then tbl = mTestTable: mTestTable = Empty
    If IsEmpty(tbl) Then Err.Raise vbObjectError + 7131, "PptData", "Tipo de gráfico desconhecido: " & kind
    On Error GoTo Fail
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = PptData_NextSheetName(PptData_SheetNames(wb))
    Set rng = Samples_WriteTable(ws.Range("A1"), tbl)
    ws.Columns(1).ColumnWidth = 16
    wb.Activate
    ws.Activate
    rng.Select
    Core_ClearLastNotice
    Core_Insert kind
    Set c = Core_ActiveChart()
    If Not c Is Nothing Then
        If Not c.Parent.Parent Is ws Then Set c = Nothing        ' grafico ativo de outro lugar: nao e o que acabou de nascer
    End If
    If c Is Nothing Then Err.Raise vbObjectError + 7132, "PptData", "Não foi possível criar o gráfico na pasta de dados: " & Core_LastNotice()
    Set PptData_NewChart = c.Parent
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set rng = Nothing
    If Not ws Is Nothing Then
        prevAlerts = Application.DisplayAlerts: Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = prevAlerts
    End If
    On Error GoTo 0
    Err.Raise en, "PptData", ed
End Function
