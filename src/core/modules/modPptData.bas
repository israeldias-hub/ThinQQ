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

' Abre para EDICAO (macros desligadas e eventos desligados durante o Open, sem atualizar vinculos, sem MRU, sem a caixa de senha) ou devolve a
' ja aberta neste Excel. Somente leitura (aberta por outra pessoa) = erro 7130 (fechada antes). Propaga erro.
Public Function PptData_OpenForEdit(path As String) As Workbook
    Dim w As Workbook, prevSec As Long, prevAlerts As Boolean, prevEv As Boolean, en As Long, ed As String
    For Each w In Application.Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then Set PptData_OpenForEdit = w: Exit Function
    Next
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
' nova que nao chegou a ser salva e fechada (nenhuma "Pasta1" sobra no Excel).
Public Function PptData_OpenOrCreate(path As String, presName As String) As Workbook
    Dim w As Workbook, prevAlerts As Boolean, en As Long, ed As String
    If CBool(CreateObject("Scripting.FileSystemObject").FileExists(path)) Then Set PptData_OpenOrCreate = PptData_OpenForEdit(path): Exit Function
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
