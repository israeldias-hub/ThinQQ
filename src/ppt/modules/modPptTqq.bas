Attribute VB_Name = "modPptTqq"
Option Explicit
' Suplemento do PowerPoint do ThinQQ (ThinQQ.ppam; spec 2026-10-14 secao 4.3). Ao abrir uma apresentacao com graficos ThinQQ vinculados
' (etiqueta TQQ_CHART, gravada pelo core ao colar) pergunta (padrao), atualiza sempre ou nao faz nada (opcao "Ao abrir") e pede ao core do
' ThinQQ no Excel para atualizar dados e anotacoes (Core_PptSyncPresentation). Excel: o aberto, se ja tem o core; senao uma instancia NOVA e
' invisivel (o Excel e de varias instancias: nunca a do usuario), com o core mais novo de %LOCALAPPDATA%\ThinQQ\versions, fechada no fim.
' Sem o core: so os dados (vinculo nativo), nunca com a pasta inexistente (o Update travaria o PowerPoint). Ligacao tardia; sem Declare.
Public Const PPT_APP_NAME As String = "ThinQQ"
Public Const PPT_SECTION As String = "PowerPoint"
Public Const PPT_KEY_OPENMODE As String = "OpenMode"
Public Const PPT_TAG_CHART As String = "TQQ_CHART"
Public Const PPT_TAG_BOOK As String = "TQQ_BOOK"
Public Const PPT_MSG_NONE As String = "Esta apresentação não tem gráficos do ThinQQ vinculados."
Public Const PPT_MSG_NO_PRES As String = "Abra uma apresentação primeiro."
Public Const PPT_MSG_NO_EXCEL As String = "O Excel não está instalado (ou não pôde ser iniciado): os gráficos do ThinQQ não foram atualizados."
Public Const PPT_MSG_DATA_ONLY As String = "O ThinQQ não está instalado no Excel deste computador: só os dados foram atualizados; setas e outras anotações do ThinQQ ficaram como estavam."
Public Const PPT_MSG_OLD_CORE As String = "Atualize o ThinQQ no Excel para atualizar também as setas."
Public Const PPT_MSG_BUSY As String = "O Excel está ocupado (uma célula em edição ou uma caixa de diálogo aberta). Termine o que está fazendo no Excel e use ThinQQ > Atualizar gráficos."
Public Const PPT_MAX_SOURCES_SHOWN As Long = 5                ' pastas de origem listadas na pergunta ao abrir (depois, "...")
Public Const PPT_REQUIRED_API As Long = 1                   ' Core_PptApiVersion minima (spec 4.9); core sem ela ou menor = so os dados
Private Const RPC_E_CALL_REJECTED As Long = -2147418111
Private Const RPC_E_RETRYLATER As Long = -2147417846
Private mEvents As clsPptEvents
Private mSeen As Collection
Private mTestXl As Object
Private mTestCoreDir As String
Private mTestCreated As Boolean                 ' gancho de teste: o Excel injetado conta como criado por nos (Quit no fim)
Private mTestCoreInfo As String                 ' gancho de teste: "caminho|ro" do core aberto pelo suplemento
Private mTestOnOpenCalls As Long                ' gancho de teste: chamadas de PptTqq_OnOpen (eventos e Auto_Open)
Private mLastMsg As String
Private mLastLog As String
Private mBusy As Boolean                        ' reentrada: pergunta/sincronizacao em andamento (eventos chegam durante o MsgBox)

' Carga do suplemento: liga os eventos do PowerPoint e trata as apresentacoes ja abertas (o projeto de teste .pptm nao liga).
Public Sub Auto_Open()
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then Exit Sub
    Set mEvents = New clsPptEvents
    Set mEvents.App = Application
    PptTqq_ScanOpen
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Set mEvents = Nothing
End Sub

' Apresentacoes ja abertas quando o suplemento carrega: as mesmas regras da abertura (inclusive a guarda de janela).
Public Sub PptTqq_ScanOpen()
    Dim i As Long, n As Long
    On Error Resume Next
    n = Application.Presentations.Count
    For i = 1 To n
        PptTqq_OnOpen Application.Presentations(i)
    Next
End Sub

' Gancho de teste: Excel injetado, pasta das versoes do core e se o Excel injetado conta como criado pelo suplemento (asCreated: abre o
' core da pasta e da Quit no fim). Inerte fora do build Test.
Public Sub PptTqq_SetExcelForTest(xl As Object, coreDir As String, Optional asCreated As Boolean = False)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    Set mTestXl = xl
    mTestCoreDir = coreDir
    mTestCreated = asCreated
    mTestCoreInfo = ""
End Sub

' Gancho de teste: zera todos os ganchos e o estado da sessao (vistas, reentrada, contadores). Inerte fora do build Test.
Public Sub PptTqq_ResetTestHooks()
    If BUILD_CONFIG <> "Test" Then Exit Sub
    Set mTestXl = Nothing
    mTestCoreDir = ""
    mTestCreated = False
    mTestCoreInfo = ""
    mTestOnOpenCalls = 0
    mBusy = False
    Set mSeen = Nothing
    mLastLog = ""
End Sub

Public Function PptTqq_CoreInfoForTest() As String
    PptTqq_CoreInfoForTest = mTestCoreInfo
End Function

Public Function PptTqq_OnOpenCallsForTest() As Long
    PptTqq_OnOpenCallsForTest = mTestOnOpenCalls
End Function

' Log silencioso (caminho dos eventos): %LOCALAPPDATA%\ThinQQ\logs\thinqq.log se a pasta existir (nunca cria); no build Test so memoria.
Public Sub PptTqq_Log(msg As String)
    Dim f As Integer, d As String
    mLastLog = msg
    If BUILD_CONFIG = "Test" Then Exit Sub
    On Error Resume Next
    d = Environ$("LOCALAPPDATA") & "\ThinQQ\logs"
    If Len(Dir$(d, vbDirectory)) = 0 Then Exit Sub
    f = FreeFile
    Open d & "\thinqq.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & "ppt" & vbTab & msg
    Close #f
End Sub

Public Function PptTqq_LastLog() As String
    PptTqq_LastLog = mLastLog
End Function

Public Function PptTqq_LastMsg() As String
    PptTqq_LastMsg = mLastMsg
End Function

Public Sub PptTqq_ClearLastMsg()
    mLastMsg = ""
End Sub

' Unico ponto de MsgBox de aviso (log-only no build Test).
Public Sub PptTqq_Msg(msg As String, Optional style As VbMsgBoxStyle = vbInformation)
    mLastMsg = msg
    If BUILD_CONFIG = "Test" Then Exit Sub
    MsgBox msg, style, PPT_APP_NAME
End Sub

' Unico ponto de pergunta (no build Test registra e responde Nao).
Private Function PptTqq_Ask(msg As String) As VbMsgBoxResult
    mLastMsg = msg
    If BUILD_CONFIG = "Test" Then PptTqq_Ask = vbNo: Exit Function
    PptTqq_Ask = MsgBox(msg, vbQuestion + vbYesNo, PPT_APP_NAME)
End Function

Public Function PptTqq_NormalizeMode(m As String) As String
    Select Case LCase$(Trim$(m))
        Case "always": PptTqq_NormalizeMode = "always"
        Case "never": PptTqq_NormalizeMode = "never"
        Case Else: PptTqq_NormalizeMode = "ask"
    End Select
End Function

' Modo "Ao abrir" (HKCU\Software\VB and VBA Program Settings\ThinQQ\PowerPoint\OpenMode). Testes: sempre "ask", sem ler o registro.
Public Function PptTqq_OpenMode() As String
    If BUILD_CONFIG = "Test" Then PptTqq_OpenMode = "ask": Exit Function
    PptTqq_OpenMode = PptTqq_NormalizeMode(GetSetting(PPT_APP_NAME, PPT_SECTION, PPT_KEY_OPENMODE, "ask"))
End Function

Public Sub PptTqq_SetOpenMode(m As String)
    If BUILD_CONFIG = "Test" Then Exit Sub
    SaveSetting PPT_APP_NAME, PPT_SECTION, PPT_KEY_OPENMODE, PptTqq_NormalizeMode(m)
End Sub

' O que fazer ao abrir: "" (nada), "ask" ou "sync". hasRemoteSource: alguma pasta de origem e de rede (\\servidor\...) ou de endereco
' http(s) e ainda nao esta aberta no Excel: "Atualizar sempre" volta a perguntar (a apresentacao escolhe quais pastas a atualizacao abre;
' uma origem remota nao e aberta sem o usuario ver o caminho). Pura.
Public Function PptTqq_OnOpenAction(mode As String, isReadOnly As Boolean, nTagged As Long, alreadySeen As Boolean, hasRemoteSource As Boolean) As String
    If nTagged = 0 Or alreadySeen Or isReadOnly Or mode = "never" Then Exit Function
    If mode = "always" And Not hasRemoteSource Then PptTqq_OnOpenAction = "sync" Else PptTqq_OnOpenAction = "ask"
End Function

' Pergunta ao abrir. sources (opcional): as pastas de origem distintas (caminho completo), listadas antes da dica: no maximo
' PPT_MAX_SOURCES_SHOWN e "..." se houver mais.
Public Function PptTqq_AskText(n As Long, Optional sources As Collection) As String
    PptTqq_AskText = "Esta apresentação tem " & n & " gráfico(s) do ThinQQ vinculado(s) ao Excel. Atualizar agora com os dados atuais?" & PptTqq_SourcesText(sources) & vbLf & vbLf & "Para não perguntar mais: aba ThinQQ > Ao abrir."
End Function

' Bloco "Pastas de origem:" da pergunta (vazio sem pastas). Pura.
Private Function PptTqq_SourcesText(sources As Collection) As String
    Dim i As Long, s As String
    If sources Is Nothing Then Exit Function
    If sources.Count = 0 Then Exit Function
    s = vbLf & vbLf & "Pastas de origem:"
    For i = 1 To sources.Count
        If i > PPT_MAX_SOURCES_SHOWN Then s = s & vbLf & "...": Exit For
        s = s & vbLf & CStr(sources(i))
    Next
    PptTqq_SourcesText = s
End Function

' Caminho de rede (\\servidor\...) ou endereco http(s). Pura.
Public Function PptTqq_IsRemotePath(p As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(p))
    PptTqq_IsRemotePath = (Left$(t, 2) = "\\" Or Left$(t, 7) = "http://" Or Left$(t, 8) = "https://")
End Function

' Alguma pasta de origem remota que nao esta entre as abertas no Excel (names: FullName das pastas abertas). Pura.
Public Function PptTqq_HasRemoteSource(sources As Collection, names As Collection) As Boolean
    Dim v As Variant
    If sources Is Nothing Then Exit Function
    For Each v In sources
        If PptTqq_IsRemotePath(CStr(v)) Then
            If Not PptTqq_InNames(CStr(v), names) Then PptTqq_HasRemoteSource = True: Exit Function
        End If
    Next
End Function

Private Function PptTqq_InNames(p As String, names As Collection) As Boolean
    Dim n As Variant
    If names Is Nothing Then Exit Function
    For Each n In names
        If StrComp(CStr(n), Trim$(p), vbTextCompare) = 0 Then PptTqq_InNames = True: Exit Function
    Next
End Function

' Mesma regra do core (PptSync_ReachableAmong): aberta no Excel (FullName igual; vale para o endereco https) ou, se nao for http(s),
' arquivo existente no disco. Nunca propaga erro.
Public Function PptTqq_ReachableAmong(path As String, names As Collection) As Boolean
    Dim p As String
    On Error GoTo Done
    p = Trim$(path)
    If Len(p) = 0 Then Exit Function
    If PptTqq_InNames(p, names) Then PptTqq_ReachableAmong = True: Exit Function
    If LCase$(Left$(p, 7)) = "http://" Or LCase$(Left$(p, 8)) = "https://" Then Exit Function
    PptTqq_ReachableAmong = PptTqq_FileExists(p)
Done:
End Function

' FullName das pastas abertas num Excel (vazio se xl e Nothing ou nao responde). Nunca propaga erro.
Private Function PptTqq_WorkbookNames(xl As Object) As Collection
    Dim out As New Collection, wb As Object
    Set PptTqq_WorkbookNames = out
    If xl Is Nothing Then Exit Function
    On Error Resume Next
    For Each wb In xl.Workbooks
        out.Add CStr(wb.FullName)
    Next
    Err.Clear
End Function

' Pastas abertas no Excel em execucao, sem nunca criar um (build Test: so o Excel injetado). Ocupado ou ausente: vazio.
Private Function PptTqq_ExcelOpenNames() As Collection
    Dim xl As Object
    If BUILD_CONFIG = "Test" Then Set PptTqq_ExcelOpenNames = PptTqq_WorkbookNames(mTestXl): Exit Function
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    Err.Clear
    On Error GoTo 0
    Set PptTqq_ExcelOpenNames = PptTqq_WorkbookNames(xl)
End Function

Private Function PptTqq_Tag(shp As Object, key As String) As String
    On Error Resume Next
    PptTqq_Tag = CStr(shp.Tags.Item(key))
    Err.Clear
End Function

' Formas soltas marcadas pelo ThinQQ na apresentacao. sources (opcional): recebe as pastas de origem distintas (sem diferenciar maiusculas),
' na ordem em que aparecem: PptTqq_SourcePath(LinkFormat.SourceFullName, TQQ_BOOK); so le o vinculo e as etiquetas (nunca o arquivo).
Public Function PptTqq_CountTagged(pres As Object, Optional sources As Collection) As Long
    Dim sl As Object, shp As Object, p As String
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If Len(PptTqq_Tag(shp, PPT_TAG_CHART)) > 0 Then
                PptTqq_CountTagged = PptTqq_CountTagged + 1
                If Not sources Is Nothing Then
                    p = PptTqq_SourcePath(PptTqq_LinkSource(shp), PptTqq_Tag(shp, PPT_TAG_BOOK))
                    If Len(p) > 0 Then
                        On Error Resume Next
                        sources.Add p, LCase$(p)                ' repetida: chave ja existe, ignorada
                        Err.Clear
                        On Error GoTo 0
                    End If
                End If
            End If
        Next
    Next
End Function

' LinkFormat.SourceFullName da forma ("" se nao ha vinculo legivel). Nunca propaga erro.
Private Function PptTqq_LinkSource(shp As Object) As String
    On Error Resume Next
    PptTqq_LinkSource = CStr(shp.LinkFormat.SourceFullName)
    Err.Clear
End Function

' Mesma regra do core (PptSync_SourcePath). Pura.
Public Function PptTqq_SourcePath(linkFull As String, tagBook As String) As String
    Dim s As String, p As Long
    s = Trim$(linkFull)
    p = InStr(s, "!")
    If p > 0 Then s = Left$(s, p - 1)
    If Len(s) = 0 Then s = Trim$(tagBook)
    PptTqq_SourcePath = s
End Function

Private Function PptTqq_IsVersion(s As String) As Boolean
    Dim p() As String, i As Long
    p = Split(s, ".")
    If UBound(p) <> 2 Then Exit Function
    For i = 0 To 2
        If Len(p(i)) = 0 Or Not IsNumeric(p(i)) Or InStr(p(i), ",") > 0 Or InStr(p(i), "-") > 0 Then Exit Function
    Next
    PptTqq_IsVersion = True
End Function

' Compara x.y.z numericamente (1 se a > b, -1 se a < b, 0 iguais).
Private Function PptTqq_CompareVersion(a As String, b As String) As Long
    Dim pa() As String, pb() As String, i As Long
    pa = Split(a, "."): pb = Split(b, ".")
    For i = 0 To 2
        If CLng(pa(i)) > CLng(pb(i)) Then PptTqq_CompareVersion = 1: Exit Function
        If CLng(pa(i)) < CLng(pb(i)) Then PptTqq_CompareVersion = -1: Exit Function
    Next
End Function

' Maior versao x.y.z numa lista "a|b|c" (nomes de pastas); "" sem nenhuma valida. Pura.
Public Function PptTqq_NewestVersion(names As String) As String
    Dim p() As String, i As Long, best As String
    p = Split(names, "|")
    For i = LBound(p) To UBound(p)
        If PptTqq_IsVersion(p(i)) Then
            If Len(best) = 0 Then
                best = p(i)
            ElseIf PptTqq_CompareVersion(p(i), best) > 0 Then
                best = p(i)
            End If
        End If
    Next
    PptTqq_NewestVersion = best
End Function

Private Function PptTqq_CoreFileName() As String
    PptTqq_CoreFileName = "ThinQQ.core" & BUILD_SUFFIX & ".xlam"
End Function

Private Function PptTqq_CoreDir() As String
    If BUILD_CONFIG = "Test" Then PptTqq_CoreDir = mTestCoreDir: Exit Function
    PptTqq_CoreDir = Environ$("LOCALAPPDATA") & "\ThinQQ\versions"
End Function

' Pura: "nome:1|nome:0|..." (1 = a pasta tem o arquivo do core) -> versoes x.y.z COM o arquivo, da mais nova para a mais velha ("a|b").
' Mesma regra do loader (Loader_ListCoreVersions): pasta parcial, sem o arquivo, fica de fora.
Public Function PptTqq_CoreCandidates(listing As String) As String
    Dim p() As String, a() As String, i As Long, j As Long, n As Long, k As Long, nm As String, t As String, s As String
    p = Split(listing, "|")
    ReDim a(0 To UBound(p) + 1)
    For i = LBound(p) To UBound(p)
        k = InStrRev(p(i), ":")
        If k > 1 Then
            nm = Left$(p(i), k - 1)
            If Mid$(p(i), k + 1) = "1" And PptTqq_IsVersion(nm) Then
                a(n) = nm
                n = n + 1
            End If
        End If
    Next
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If PptTqq_CompareVersion(a(j), a(i)) > 0 Then t = a(i): a(i) = a(j): a(j) = t
        Next
    Next
    For i = 0 To n - 1
        If i > 0 Then s = s & "|"
        s = s & a(i)
    Next
    PptTqq_CoreCandidates = s
End Function

' Versoes candidatas na pasta d (subpastas com o arquivo do core), da mais nova para a mais velha.
Public Function PptTqq_ListCoreCandidates(d As String) As String
    Dim fso As Object, sf As Object, s As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function
    If Not fso.FolderExists(d) Then Exit Function
    For Each sf In fso.GetFolder(d).SubFolders
        s = s & sf.Name & ":" & IIf(fso.FileExists(sf.Path & "\" & PptTqq_CoreFileName()), "1", "0") & "|"
    Next
    Err.Clear
    PptTqq_ListCoreCandidates = PptTqq_CoreCandidates(s)
End Function

' Pura: resultado do Core_Ping. 1 = core respondeu; -1 = Excel ocupado (RPC); 0 = sem core.
Public Function PptTqq_PingClass(en As Long, v As String) As Long
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Then PptTqq_PingClass = -1: Exit Function
    If en = 0 And v = "pong" Then PptTqq_PingClass = 1
End Function

Private Function PptTqq_PingState(xl As Object) As Long
    Dim v As Variant, en As Long, s As String
    On Error Resume Next
    v = xl.Run("'" & PptTqq_CoreFileName() & "'!Core_Ping")
    en = Err.Number
    Err.Clear
    s = CStr(v)
    Err.Clear
    PptTqq_PingState = PptTqq_PingClass(en, s)
End Function

Private Sub PptTqq_RaiseBusy()
    Err.Raise RPC_E_CALL_REJECTED, "PptTqq", "Excel ocupado"
End Sub

' Excel a usar: no build Test, so o injetado (created conforme o gancho). Senao: o aberto (GetObject) SE ele tem o core (ocupado: erro de
' ocupado, sem cair para outra instancia); senao uma instancia nova e invisivel (created). O retorno e atribuido antes de configurar: a
' instancia criada nunca fica sem dono.
Private Function PptTqq_GetExcel(ByRef created As Boolean) As Object
    Dim xl As Object, st As Long
    created = False
    If BUILD_CONFIG = "Test" Then
        created = mTestCreated
        Set PptTqq_GetExcel = mTestXl
        Exit Function
    End If
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    Err.Clear
    On Error GoTo 0
    If Not xl Is Nothing Then
        st = PptTqq_PingState(xl)
        If st = 1 Then Set PptTqq_GetExcel = xl: Exit Function
        If st = -1 Then PptTqq_RaiseBusy
    End If
    Set xl = Nothing
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    Err.Clear
    If xl Is Nothing Then Exit Function
    created = True
    Set PptTqq_GetExcel = xl
    xl.Visible = False
    xl.DisplayAlerts = False
    Err.Clear
End Function

' Abre um core (somente leitura) e confere o Core_Ping; se nao responder, fecha o que abriu.
Private Function PptTqq_TryOpenCore(xl As Object, p As String) As Boolean
    Dim wb As Object
    On Error Resume Next
    Set wb = xl.Workbooks.Open(p, 0, True)
    Err.Clear
    If PptTqq_PingState(xl) = 1 Then
        PptTqq_TryOpenCore = True
        If BUILD_CONFIG = "Test" Then mTestCoreInfo = CStr(xl.Workbooks(PptTqq_CoreFileName()).FullName) & "|" & IIf(xl.Workbooks(PptTqq_CoreFileName()).ReadOnly, "ro", "rw")
        Err.Clear
        Exit Function
    End If
    If wb Is Nothing Then Set wb = xl.Workbooks(PptTqq_CoreFileName())
    If Not wb Is Nothing Then wb.Close False
    Err.Clear
End Function

' Nome do core respondendo no Excel ("" sem core). Na instancia criada por nos, abre o core instalado: da versao mais nova para a mais velha,
' a primeira que responder (como o loader: versao nova que nao carrega cai para a anterior).
Private Function PptTqq_EnsureCore(xl As Object, created As Boolean) As String
    Dim st As Long, d As String, cands As String, c() As String, i As Long
    st = PptTqq_PingState(xl)
    If st = -1 Then PptTqq_RaiseBusy
    If st = 1 Then PptTqq_EnsureCore = PptTqq_CoreFileName(): Exit Function
    If Not created Then Exit Function
    d = PptTqq_CoreDir()
    If Len(d) = 0 Then Exit Function
    cands = PptTqq_ListCoreCandidates(d)
    If Len(cands) = 0 Then Exit Function
    c = Split(cands, "|")
    For i = 0 To UBound(c)
        If PptTqq_TryOpenCore(xl, d & "\" & c(i) & "\" & PptTqq_CoreFileName()) Then
            PptTqq_EnsureCore = PptTqq_CoreFileName()
            Exit Function
        End If
        PptTqq_Log "PptTqq_EnsureCore: core " & c(i) & " nao carregou; tentando a versao anterior"
    Next
End Function

' Quit so na instancia que o suplemento criou.
Private Sub PptTqq_Release(xl As Object, created As Boolean)
    On Error Resume Next
    If created And Not xl Is Nothing Then
        xl.DisplayAlerts = False
        xl.Quit
        If BUILD_CONFIG = "Test" Then Set mTestXl = Nothing
    End If
    Set xl = Nothing
End Sub

Public Function PptTqq_ErrorText(en As Long, ed As String) As String
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Then
        PptTqq_ErrorText = PPT_MSG_BUSY
    ElseIf en = 429 Then
        PptTqq_ErrorText = PPT_MSG_NO_EXCEL
    Else
        PptTqq_ErrorText = "Não foi possível atualizar os gráficos do ThinQQ: " & ed & " (" & en & ")"
    End If
End Function

' Modo degradado (sem o core): so os dados pelo vinculo nativo, e so dos graficos cuja pasta e alcancavel pela regra do core
' (PptTqq_ReachableAmong: endereco https so se estiver aberto no Excel; names = pastas abertas no Excel usado).
Public Function PptTqq_DataOnly(pres As Object, Optional names As Collection) As String
    Dim sl As Object, shp As Object, n As Long, nMiss As Long, p As String, lf As String
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If Len(PptTqq_Tag(shp, PPT_TAG_CHART)) > 0 Then
                lf = ""
                On Error Resume Next
                lf = CStr(shp.LinkFormat.SourceFullName)
                Err.Clear
                On Error GoTo 0
                p = PptTqq_SourcePath(lf, PptTqq_Tag(shp, PPT_TAG_BOOK))
                If PptTqq_ReachableAmong(p, names) Then
                    On Error Resume Next
                    shp.LinkFormat.Update
                    If Err.Number = 0 Then n = n + 1
                    Err.Clear
                    On Error GoTo 0
                Else
                    nMiss = nMiss + 1
                End If
            End If
        Next
    Next
    PptTqq_DataOnly = "Dados atualizados em " & n & " gráfico(s)" & IIf(nMiss > 0, "; " & nMiss & " sem a pasta de origem (ficaram como estavam)", "") & "."
End Function

Private Function PptTqq_FileExists(p As String) As Boolean
    On Error Resume Next
    If Len(p) = 0 Then Exit Function
    PptTqq_FileExists = CreateObject("Scripting.FileSystemObject").FileExists(p)
    Err.Clear
End Function

' Versao da API do core (0 = core sem Core_PptApiVersion, anterior a 1.4.0). Excel ocupado: erro de ocupado.
Private Function PptTqq_CoreApi(xl As Object, core As String) As Long
    Dim v As Variant, en As Long
    On Error Resume Next
    v = xl.Run("'" & core & "'!Core_PptApiVersion")
    en = Err.Number
    Err.Clear
    On Error GoTo 0
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Then PptTqq_RaiseBusy
    If en = 0 And IsNumeric(v) Then PptTqq_CoreApi = CLng(v)
End Function

' Contrato (spec 4.9): so chama a API se o core a tem na versao exigida. Pura.
Public Function PptTqq_ApiCompatible(coreApi As Long) As Boolean
    PptTqq_ApiCompatible = (coreApi >= PPT_REQUIRED_API)
End Function

' Core_PptSyncPresentation no Excel. Excel ocupado: o erro sobe (aviso proprio). Core antigo ou API menor: so os dados + aviso. Core com a
' API que falha na chamada: o texto do erro (nao o de core antigo).
Private Function PptTqq_RunCoreSync(xl As Object, core As String, pres As Object, force As Boolean) As String
    Dim v As Variant, en As Long, ed As String
    If Not PptTqq_ApiCompatible(PptTqq_CoreApi(xl, core)) Then PptTqq_RunCoreSync = PptTqq_DataOnly(pres, PptTqq_WorkbookNames(xl)) & vbLf & PPT_MSG_OLD_CORE: Exit Function
    On Error Resume Next
    v = xl.Run("'" & core & "'!Core_PptSyncPresentation", pres, force)
    en = Err.Number: ed = Err.Description
    Err.Clear
    On Error GoTo 0
    If en = 0 Then PptTqq_RunCoreSync = CStr(v): Exit Function
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Then Err.Raise en, "PptTqq", ed
    PptTqq_RunCoreSync = PptTqq_ErrorText(en, ed)
End Function

' Atualiza os graficos ThinQQ da apresentacao (force: inclusive os que nao mudaram). Aviso final com o resumo. Nunca propaga erro.
' Reentrada (evento durante outra pergunta/sincronizacao): ignora.
Public Sub PptTqq_Sync(pres As Object, force As Boolean)
    If mBusy Then Exit Sub
    mBusy = True
    On Error Resume Next
    PptTqq_DoSync pres, force
    mBusy = False
End Sub

Private Sub PptTqq_DoSync(pres As Object, force As Boolean)
    Dim xl As Object, created As Boolean, core As String, msg As String, en As Long, ed As String
    On Error GoTo Fail
    If PptTqq_CountTagged(pres) = 0 Then PptTqq_Msg PPT_MSG_NONE: Exit Sub
    Set xl = PptTqq_GetExcel(created)
    If xl Is Nothing Then PptTqq_Msg PPT_MSG_NO_EXCEL, vbExclamation: Exit Sub
    core = PptTqq_EnsureCore(xl, created)
    If Len(core) > 0 Then
        msg = PptTqq_RunCoreSync(xl, core, pres, force)
    Else
        msg = PptTqq_DataOnly(pres, PptTqq_WorkbookNames(xl)) & vbLf & PPT_MSG_DATA_ONLY
    End If
    PptTqq_Release xl, created
    PptTqq_Msg msg, vbInformation
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptTqq_Release xl, created
    PptTqq_Msg PptTqq_ErrorText(en, ed), vbExclamation
End Sub

Private Function PptTqq_Seen(key As String) As Boolean
    Dim v As Variant
    If mSeen Is Nothing Then Exit Function
    On Error Resume Next
    v = mSeen(key)
    PptTqq_Seen = (Err.Number = 0)
    Err.Clear
End Function

Private Sub PptTqq_MarkSeen(key As String)
    If mSeen Is Nothing Then Set mSeen = New Collection
    On Error Resume Next
    mSeen.Add key, key
    Err.Clear
End Sub

' Evento PresentationCloseFinal: esquece a apresentacao (reabrir depois que os dados mudaram pergunta de novo). Nunca propaga erro.
Public Sub PptTqq_OnClose(pres As Object)
    On Error Resume Next
    If mSeen Is Nothing Then Exit Sub
    mSeen.Remove LCase$(CStr(pres.FullName))
    Err.Clear
End Sub

' Pura: so trata a apresentacao com janela num PowerPoint visivel (abertura por automacao, sem janela, nunca pergunta nem atualiza).
Public Function PptTqq_ShouldHandle(nWindows As Long, appVisible As Boolean) As Boolean
    PptTqq_ShouldHandle = (nWindows > 0 And appVisible)
End Function

' Eventos (AfterPresentationOpen, WindowActivate) e Auto_Open. Sem janela ou PowerPoint invisivel: nada (WindowActivate cobre a janela
' que aparece depois). Erro: so log (nada de MsgBox no caminho dos eventos).
Public Sub PptTqq_OnOpen(pres As Object)
    Dim nw As Long, vis As Boolean
    On Error GoTo Fail
    If BUILD_CONFIG = "Test" Then mTestOnOpenCalls = mTestOnOpenCalls + 1
    nw = pres.Windows.Count
    vis = (CLng(Application.Visible) <> 0)
    If Not PptTqq_ShouldHandle(nw, vis) Then Exit Sub
    PptTqq_HandleOpen pres
    Exit Sub
Fail:
    PptTqq_Log "PptTqq_OnOpen: " & Err.Number & " " & Err.Description
End Sub

' Decisao ao abrir: a apresentacao e marcada como vista na primeira vez (com ou sem graficos marcados: colar depois nao pergunta) e
' esquecida ao fechar. Pergunta/atualiza conforme PptTqq_OnOpenAction. Reentrada ignorada. Nunca propaga erro.
Public Sub PptTqq_HandleOpen(pres As Object)
    Dim n As Long, act As String, key As String, seen As Boolean, md As String, srcs As Collection, remote As Boolean
    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Fail
    key = LCase$(CStr(pres.FullName))
    seen = PptTqq_Seen(key)
    PptTqq_MarkSeen key
    If Not seen Then
        Set srcs = New Collection
        n = PptTqq_CountTagged(pres, srcs)
        md = PptTqq_OpenMode()
        If md = "always" And n > 0 Then remote = PptTqq_HasRemoteSource(srcs, PptTqq_ExcelOpenNames())
        act = PptTqq_OnOpenAction(md, (CLng(pres.ReadOnly) <> 0), n, False, remote)
        If act = "sync" Then
            PptTqq_DoSync pres, False
        ElseIf act = "ask" Then
            If PptTqq_Ask(PptTqq_AskText(n, srcs)) = vbYes Then PptTqq_DoSync pres, False
        End If
    End If
    mBusy = False
    Exit Sub
Fail:
    mBusy = False
    PptTqq_Log "Não foi possível verificar os gráficos do ThinQQ desta apresentação: " & Err.Number & " " & Err.Description
End Sub
