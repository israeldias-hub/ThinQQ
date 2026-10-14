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
Public Const PPT_REQUIRED_API As Long = 1                   ' Core_PptApiVersion minima (spec 4.9); core sem ela ou menor = so os dados
Private Const RPC_E_CALL_REJECTED As Long = -2147418111
Private Const RPC_E_RETRYLATER As Long = -2147417846
Private mEvents As clsPptEvents
Private mSeen As Collection
Private mTestXl As Object
Private mTestCoreDir As String
Private mLastMsg As String

' Carga do suplemento: liga os eventos do PowerPoint (o projeto de teste .pptm nao liga).
Public Sub Auto_Open()
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then Exit Sub
    Set mEvents = New clsPptEvents
    Set mEvents.App = Application
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Set mEvents = Nothing
End Sub

' Gancho de teste: Excel injetado (o do teste, com o core de teste aberto) e pasta das versoes do core. Inerte fora do build Test.
Public Sub PptTqq_SetExcelForTest(xl As Object, coreDir As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    Set mTestXl = xl
    mTestCoreDir = coreDir
End Sub

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

' O que fazer ao abrir: "" (nada), "ask" ou "sync". Pura.
Public Function PptTqq_OnOpenAction(mode As String, isReadOnly As Boolean, nTagged As Long, alreadySeen As Boolean) As String
    If nTagged = 0 Or alreadySeen Or isReadOnly Or mode = "never" Then Exit Function
    If mode = "always" Then PptTqq_OnOpenAction = "sync" Else PptTqq_OnOpenAction = "ask"
End Function

Public Function PptTqq_AskText(n As Long) As String
    PptTqq_AskText = "Esta apresentação tem " & n & " gráfico(s) do ThinQQ vinculado(s) ao Excel. Atualizar agora com os dados atuais?" & vbLf & vbLf & "Para não perguntar mais: aba ThinQQ > Ao abrir."
End Function

Private Function PptTqq_Tag(shp As Object, key As String) As String
    On Error Resume Next
    PptTqq_Tag = CStr(shp.Tags.Item(key))
    Err.Clear
End Function

' Formas soltas marcadas pelo ThinQQ na apresentacao.
Public Function PptTqq_CountTagged(pres As Object) As Long
    Dim sl As Object, shp As Object
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If Len(PptTqq_Tag(shp, PPT_TAG_CHART)) > 0 Then PptTqq_CountTagged = PptTqq_CountTagged + 1
        Next
    Next
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

' Caminho do core mais novo instalado ("" se nao ha).
Private Function PptTqq_CorePath() As String
    Dim d As String, nm As String, names As String, v As String, p As String
    d = PptTqq_CoreDir()
    If Len(d) = 0 Then Exit Function
    On Error Resume Next
    nm = Dir$(d & "\*", vbDirectory)
    Do While Len(nm) > 0
        If nm <> "." And nm <> ".." Then names = names & nm & "|"
        nm = Dir$()
    Loop
    On Error GoTo 0
    v = PptTqq_NewestVersion(names)
    If Len(v) = 0 Then Exit Function
    p = d & "\" & v & "\" & PptTqq_CoreFileName()
    If Len(Dir$(p)) > 0 Then PptTqq_CorePath = p
End Function

Private Function PptTqq_Ping(xl As Object) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = xl.Run("'" & PptTqq_CoreFileName() & "'!Core_Ping")
    PptTqq_Ping = (Err.Number = 0 And CStr(v) = "pong")
    Err.Clear
End Function

' Excel a usar: no build Test, so o injetado. Senao: o aberto (GetObject) SE ele tem o core; senao uma instancia nova e invisivel (created).
Private Function PptTqq_GetExcel(ByRef created As Boolean) As Object
    Dim xl As Object
    created = False
    If BUILD_CONFIG = "Test" Then Set PptTqq_GetExcel = mTestXl: Exit Function
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    If Not xl Is Nothing Then
        If PptTqq_Ping(xl) Then Set PptTqq_GetExcel = xl: Exit Function
    End If
    Set xl = Nothing
    Err.Clear
    Set xl = CreateObject("Excel.Application")
    If xl Is Nothing Then Exit Function
    On Error GoTo 0
    created = True
    xl.Visible = False
    xl.DisplayAlerts = False
    Set PptTqq_GetExcel = xl
End Function

' Nome do core respondendo no Excel ("" sem core). Na instancia criada por nos, abre o core mais novo instalado (somente leitura).
Private Function PptTqq_EnsureCore(xl As Object, created As Boolean) As String
    Dim p As String
    If PptTqq_Ping(xl) Then PptTqq_EnsureCore = PptTqq_CoreFileName(): Exit Function
    If Not created Then Exit Function
    p = PptTqq_CorePath()
    If Len(p) = 0 Then Exit Function
    On Error Resume Next
    xl.Workbooks.Open p, 0, True
    Err.Clear
    On Error GoTo 0
    If PptTqq_Ping(xl) Then PptTqq_EnsureCore = PptTqq_CoreFileName()
End Function

Private Sub PptTqq_Release(xl As Object, created As Boolean)
    On Error Resume Next
    If created And Not xl Is Nothing Then
        xl.DisplayAlerts = False
        xl.Quit
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

' Modo degradado (sem o core): so os dados pelo vinculo nativo, e so dos graficos cuja pasta existe.
Public Function PptTqq_DataOnly(pres As Object) As String
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
                If PptTqq_FileExists(p) Then
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

' Versao da API do core (0 = core sem Core_PptApiVersion, anterior a 1.4.0). Pura quanto ao suplemento: so le.
Private Function PptTqq_CoreApi(xl As Object, core As String) As Long
    Dim v As Variant
    On Error Resume Next
    v = xl.Run("'" & core & "'!Core_PptApiVersion")
    If Err.Number = 0 And IsNumeric(v) Then PptTqq_CoreApi = CLng(v)
    Err.Clear
End Function

' Contrato (spec 4.9): so chama a API se o core a tem na versao exigida. Pura.
Public Function PptTqq_ApiCompatible(coreApi As Long) As Boolean
    PptTqq_ApiCompatible = (coreApi >= PPT_REQUIRED_API)
End Function

' Core_PptSyncPresentation no Excel. Excel ocupado: o erro sobe (aviso proprio). Core antigo ou API menor: so os dados + aviso.
Private Function PptTqq_RunCoreSync(xl As Object, core As String, pres As Object, force As Boolean) As String
    Dim v As Variant, en As Long, ed As String
    If Not PptTqq_ApiCompatible(PptTqq_CoreApi(xl, core)) Then PptTqq_RunCoreSync = PptTqq_DataOnly(pres) & vbLf & PPT_MSG_OLD_CORE: Exit Function
    On Error Resume Next
    v = xl.Run("'" & core & "'!Core_PptSyncPresentation", pres, force)
    en = Err.Number: ed = Err.Description
    Err.Clear
    On Error GoTo 0
    If en = 0 Then PptTqq_RunCoreSync = CStr(v): Exit Function
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Then Err.Raise en, "PptTqq", ed
    PptTqq_RunCoreSync = PptTqq_DataOnly(pres) & vbLf & PPT_MSG_OLD_CORE
End Function

' Atualiza os graficos ThinQQ da apresentacao (force: inclusive os que nao mudaram). Aviso final com o resumo. Nunca propaga erro.
Public Sub PptTqq_Sync(pres As Object, force As Boolean)
    Dim xl As Object, created As Boolean, core As String, msg As String, en As Long, ed As String
    On Error GoTo Fail
    If PptTqq_CountTagged(pres) = 0 Then PptTqq_Msg PPT_MSG_NONE: Exit Sub
    Set xl = PptTqq_GetExcel(created)
    If xl Is Nothing Then PptTqq_Msg PPT_MSG_NO_EXCEL, vbExclamation: Exit Sub
    core = PptTqq_EnsureCore(xl, created)
    If Len(core) > 0 Then
        msg = PptTqq_RunCoreSync(xl, core, pres, force)
    Else
        msg = PptTqq_DataOnly(pres) & vbLf & PPT_MSG_DATA_ONLY
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

' Gancho de teste: esquece as apresentacoes ja vistas nesta sessao.
Public Sub PptTqq_ResetSeenForTest()
    Set mSeen = Nothing
End Sub

' Eventos (AfterPresentationOpen e WindowActivate): uma vez por apresentacao (caminho) e por sessao. Nunca propaga erro.
Public Sub PptTqq_OnOpen(pres As Object)
    Dim n As Long, act As String, key As String
    On Error GoTo Fail
    key = LCase$(CStr(pres.FullName))
    n = PptTqq_CountTagged(pres)
    act = PptTqq_OnOpenAction(PptTqq_OpenMode(), (CLng(pres.ReadOnly) <> 0), n, PptTqq_Seen(key))
    If n > 0 Then PptTqq_MarkSeen key
    Select Case act
        Case "sync"
            PptTqq_Sync pres, False
        Case "ask"
            If PptTqq_Ask(PptTqq_AskText(n)) = vbYes Then PptTqq_Sync pres, False
    End Select
    Exit Sub
Fail:
    PptTqq_Msg "Não foi possível verificar os gráficos do ThinQQ desta apresentação: " & Err.Description, vbExclamation
End Sub
