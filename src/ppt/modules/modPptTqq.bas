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
Public Const PPT_MSG_BUSY As String = "O Excel está ocupado (uma célula em edição ou uma caixa de diálogo aberta). Termine o que está fazendo no Excel e tente de novo."
Public Const PPT_MAX_SOURCES_SHOWN As Long = 5                ' pastas de origem listadas na pergunta ao abrir (depois, "...")
Public Const PPT_MAX_PATH_SHOWN As Long = 150                 ' caracteres de cada pasta na pergunta (cortada no meio)
Public Const PPT_MAX_ASK_LEN As Long = 900                    ' tamanho maximo da pergunta (o MsgBox corta perto de 1024)
Public Const PPT_REQUIRED_API As Long = 1                   ' Core_PptApiVersion minima (spec 4.9); core sem ela ou menor = so os dados
Public Const PPT_TAG_AUTO As String = "TQQ_AUTO"
Public Const PPT_TAG_INDEX As String = "TQQ_AUTOIDX"
Public Const PPT_API_SHAPES As Long = 2                     ' Core_PptSyncShapes (core 1.6.0); conferida POR RECURSO (ao abrir basta a 1)
Public Const PPT_IDX_MAX_ENTRIES As Long = 2000             ' = PPTIDX_MAX_ENTRIES do core
Public Const PPT_IDX_MAX_LEN As Long = 1000000              ' = PPTIDX_MAX_LEN do core
Private Const PPT_IDX_FS As String = "|"                    ' indice: separador de campos (= PPTIDX_FS do core)
Private Const PPT_IDX_ES As String = vbLf                   ' indice: separador de entradas (= PPTIDX_ES do core)
Public Const PPT_MSG_SELECT As String = "Selecione um ou mais gráficos do ThinQQ vinculados ao Excel."
Public Const PPT_MSG_AUTO_OLDCORE As String = "Atualize o ThinQQ no Excel para a atualização automática funcionar (a marca foi gravada)."
Public Const PPT_MSG_AUTO_NOCORE As String = "Marcado para atualizar automaticamente. Para os gráficos acompanharem a pasta de origem, abra o Excel com o ThinQQ (1.6.0 ou mais novo)."
Public Const PPT_MSG_AUTO_OFF As String = "A atualização automática está desligada no Excel: aba ThinQQ > PowerPoint > Atualização automática ligada."
Public Const PPT_MSG_SEL_OLDCORE As String ="Atualize o ThinQQ no Excel para atualizar só os gráficos selecionados (desta vez a apresentação inteira foi atualizada)."
Public Const PPT_TAG_DATA As String = "TQQ_DATA"
Public Const PPT_TAG_DATABOOK As String = "TQQ_DATABOOK"
Public Const PPT_DATA_SUFFIX As String = " - dados ThinQQ.xlsx"
Public Const PPT_API_DATA As Long = 3                       ' Core_PptInsertChart, Core_PptEditData, Core_PptDataSave (core 1.8.0)
Public Const PPT_MSG_SAVE_FIRST As String = "Salve a apresentação primeiro: os dados do gráfico ficam numa pasta do Excel ao lado dela (""<nome> - dados ThinQQ.xlsx"")."
Public Const PPT_MSG_WEB_PRES As String = "A apresentação está aberta por um endereço da Web (OneDrive/SharePoint). Salve-a numa pasta do computador (ou na pasta sincronizada do OneDrive) para inserir gráficos do ThinQQ."
Public Const PPT_MSG_DATA_NOCORE As String = "O ThinQQ não está instalado no Excel deste computador: não é possível criar nem editar gráficos do ThinQQ."
Public Const PPT_MSG_OLD_CORE_DATA As String = "Atualize o ThinQQ no Excel para criar gráficos pelo PowerPoint."
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
Private mEditXl As Object                       ' Excel criado aqui para criar/editar dados (fica com o usuario; atalho para reencontra-lo)
Private mHintShown As Boolean                   ' aviso de primeira vez de "Inserir grafico" ja mostrado nesta sessao

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
    mHintShown = False
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

' Pergunta ao abrir. sources (opcional): as pastas de origem distintas, listadas antes da dica (PptTqq_SourcesText: remotas primeiro, no
' maximo PPT_MAX_SOURCES_SHOWN, cada uma limpa e com no maximo PPT_MAX_PATH_SHOWN caracteres). A pergunta inteira fica em
' PPT_MAX_ASK_LEN caracteres (o MsgBox corta perto de 1024): a frase da pergunta e a dica sempre aparecem.
Public Function PptTqq_AskText(n As Long, Optional sources As Collection) As String
    Dim head As String, tail As String
    head = "Esta apresentação tem " & n & " gráfico(s) do ThinQQ vinculado(s) ao Excel. Atualizar agora com os dados atuais?"
    tail = vbLf & vbLf & "Para não perguntar mais: aba ThinQQ > Ao abrir."
    PptTqq_AskText = head & PptTqq_SourcesText(sources, PPT_MAX_ASK_LEN - Len(head) - Len(tail)) & tail
End Function

' Bloco "Pastas de origem:" da pergunta (vazio sem pastas): remotas primeiro (na ordem em que aparecem), depois as locais; para no
' PPT_MAX_SOURCES_SHOWN-esimo caminho ou quando o proximo passaria do limite budget, com "..." se sobrou alguma. Pura.
Private Function PptTqq_SourcesText(sources As Collection, budget As Long) As String
    Dim ordered As New Collection, v As Variant, s As String, ln As String, k As Long, pass As Long
    If sources Is Nothing Then Exit Function
    If sources.Count = 0 Then Exit Function
    For pass = 1 To 2
        For Each v In sources
            If PptTqq_IsRemotePath(CStr(v)) = (pass = 1) Then ordered.Add PptTqq_CleanPath(CStr(v))
        Next
    Next
    s = vbLf & vbLf & "Pastas de origem:"
    For Each v In ordered
        ln = vbLf & CStr(v)
        If k >= PPT_MAX_SOURCES_SHOWN Or Len(s) + Len(ln) + 4 > budget Then s = s & vbLf & "...": Exit For
        s = s & ln
        k = k + 1
    Next
    PptTqq_SourcesText = s
End Function

' Caminho para exibir: cada caractere de controle (codigo < 32 ou 127) vira espaco, espacos repetidos viram um, e acima de
' PPT_MAX_PATH_SHOWN caracteres fica o comeco e o fim com "..." no meio. Pura.
Public Function PptTqq_CleanPath(p As String) As String
    Dim i As Long, c As Long, s As String, keepHead As Long
    s = p
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If (c >= 0 And c < 32) Or c = 127 Then Mid$(s, i, 1) = " "
    Next
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)
    If Len(s) > PPT_MAX_PATH_SHOWN Then
        keepHead = (PPT_MAX_PATH_SHOWN - 3) \ 2
        s = Left$(s, keepHead) & "..." & Right$(s, PPT_MAX_PATH_SHOWN - 3 - keepHead)
    End If
    PptTqq_CleanPath = s
End Function

' Origem remota: caminho de rede (\\servidor\..., inclusive \\?\UNC\ e WebDAV \\servidor@SSL\DavWWWRoot\), com barras normais ou
' mistas (//servidor/...), endereco http(s) ou URL file: (sempre remota, por cautela). \\?\C:\ (unidade local com prefixo longo) e
' letras de unidade sao locais: uma unidade de rede mapeada (Z:\) nao e distinguivel aqui (limite documentado). Pura.
Public Function PptTqq_IsRemotePath(p As String) As Boolean
    Dim t As String
    t = Replace(LCase$(Trim$(p)), "/", "\")
    If Left$(t, 5) = "file:" Or Left$(t, 7) = "http:\\" Or Left$(t, 8) = "https:\\" Then PptTqq_IsRemotePath = True: Exit Function
    If Left$(t, 2) <> "\\" Then Exit Function
    If Left$(t, 8) = "\\?\unc\" Then PptTqq_IsRemotePath = True: Exit Function
    If (Left$(t, 4) = "\\?\" Or Left$(t, 4) = "\\.\") And Mid$(t, 6, 1) = ":" Then Exit Function    ' \\?\C:\ ou \\.\C:\: unidade local
    PptTqq_IsRemotePath = True
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
' instancia criada nunca fica sem dono. noCreate: so o Excel ja aberto (com ou sem o core), nunca uma instancia nova (Nothing se nao ha).
Private Function PptTqq_GetExcel(ByRef created As Boolean, Optional noCreate As Boolean = False) As Object
    Dim xl As Object, st As Long
    created = False
    If BUILD_CONFIG = "Test" Then
        created = mTestCreated And Not noCreate
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
        If noCreate Then Set PptTqq_GetExcel = xl: Exit Function
    End If
    If noCreate Then Exit Function
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
' API que falha na chamada: o texto do erro (nao o de core antigo). api: a versao ja lida (-1 = ler agora).
Private Function PptTqq_RunCoreSync(xl As Object, core As String, pres As Object, force As Boolean, Optional api As Long = -1) As String
    Dim v As Variant, en As Long, ed As String
    If api < 0 Then api = PptTqq_CoreApi(xl, core)
    If Not PptTqq_ApiCompatible(api) Then PptTqq_RunCoreSync = PptTqq_DataOnly(pres, PptTqq_WorkbookNames(xl)) & vbLf & PPT_MSG_OLD_CORE: Exit Function
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
        PptTqq_RelinkData pres                    ' apresentacao movida: graficos criados pelo PowerPoint voltam a achar os dados
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

' ---- 1.6.0: atualizacao automatica por elemento. Indice: COPIA EXATA das PptIdx_* do core (mesmos testes de cadeia). Puras e totais: a
' etiqueta vem da apresentacao (de quem a fez), entao lixo nunca levanta erro, o SlideID so vale com 1 a 9 algarismos, entradas alem de
' PPT_IDX_MAX_ENTRIES sao ignoradas e uma etiqueta maior que PPT_IDX_MAX_LEN e recusada inteira. Entradas por vbLf, campos por "|" (o .pptx
' salvo perde Chr(30)/Chr(31) e troca Tab por espaco: medido no core). Os separadores nunca entram num campo. ----

Public Function PptTqq_IdxSafeField(s As String) As String
    Dim t As String
    t = Replace(Replace(Replace(s, PPT_IDX_FS, "_"), vbLf, "_"), vbCr, "_")
    PptTqq_IdxSafeField = Replace(Replace(Replace(t, vbTab, "_"), Chr$(30), "_"), Chr$(31), "_")
End Function

Public Function PptTqq_IdxEntry(slideId As Long, shapeName As String, book As String) As String
    PptTqq_IdxEntry = CStr(slideId) & PPT_IDX_FS & PptTqq_IdxSafeField(shapeName) & PPT_IDX_FS & PptTqq_IdxSafeField(book)
End Function

Public Function PptTqq_IdxParse(idx As String) As Collection
    Dim out As New Collection, parts() As String, f() As String, i As Long
    Set PptTqq_IdxParse = out
    On Error GoTo Done
    If Len(idx) = 0 Or Len(idx) > PPT_IDX_MAX_LEN Then Exit Function
    parts = Split(idx, PPT_IDX_ES)
    For i = LBound(parts) To UBound(parts)
        If out.Count >= PPT_IDX_MAX_ENTRIES Then Exit For
        If Len(parts(i)) > 0 Then
            f = Split(parts(i), PPT_IDX_FS)
            If UBound(f) = 2 Then
                If Len(f(0)) > 0 And Len(f(0)) < 10 And Len(f(1)) > 0 Then
                    If Not f(0) Like "*[!0-9]*" Then out.Add Array(CLng(f(0)), f(1), f(2))
                End If
            End If
        End If
    Next
Done:
End Function

Public Function PptTqq_IdxJoin(entries As Collection) As String
    Dim e As Variant, a() As String, n As Long
    If entries.Count = 0 Then Exit Function
    ReDim a(0 To entries.Count - 1)
    For Each e In entries
        a(n) = PptTqq_IdxEntry(CLng(e(0)), CStr(e(1)), CStr(e(2)))
        n = n + 1
    Next
    PptTqq_IdxJoin = Join(a, PPT_IDX_ES)
End Function

Public Function PptTqq_IdxSet(idx As String, slideId As Long, shapeName As String, book As String, turnOn As Boolean) As String
    Dim out As New Collection, e As Variant, nm As String
    nm = PptTqq_IdxSafeField(shapeName)
    For Each e In PptTqq_IdxParse(idx)
        If Not (CLng(e(0)) = slideId And StrComp(CStr(e(1)), nm, vbBinaryCompare) = 0) Then out.Add e
    Next
    If turnOn Then out.Add Array(slideId, nm, book)
    PptTqq_IdxSet = PptTqq_IdxJoin(out)
End Function

Private Sub PptTqq_TagDelete(o As Object, key As String)
    On Error Resume Next
    o.Tags.Delete key
    Err.Clear
End Sub

' Medido na C3: Shapes(nome) do PowerPoint acha a forma IGNORANDO a caixa ("tqq_A" e "tqq_a" no slide: os dois nomes acham a primeira), entao
' a comparacao de nomes aqui e sem diferenciar maiusculas (vbTextCompare), como no core.
Private Function PptTqq_SlideHasName(sl As Object, nm As String) As Boolean
    Dim s As Object
    For Each s In sl.Shapes
        If StrComp(CStr(s.Name), nm, vbTextCompare) = 0 Then PptTqq_SlideHasName = True: Exit Function
    Next
End Function

' Copia de PptSync_EnsureUniqueName do core: separadores do indice e caracteres que o .pptx salvo muda viram "_" (PptTqq_IdxSafeField; medido
' na revisao da C1: o PowerPoint ACEITA Chr(30) no nome de uma forma) e nome repetido no slide (sem diferenciar maiusculas; outra forma =
' Shape.Id diferente) vira "_2", "_3"...
Private Sub PptTqq_EnsureUniqueName(shp As Object)
    Dim sl As Object, s As Object, base As String, n As Long, clash As Boolean, myId As Long
    Set sl = shp.Parent
    base = CStr(shp.Name): myId = CLng(shp.Id)
    If PptTqq_IdxSafeField(base) <> base Then
        base = PptTqq_IdxSafeField(base)
        If Not PptTqq_SlideHasName(sl, base) Then shp.Name = base: Exit Sub
        clash = True
    End If
    If Not clash Then
        For Each s In sl.Shapes
            If CLng(s.Id) <> myId And StrComp(CStr(s.Name), base, vbTextCompare) = 0 Then clash = True: Exit For
        Next
    End If
    If Not clash Then Exit Sub
    For n = 2 To 999
        If Not PptTqq_SlideHasName(sl, base & "_" & n) Then shp.Name = base & "_" & n: Exit Sub
    Next
End Sub

' Formas do ThinQQ (TQQ_CHART) da selecao da janela ativa (formas ou texto dentro de uma forma), so as de um slide (mestre/layout nao tem
' SlideID e o indice nao as acha). Vazia sem janela/selecao. Nunca propaga erro (chamada pelos callbacks da faixa a cada mudanca de selecao).
Public Function PptTqq_SelectedTagged() As Collection
    Dim out As New Collection, sel As Object, sr As Object, shp As Object, i As Long, sid As Long
    Set PptTqq_SelectedTagged = out
    On Error Resume Next
    If Application.Windows.Count = 0 Then Exit Function
    Set sel = Application.ActiveWindow.Selection
    If sel Is Nothing Then Exit Function
    If sel.Type <> 2 And sel.Type <> 3 Then Exit Function          ' ppSelectionShapes, ppSelectionText
    Set sr = sel.ShapeRange
    If sr Is Nothing Then Exit Function
    For i = 1 To sr.Count
        Set shp = Nothing
        Set shp = sr(i)
        If Not shp Is Nothing Then
            If Len(PptTqq_Tag(shp, PPT_TAG_CHART)) > 0 Then
                Err.Clear
                sid = CLng(shp.Parent.SlideID)
                If Err.Number = 0 Then out.Add shp
            End If
        End If
        Err.Clear
    Next
    Err.Clear
End Function

' Todas as formas dadas sao automaticas (e ha pelo menos uma).
Public Function PptTqq_ShapesAuto(shapes As Collection) As Boolean
    Dim s As Object
    If shapes.Count = 0 Then Exit Function
    For Each s In shapes
        If PptTqq_Tag(s, PPT_TAG_AUTO) <> "1" Then Exit Function
    Next
    PptTqq_ShapesAuto = True
End Function

' Liga/desliga "Atualizar automaticamente" (etiqueta + indice da apresentacao), sem o Excel. Mesma ordem do core (PptSync_SetAuto): o nome
' fica definitivo antes do indice e a etiqueta da forma (a verdade) e gravada antes do indice (o atalho, reconstruido pela sincronizacao).
Public Sub PptTqq_SetAutoShapes(shapes As Collection, turnOn As Boolean)
    Dim s As Object, pres As Object, idx As String, book As String
    For Each s In shapes
        Set pres = s.Parent.Parent
        If turnOn Then PptTqq_EnsureUniqueName s
        If turnOn Then s.Tags.Add PPT_TAG_AUTO, "1" Else PptTqq_TagDelete s, PPT_TAG_AUTO
        book = PptTqq_SourcePath(PptTqq_LinkSource(s), PptTqq_Tag(s, PPT_TAG_BOOK))
        idx = PptTqq_IdxSet(PptTqq_Tag(pres, PPT_TAG_INDEX), CLng(s.Parent.SlideID), CStr(s.Name), book, turnOn)
        If Len(idx) = 0 Then PptTqq_TagDelete pres, PPT_TAG_INDEX Else pres.Tags.Add PPT_TAG_INDEX, idx
    Next
End Sub

' Entradas (formato do indice, pasta vazia) das formas dadas: o que Core_PptSyncShapes recebe.
Public Function PptTqq_EntriesOf(shapes As Collection) As String
    Dim c As New Collection, s As Object
    For Each s In shapes
        c.Add Array(CLng(s.Parent.SlideID), CStr(s.Name), "")
    Next
    PptTqq_EntriesOf = PptTqq_IdxJoin(c)
End Function

' O core tem Core_PptSyncShapes (API por recurso: a sincronizacao ao abrir so exige PPT_REQUIRED_API). Pura.
Public Function PptTqq_ShapesApiOk(coreApi As Long) As Boolean
    PptTqq_ShapesApiOk = (coreApi >= PPT_API_SHAPES)
End Function

' Atualiza as entradas dadas. Core com API 2 => Core_PptSyncShapes (o texto do core, inclusive o de PowerPoint ocupado, e exibido como veio).
' "Atualizar uma vez" (fromAuto False): core mais velho => a apresentacao inteira com UM aviso (sem Core_PptApiVersion: so os dados e o aviso
' de core antigo; API 1: Core_PptSyncPresentation e o aviso de "so os selecionados"); sem o core => so os dados; sem Excel => cria um.
' "Atualizar automaticamente" (fromAuto, chamado depois de gravar a marca; design secao 8): so o Excel ja aberto (nunca cria um); sem Excel ou
' sem o core => so a dica de abrir o Excel com o ThinQQ; core mais velho => so o aviso de atualizar o ThinQQ (nada e atualizado); API 2 =>
' atualiza e, se a atualizacao automatica esta desligada no Excel, acrescenta onde liga-la. A API e lida uma vez. Reentrada ignorada;
' mBusy liberado em todo caminho antes do aviso. Nunca propaga.
Public Sub PptTqq_SyncEntries(pres As Object, entries As String, force As Boolean, Optional fromAuto As Boolean = False)
    Dim xl As Object, created As Boolean, core As String, msg As String, v As Variant, en As Long, ed As String, api As Long
    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Fail
    Set xl = PptTqq_GetExcel(created, fromAuto)
    If xl Is Nothing Then
        mBusy = False
        If fromAuto Then PptTqq_Msg PPT_MSG_AUTO_NOCORE, vbInformation Else PptTqq_Msg PPT_MSG_NO_EXCEL, vbExclamation
        Exit Sub
    End If
    core = PptTqq_EnsureCore(xl, created)
    If Len(core) = 0 Then
        If fromAuto Then msg = PPT_MSG_AUTO_NOCORE Else msg = PptTqq_DataOnly(pres, PptTqq_WorkbookNames(xl)) & vbLf & PPT_MSG_DATA_ONLY
    Else
        api = PptTqq_CoreApi(xl, core)
        If PptTqq_ShapesApiOk(api) Then
            v = xl.Run("'" & core & "'!Core_PptSyncShapes", pres, entries, force)
            msg = CStr(v)
            If fromAuto Then
                If Not PptTqq_CoreAutoOn(xl, core) Then msg = msg & vbLf & PPT_MSG_AUTO_OFF
            End If
        ElseIf fromAuto Then
            msg = PPT_MSG_AUTO_OLDCORE
        ElseIf PptTqq_ApiCompatible(api) Then
            msg = PptTqq_RunCoreSync(xl, core, pres, force, api) & vbLf & PPT_MSG_SEL_OLDCORE
        Else
            msg = PptTqq_RunCoreSync(xl, core, pres, force, api)          ' so os dados + PPT_MSG_OLD_CORE (um aviso so)
        End If
    End If
    PptTqq_Release xl, created
    mBusy = False
    PptTqq_Msg msg, vbInformation
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptTqq_Release xl, created
    mBusy = False
    PptTqq_Msg PptTqq_ErrorText(en, ed), vbExclamation
End Sub

' Opcao "autoon" do core (API 2). Erro ou Excel ocupado: True (nenhuma nota; o resumo ja saiu). Nunca propaga.
Private Function PptTqq_CoreAutoOn(xl As Object, core As String) As Boolean
    Dim v As Variant
    PptTqq_CoreAutoOn = True
    On Error Resume Next
    v = xl.Run("'" & core & "'!Core_PptGetFlag", "autoon")
    If Err.Number = 0 Then PptTqq_CoreAutoOn = CBool(v)
    Err.Clear
End Function

' Botao "Atualizar automaticamente": grava a marca (e o indice) e, ao ligar, ja atualiza uma vez so essas formas (como o "Update
' Automatically" do think-cell). Desligar nao chama o Excel. Nunca propaga erro (falha ao gravar a marca: aviso proprio).
Public Sub PptTqq_ToggleAuto(shapes As Collection, turnOn As Boolean)
    Dim pres As Object, en As Long, ed As String
    If shapes.Count = 0 Then PptTqq_Msg PPT_MSG_SELECT: Exit Sub
    On Error GoTo Fail
    Set pres = shapes(1).Parent.Parent
    PptTqq_SetAutoShapes shapes, turnOn
    On Error GoTo 0
    If turnOn Then PptTqq_SyncEntries pres, PptTqq_EntriesOf(shapes), True, True
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    If turnOn Then
        PptTqq_Msg "Não foi possível marcar o gráfico como automático: " & ed & " (" & en & ")", vbExclamation
    Else
        PptTqq_Msg "Não foi possível voltar o gráfico para a atualização manual: " & ed & " (" & en & ")", vbExclamation
    End If
End Sub

' Botao "Atualizar uma vez": deixa o nome de cada forma unico no slide ANTES de montar as entradas (um grafico copiado e colado no mesmo
' slide tem o mesmo nome, e a busca do core por nome ignora a caixa: atualizaria a outra forma) e atualiza so essas formas. Nunca propaga.
Public Sub PptTqq_UpdateSelected(shapes As Collection)
    Dim s As Object, en As Long, ed As String
    If shapes.Count = 0 Then PptTqq_Msg PPT_MSG_SELECT: Exit Sub
    On Error GoTo Fail
    For Each s In shapes
        PptTqq_EnsureUniqueName s
    Next
    On Error GoTo 0
    PptTqq_SyncEntries shapes(1).Parent.Parent, PptTqq_EntriesOf(shapes), True
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    PptTqq_Msg PptTqq_ErrorText(en, ed), vbExclamation
End Sub

' ---- 1.8.0 (Parte G): criar graficos pelo PowerPoint. spec 2026-10-15 secao 14. ----

' Copia exata de PptData_PathFor do core (mesmos testes de cadeia). Pura.
Public Function PptTqq_DataPathFor(presFullName As String) As String
    Dim p As Long, k As Long, base As String, f As String
    f = Trim$(presFullName)
    If LCase$(Left$(f, 7)) = "http://" Or LCase$(Left$(f, 8)) = "https://" Then Exit Function
    p = InStrRev(f, "\")
    If p < 2 Then Exit Function
    base = Mid$(f, p + 1)
    k = InStrRev(base, ".")
    If k > 1 Then base = Left$(base, k - 1)
    PptTqq_DataPathFor = Left$(f, p - 1) & "\" & base & PPT_DATA_SUFFIX
End Function

' Slide ativo da apresentacao (modo Normal; senao o selecionado). Nothing em modos de mestre ou sem janela.
Public Function PptTqq_ActiveSlide(pres As Object) As Object
    Dim s As Object, k As Long
    On Error Resume Next
    Set s = Application.ActiveWindow.View.Slide
    k = 0: k = s.SlideIndex
    If k = 0 Then Set s = Application.ActiveWindow.Selection.SlideRange(1): k = s.SlideIndex
    If k = 0 Then Set s = Nothing
    If Not s Is Nothing Then
        If StrComp(CStr(s.Parent.FullName), CStr(pres.FullName), vbTextCompare) <> 0 Then Set s = Nothing
    End If
    Err.Clear
    Set PptTqq_ActiveSlide = s
End Function

Private Sub PptTqq_OpenLoader(xl As Object)
    Dim p As String
    On Error Resume Next
    p = Environ$("LOCALAPPDATA") & "\ThinQQ\ThinQQ.xlam"
    If Len(Dir$(p)) = 0 Then Exit Sub
    xl.Workbooks.Open p
    Err.Clear
End Sub

' Excel para criar/editar dados, que FICA com o usuario (spec 14.3): (1) o criado antes por aqui, se ainda responde; (2) o Excel aberto COM o
' core; (3) senao um novo, VISIVEL, UserControl, com o loader instalado aberto (carrega core e faixa) ou, sem loader, o core pelo caminho de
' sempre. Nunca carrega o ThinQQ num Excel do usuario que nao o tem; nunca da Quit no Excel criado aqui (sonda PG1: ele sobrevive ao fim da
' referencia e ao fechamento do PowerPoint; mEditXl so serve de atalho para reencontra-lo). Ocupado: erro de ocupado.
Public Function PptTqq_GetEditExcel(ByRef core As String) As Object
    Dim xl As Object, st As Long
    core = ""
    If BUILD_CONFIG = "Test" Then
        If mTestXl Is Nothing Then Exit Function
        core = PptTqq_EnsureCore(mTestXl, mTestCreated)
        Set PptTqq_GetEditExcel = mTestXl
        Exit Function
    End If
    If Not mEditXl Is Nothing Then
        If PptTqq_PingState(mEditXl) = 1 Then core = PptTqq_CoreFileName(): Set PptTqq_GetEditExcel = mEditXl: Exit Function
        Set mEditXl = Nothing                                       ' o usuario fechou
    End If
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    Err.Clear
    On Error GoTo 0
    If Not xl Is Nothing Then
        st = PptTqq_PingState(xl)
        If st = 1 Then core = PptTqq_CoreFileName(): Set PptTqq_GetEditExcel = xl: Exit Function
        If st = -1 Then PptTqq_RaiseBusy
    End If
    Set xl = Nothing
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    Err.Clear
    On Error GoTo 0
    If xl Is Nothing Then Exit Function
    Set mEditXl = xl
    On Error Resume Next
    xl.Visible = True
    xl.UserControl = True
    Err.Clear
    On Error GoTo 0
    PptTqq_OpenLoader xl
    If PptTqq_PingState(xl) = 1 Then core = PptTqq_CoreFileName() Else core = PptTqq_EnsureCore(xl, True)
    Set PptTqq_GetEditExcel = xl
End Function

' Texto de falha de criar/editar dados: ocupado e Excel ausente com os textos de sempre; o resto com o prefixo da acao. Pura.
Private Function PptTqq_DataErrorText(prefix As String, en As Long, ed As String) As String
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Or en = 429 Then
        PptTqq_DataErrorText = PptTqq_ErrorText(en, ed)
    Else
        PptTqq_DataErrorText = prefix & ed & " (" & en & ")"
    End If
End Function

' Abre a planilha dos dados da forma no Excel dado e traz a janela a frente (AppActivate com a legenda devolvida pelo core; sonda PG2,
' verificacao manual). "!" do core = a mensagem. Falha: aviso proprio (o grafico ja inserido continua). Nunca propaga.
Private Sub PptTqq_EditShape(shp As Object, xl As Object, core As String)
    Dim cap As String, c As New Collection, en As Long, ed As String
    On Error GoTo Fail
    c.Add shp
    cap = CStr(xl.Run("'" & core & "'!Core_PptEditData", shp.Parent.Parent, PptTqq_EntriesOf(c)))
    If Left$(cap, 1) = "!" Then PptTqq_Msg Mid$(cap, 2), vbExclamation: Exit Sub
    If BUILD_CONFIG = "Test" Then mLastMsg = "edit:" & cap: Exit Sub
    On Error Resume Next
    AppActivate cap
    If Err.Number <> 0 Then Err.Clear: PptTqq_Msg "O Excel abriu a planilha de dados do gráfico (veja a barra de tarefas)."
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    PptTqq_Msg PptTqq_DataErrorText("Não foi possível abrir os dados do gráfico: ", en, ed), vbExclamation
End Sub

' Aviso unico por sessao depois do primeiro grafico criado.
Private Sub PptTqq_FirstHint(pres As Object)
    Dim d As String
    If mHintShown Then Exit Sub
    mHintShown = True
    d = PptTqq_DataPathFor(CStr(pres.FullName))
    PptTqq_Msg "Digite os dados na planilha do Excel que abriu: o slide acompanha em cerca de 2 segundos." & vbLf & _
               "Os dados ficam em '" & Mid$(d, InStrRev(d, "\") + 1) & "', ao lado da apresentação: envie os dois arquivos juntos (ThinQQ > Empacotar para enviar)."
End Sub

' "Inserir grafico" (1.8.0): apresentacao salva (senao a mensagem), slide de destino (o dado; Nothing = um slide novo no fim), Excel para editar
' (fica com o usuario), core com a API 3; o core cria o grafico na pasta de dados e cola no slide; depois abre a planilha dos dados. Nunca propaga.
Public Sub PptTqq_InsertChartIn(pres As Object, slide As Object, kind As String)
    Dim xl As Object, core As String, msg As String, sl As Object, n0 As Long, en As Long, ed As String
    If mBusy Then Exit Sub
    On Error GoTo Fail
    If Len(CStr(pres.Path)) = 0 Then PptTqq_Msg PPT_MSG_SAVE_FIRST: Exit Sub
    If Len(PptTqq_DataPathFor(CStr(pres.FullName))) = 0 Then PptTqq_Msg PPT_MSG_WEB_PRES: Exit Sub
    mBusy = True
    Set xl = PptTqq_GetEditExcel(core)
    If xl Is Nothing Then mBusy = False: PptTqq_Msg PPT_MSG_NO_EXCEL, vbExclamation: Exit Sub
    If Len(core) = 0 Then mBusy = False: PptTqq_Msg PPT_MSG_DATA_NOCORE, vbExclamation: Exit Sub
    If PptTqq_CoreApi(xl, core) < PPT_API_DATA Then mBusy = False: PptTqq_Msg PPT_MSG_OLD_CORE_DATA: Exit Sub
    Set sl = slide
    If sl Is Nothing Then Set sl = pres.Slides.Add(pres.Slides.Count + 1, 12)
    n0 = sl.Shapes.Count
    msg = CStr(xl.Run("'" & core & "'!Core_PptInsertChart", sl, kind))
    mBusy = False
    If Len(msg) > 0 Then PptTqq_Msg msg, vbExclamation: Exit Sub
    If sl.Shapes.Count > n0 Then PptTqq_EditShape sl.Shapes(sl.Shapes.Count), xl, core
    PptRx_InvalidateElem
    PptTqq_FirstHint pres
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    mBusy = False
    PptTqq_Msg PptTqq_DataErrorText("Não foi possível criar o gráfico: ", en, ed), vbExclamation
End Sub

' "Editar dados" das formas dadas (a 1a): reaponta antes (apresentacao movida), Excel para editar, Core_PptEditData. Nunca propaga.
Public Sub PptTqq_EditShapes(pres As Object, shapes As Collection)
    Dim xl As Object, core As String, en As Long, ed As String
    If mBusy Then Exit Sub
    If shapes.Count = 0 Then PptTqq_Msg PPT_MSG_SELECT: Exit Sub
    On Error GoTo Fail
    PptTqq_RelinkData pres
    mBusy = True
    Set xl = PptTqq_GetEditExcel(core)
    If xl Is Nothing Then mBusy = False: PptTqq_Msg PPT_MSG_NO_EXCEL, vbExclamation: Exit Sub
    If Len(core) = 0 Then mBusy = False: PptTqq_Msg PPT_MSG_DATA_NOCORE, vbExclamation: Exit Sub
    If PptTqq_CoreApi(xl, core) < PPT_API_DATA Then mBusy = False: PptTqq_Msg PPT_MSG_OLD_CORE_DATA: Exit Sub
    PptTqq_EditShape shapes(1), xl, core
    mBusy = False
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    mBusy = False
    PptTqq_Msg PptTqq_DataErrorText("Não foi possível abrir os dados do gráfico: ", en, ed), vbExclamation
End Sub

' Apresentacao movida/renomeada (spec 14.2): formas TQQ_DATA cuja pasta de dados sumiu sao reapontadas para a pasta de dados NA PASTA DA
' APRESENTACAO (nome atual da apresentacao, ou o nome antigo do arquivo), se existir; o indice das automaticas acompanha. Sem Excel, sem
' perguntar. Devolve quantas formas. Barato quando TQQ_DATABOOK existe. Nunca propaga.
' Sonda PG4: SourceFullName e gravavel (persiste, mantem forma e animacao) e o PowerPoint NAO resolve caminho relativo. LinkFormat.Update com o
' arquivo ausente abre uma caixa modal que trava a chamada: aqui o candidato e conferido (FileExists) ANTES de gravar e nunca ha Update.
Public Function PptTqq_RelinkData(pres As Object) As Long
    Dim fso As Object, mark As String, sl As Object, shp As Object, oldP As String, cand As String, dirP As String, idx As String
    On Error Resume Next
    mark = PptTqq_Tag(pres, PPT_TAG_DATABOOK)
    If Len(mark) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function
    If CBool(fso.FileExists(mark)) Then Exit Function
    dirP = CStr(pres.Path)
    If Len(dirP) = 0 Then Exit Function
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If PptTqq_Tag(shp, PPT_TAG_DATA) = "1" Then
                oldP = PptTqq_SourcePath(PptTqq_LinkSource(shp), PptTqq_Tag(shp, PPT_TAG_BOOK))
                If Not CBool(fso.FileExists(oldP)) Then
                    cand = PptTqq_DataPathFor(CStr(pres.FullName))
                    If Not CBool(fso.FileExists(cand)) Then cand = dirP & "\" & fso.GetFileName(oldP)
                    If CBool(fso.FileExists(cand)) Then
                        Err.Clear
                        shp.LinkFormat.SourceFullName = cand
                        If Err.Number = 0 Then
                            shp.Tags.Add PPT_TAG_BOOK, cand
                            pres.Tags.Add PPT_TAG_DATABOOK, cand
                            If PptTqq_Tag(shp, PPT_TAG_AUTO) = "1" Then
                                idx = PptTqq_IdxSet(PptTqq_Tag(pres, PPT_TAG_INDEX), CLng(sl.SlideID), CStr(shp.Name), cand, True)
                                pres.Tags.Add PPT_TAG_INDEX, idx
                            End If
                            PptTqq_RelinkData = PptTqq_RelinkData + 1
                        Else
                            PptTqq_Log "PptTqq_RelinkData: " & Err.Number & " " & Err.Description
                        End If
                        Err.Clear
                    End If
                End If
            End If
        Next
    Next
    Err.Clear
End Function

' Evento PresentationSave: salva a pasta de dados aberta e alterada no Excel com o ThinQQ (o de edicao, senao o aberto); nunca abre nem cria
' um Excel; nunca mostra mensagem. Sonda PG5: o evento chega DEPOIS da gravacao do .pptx (arquivo ja novo, Saved = True; o BeforeSave vem
' antes) e, no Salvar como, ja com o nome novo; para a pasta de dados separada a ordem nao importa. Nunca propaga.
Public Sub PptTqq_OnSave(pres As Object)
    Dim xl As Object
    On Error Resume Next
    If Len(PptTqq_Tag(pres, PPT_TAG_DATABOOK)) = 0 Then Exit Sub
    If BUILD_CONFIG = "Test" Then
        Set xl = mTestXl
    Else
        If Not mEditXl Is Nothing Then
            If PptTqq_PingState(mEditXl) = 1 Then Set xl = mEditXl Else Set mEditXl = Nothing      ' fechado pelo usuario: o aberto
        End If
        If xl Is Nothing Then Set xl = GetObject(, "Excel.Application")
    End If
    Err.Clear
    If xl Is Nothing Then Exit Sub
    If PptTqq_PingState(xl) <> 1 Then Exit Sub
    If PptTqq_CoreApi(xl, PptTqq_CoreFileName()) < PPT_API_DATA Then Exit Sub
    xl.Run "'" & PptTqq_CoreFileName() & "'!Core_PptDataSave", pres
    Err.Clear
End Sub
