Attribute VB_Name = "modUpdater"
Option Explicit
' Verifica o manifesto (HTTP ou pasta), avisa o usuário e instala core/loader.
' Nenhuma mensagem sai daqui por MsgBox: tudo passa por Loader_Msg (em build de Teste apenas registra).
Private mBusy As Boolean                 ' protecao contra reentrada (o Transport usa DoEvents)
Public Updater_TestFormResult As String  ' gancho de teste: "update" | "skip" | "later" ("" = later no build de Teste)
Public Updater_FormShownCount As Long    ' quantas vezes o formulario foi solicitado (usado nos testes)
Public Updater_PendingScheduleCount As Long  ' vezes que a troca do loader seria agendada (build de Teste nao agenda OnTime)
Private Const NO_SHA_MSG As String = "O manifesto não informa o hash de verificação (SHA-256) do arquivo; por segurança nada foi instalado."
Private Const FAILURE_RETRY_HOURS As Double = 1
Private Const MIN_CORE_FOR_PENDING As String = "1.0.0"   ' primeira versao do core que exporta Core_ApplyPendingLoader
' Suplemento do PowerPoint (ThinQQ.ppam; spec 2026-10-14 secao 4.9)
Public Updater_PpamRetryScheduleCount As Long          ' gancho de teste: novas tentativas que seriam agendadas (Test nao agenda OnTime)
Public Updater_PpamReadyAtCoreSwap As Long            ' marcador de teste (so o build Test grava): 1 = o .ppam conferido ja existia quando a troca do core comecou
Private mTestPptRunning As Variant                    ' gancho de teste: Empty = consulta real
Private mTestPpamAuto As Variant                       ' gancho de teste: Empty = configuracao real
Private mLastPpamNote As String                        ' ultima frase sobre o suplemento (acrescentada ao aviso do atualizador)
Private mPpamRetryAt As Date                           ' horario da nova tentativa agendada (0 = nenhuma): nunca empilha OnTime
Private Const PPAM_RETRY_MINUTES As Long = 10

Public Function Updater_ShouldCheckNow(lastCheckText As String, nowValue As Date, intervalHours As Double, Optional lastFailureText As String = "") As Boolean
    Dim diff As Double
    ' Falha recente (< 1 h): nao insiste a cada abertura do Excel.
    If Len(lastFailureText) > 0 Then
        If IsDate(lastFailureText) Then
            diff = (nowValue - CDate(lastFailureText)) * 24
            If diff >= 0 And diff < FAILURE_RETRY_HOURS Then Updater_ShouldCheckNow = False: Exit Function
        End If
    End If
    If Len(lastCheckText) = 0 Then Updater_ShouldCheckNow = True: Exit Function
    If Not IsDate(lastCheckText) Then Updater_ShouldCheckNow = True: Exit Function
    diff = (nowValue - CDate(lastCheckText)) * 24
    If diff < 0 Then Updater_ShouldCheckNow = True: Exit Function    ' relogio voltou: verifica
    Updater_ShouldCheckNow = (diff >= intervalHours)
End Function

Public Function Updater_Evaluate(json As String, installedCore As String, installedLoader As String) As String
    Dim needCore As Boolean, needLoader As Boolean
    If JsonGetNumber(json, "schemaVersion") > 1 Then
        Loader_Log "Updater_Evaluate: schemaVersion " & JsonGetNumber(json, "schemaVersion") & " nao suportado; manifesto ignorado."
        Updater_Evaluate = "none"
        Exit Function
    End If
    needCore = CompareVersions(JsonGetString(json, "coreVersion"), installedCore) > 0
    needLoader = CompareVersions(JsonGetString(json, "loaderVersion"), installedLoader) > 0
    If needCore And needLoader Then
        Updater_Evaluate = "both"
    ElseIf needCore Then
        Updater_Evaluate = "core"
    ElseIf needLoader Then
        Updater_Evaluate = "loader"
    Else
        Updater_Evaluate = "none"
    End If
End Function

' Mensagem amigavel em pt-BR para um erro de Transport/Updater. Erros de rede crus (COM) viram texto generico.
Public Function Updater_FriendlyError(errNumber As Long, errDescription As String) As String
    Select Case errNumber
        Case vbObjectError + 1001: Updater_FriendlyError = "O servidor de atualização respondeu com erro. " & errDescription
        Case vbObjectError + 1002: Updater_FriendlyError = "Arquivo de atualização não encontrado. " & errDescription
        Case vbObjectError + 1003: Updater_FriendlyError = "Não foi possível verificar a integridade do arquivo baixado."
        Case vbObjectError + 1004: Updater_FriendlyError = "A fonte de atualização demorou demais para responder."
        Case vbObjectError + 1005: Updater_FriendlyError = "Sem conexão com a fonte de atualização. Verifique a rede e o proxy. " & errDescription
        Case vbObjectError + 2000 To vbObjectError + 2009: Updater_FriendlyError = errDescription
        Case 1 To 65535: Updater_FriendlyError = "Erro ao gravar os arquivos da atualização: " & errDescription
        Case Else: Updater_FriendlyError = "Sem conexão com a fonte de atualização. Verifique a rede e o endereço configurado."
    End Select
End Function

Public Function Updater_IsBusy() As Boolean
    Updater_IsBusy = mBusy
End Function

Public Sub Updater_TestSetBusy(value As Boolean)
    If BUILD_CONFIG <> "Test" Then Exit Sub     ' gancho de teste: inerte em Release
    mBusy = value
End Sub

' Chamado por Application.OnTime: qualquer erro aqui abriria um dialogo modal, entao tudo e contido.
Public Sub Updater_StartupCheck()
    On Error GoTo Fail
    If Not Updater_ShouldCheckNow(Settings_Get(KEY_LAST_CHECK), Now, Val(Settings_Get(KEY_INTERVAL, "24")), Settings_Get(KEY_LAST_FAILURE)) Then Exit Sub
    Updater_Check True
    Exit Sub
Fail:
    Loader_Log "Updater_StartupCheck: " & Err.Number & " " & Err.Description
End Sub

Public Sub Updater_CheckInteractive()
    On Error GoTo Fail
    Updater_Check False
    Exit Sub
Fail:
    Loader_Log "Updater_CheckInteractive: " & Err.Number & " " & Err.Description
End Sub

' Exibe o formulario de atualizacao. Em build de Teste nunca abre o modal: devolve o gancho (ou "later").
Private Function ShowUpdateForm(title As String, versionLine As String, notes As String, showSkip As Boolean) As String
    Dim frm As frmUpdate, en As Long, ed As String
    Updater_FormShownCount = Updater_FormShownCount + 1
    If BUILD_CONFIG = "Test" Then
        If Len(Updater_TestFormResult) > 0 Then ShowUpdateForm = Updater_TestFormResult Else ShowUpdateForm = "later"
        Exit Function
    End If
    On Error GoTo Fail
    Set frm = New frmUpdate
    frm.Prepare title, versionLine, notes, showSkip, showPpam:=Updater_PpamInstalled()
    frm.Show
    ShowUpdateForm = frm.Result
    Unload frm
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Unload frm
    On Error GoTo 0
    Err.Raise en, "Updater", ed
End Function

' Linha de versao do formulario (diz explicitamente quando so o carregador muda).
Public Function Updater_VersionLine(what As String, installedCore As String, json As String) As String
    Dim pub As String
    pub = "  (" & JsonGetString(json, "publishedAt") & ")"
    Select Case what
        Case "loader"
            Updater_VersionLine = "Atualização do carregador (loader) - Instalado: " & APP_VERSION & "   -   Novo: " & JsonGetString(json, "loaderVersion") & pub
        Case "both"
            Updater_VersionLine = "Core: " & installedCore & " > " & JsonGetString(json, "coreVersion") & "   |   Carregador: " & APP_VERSION & " > " & JsonGetString(json, "loaderVersion") & pub
        Case Else
            Updater_VersionLine = "Instalada: " & installedCore & "   -   Nova: " & JsonGetString(json, "coreVersion") & pub
    End Select
    If what <> "loader" Then
        If Updater_PpamNeeded(json) Then Updater_VersionLine = Updater_VersionLine & "   |   Suplemento do PowerPoint: " & Settings_Get(KEY_PPAM_VERSION, "0.0.0") & " > " & JsonGetString(json, "pptAddinVersion")
    End If
End Function

' Chave de "ignorar esta versao": combina core e loader do manifesto.
Public Function Updater_SkipKey(json As String) As String
    Updater_SkipKey = JsonGetString(json, "coreVersion") & "|" & JsonGetString(json, "loaderVersion")
End Function

' Agenda (OnTime, fora da pilha do loader) a troca do loader pendente pelo core. Nao agenda se o core
' estiver ausente ou for antigo demais (um OnTime para macro inexistente abriria dialogo modal).
Public Sub Updater_SchedulePendingLoader(delaySeconds As Long)
    On Error GoTo Fail
    Dim p As String
    p = Settings_Get(KEY_PENDING_LOADER)
    If Len(p) = 0 Then Exit Sub
    If Len(Dir(p)) = 0 Then Exit Sub
    If Loader_SwapInProgress() Then Exit Sub        ' o core ja esta aplicando esta troca (novo loader acabou de subir)
    If Loader_CoreWorkbook() Is Nothing Then Loader_Log "Loader pendente aguardando core: " & p: Exit Sub
    If CompareVersions(Loader_CurrentCoreVersion(), MIN_CORE_FOR_PENDING) < 0 Then
        Loader_Log "Core sem Core_ApplyPendingLoader; loader pendente mantido: " & p
        Exit Sub
    End If
    If BUILD_CONFIG = "Test" Then Updater_PendingScheduleCount = Updater_PendingScheduleCount + 1: Exit Sub
    Application.OnTime Now + TimeSerial(0, 0, delaySeconds), "'" & Loader_CoreFileName() & "'!Core_ApplyPendingLoader"
    Exit Sub
Fail:
    Loader_Log "Updater_SchedulePendingLoader: " & Err.Number & " " & Err.Description
End Sub

' ---------- Suplemento do PowerPoint (ThinQQ.ppam; spec 2026-10-14 secao 4.9) ----------
' O atualizador so baixa, confere o SHA-256 e renomeia o arquivo: nunca abre nem executa o .ppam.

Public Sub Updater_TestSetPptRunning(v As Variant)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestPptRunning = v
End Sub

Public Sub Updater_TestSetPpamAuto(v As Variant)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestPpamAuto = v
End Sub

Public Function Updater_LastPpamNote() As String
    Updater_LastPpamNote = mLastPpamNote
End Function

Private Function PpamPath() As String
    PpamPath = Loader_Root() & "\ThinQQ.ppam"
End Function

' Build de Teste sem raiz de teste: a raiz e a instalacao REAL do usuario; os testes nunca a tratam como instalada nem mexem nela.
Private Function PpamRootIsRealInTest() As Boolean
    If BUILD_CONFIG <> "Test" Then Exit Function
    PpamRootIsRealInTest = (StrComp(Loader_Root(), Environ$("LOCALAPPDATA") & "\ThinQQ", vbTextCompare) = 0)
End Function

' PowerPoint em execucao (WMI; nunca GetObject no PowerPoint). Na duvida (WMI indisponivel): em execucao. So decide a frase do aviso.
Private Function Updater_PptRunning() As Boolean
    If BUILD_CONFIG = "Test" And Not IsEmpty(mTestPptRunning) Then Updater_PptRunning = CBool(mTestPptRunning): Exit Function
    On Error Resume Next
    Updater_PptRunning = True
    Updater_PptRunning = (GetObject("winmgmts:" & Chr(92) & Chr(92) & "." & Chr(92) & "root" & Chr(92) & "cimv2").ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'POWERPNT.EXE'").Count > 0)
End Function

Private Function Updater_PpamAutoOn() As Boolean
    If BUILD_CONFIG = "Test" And Not IsEmpty(mTestPpamAuto) Then Updater_PpamAutoOn = (CStr(mTestPpamAuto) <> "0"): Exit Function
    Updater_PpamAutoOn = (Settings_Get(KEY_PPAM_AUTO, "1") <> "0")
End Function

' (5) Instalado = o instalador deixou <raiz>\ThinQQ.ppam. O atualizador NUNCA instala o suplemento para quem nao o tem.
Public Function Updater_PpamInstalled() As Boolean
    On Error Resume Next
    If PpamRootIsRealInTest() Then Exit Function
    Updater_PpamInstalled = (Len(Dir(PpamPath())) > 0)
    If Not Updater_PpamInstalled Then Updater_PpamInstalled = (Len(Dir(PpamPath() & ".old")) > 0)   ' troca interrompida: o .old e a copia instalada
End Function

' Troca interrompida (ThinQQ.ppam ausente, .old presente): devolve o .old ao lugar. True = ThinQQ.ppam existe ao final.
Private Function PpamRecoverOld() As Boolean
    Dim dest As String
    On Error Resume Next
    dest = PpamPath()
    If Len(Dir(dest)) = 0 And Len(Dir(dest & ".old")) > 0 Then
        Name dest & ".old" As dest
        If Err.Number <> 0 Then Loader_Log "PpamRecoverOld: " & Err.Number & " " & Err.Description Else Loader_Log "Suplemento do PowerPoint: .old de uma troca interrompida devolvido"
        Err.Clear
    End If
    PpamRecoverOld = (Len(Dir(dest)) > 0)
End Function

' 64 caracteres hexadecimais (maiusculas ou minusculas).
Private Function IsSha256Hex(s As String) As Boolean
    Dim i As Long
    If Len(s) <> 64 Then Exit Function
    For i = 1 To 64
        If Not (Mid$(s, i, 1) Like "[0-9a-fA-F]") Then Exit Function
    Next
    IsSha256Hex = True
End Function

' (1) Ha suplemento mais novo para este usuario: opcao ligada, instalado, manifesto com versao/url/SHA validos, versao maior que a
' instalada (PpamVersion; ausente = 0.0.0) e diferente da que ja esta pendente.
Public Function Updater_PpamNeeded(json As String) As Boolean
    Dim v As String
    On Error GoTo Fail
    If JsonGetNumber(json, "schemaVersion") > 1 Then
        Loader_Log "Updater_PpamNeeded: schemaVersion " & JsonGetNumber(json, "schemaVersion") & " nao suportado; suplemento ignorado."
        Exit Function
    End If
    If Not Updater_PpamAutoOn() Then Exit Function
    If Not Updater_PpamInstalled() Then Exit Function
    v = JsonGetString(json, "pptAddinVersion")
    If Not IsValidVersion(v) Then Exit Function
    If Len(JsonGetString(json, "pptAddinUrl")) = 0 Then Exit Function
    If Not IsSha256Hex(Trim$(JsonGetString(json, "pptAddinSha256"))) Then
        Loader_Log "Updater_PpamNeeded: manifesto com pptAddinUrl mas pptAddinSha256 ausente ou malformado; suplemento ignorado."
        Exit Function
    End If
    If Settings_Get(KEY_PENDING_PPAM) = "1" And Settings_Get(KEY_PENDING_PPAM_VER) = v Then Exit Function
    Updater_PpamNeeded = (CompareVersions(v, Settings_Get(KEY_PPAM_VERSION, "0.0.0")) > 0)
    Exit Function
Fail:
    Loader_Log "Updater_PpamNeeded: " & Err.Number & " " & Err.Description
    Updater_PpamNeeded = False
End Function

' So o suplemento, sem o core: precisa (Updater_PpamNeeded) e o core INSTALADO ja o acompanha (pptAddinVersion <= versao do core em uso; o
' suplemento sai junto com o core da mesma versao, cuja API ele exige). Nunca "suplemento novo + core velho".
Private Function Updater_PpamOnlyNeeded(json As String, installedCore As String) As Boolean
    On Error GoTo Fail
    If Not Updater_PpamNeeded(json) Then Exit Function
    If CompareVersions(JsonGetString(json, "pptAddinVersion"), installedCore) > 0 Then
        Loader_Log "Suplemento do PowerPoint " & JsonGetString(json, "pptAddinVersion") & " aguarda o core (instalado: " & installedCore & ")"
        Exit Function
    End If
    Updater_PpamOnlyNeeded = True
    Exit Function
Fail:
    Loader_Log "Updater_PpamOnlyNeeded: " & Err.Number & " " & Err.Description
    Updater_PpamOnlyNeeded = False
End Function

' (7) Baixa para tmp e confere o SHA-256 do manifesto (obrigatorio: AllowUnverified nao vale para o suplemento). False = nada fica em tmp.
Private Function Updater_PpamDownload(json As String, tmp As String) As Boolean
    On Error GoTo Fail
    If Len(Dir(tmp)) > 0 Then Kill tmp
    Transport_DownloadFile JsonGetString(json, "pptAddinUrl"), tmp
    If LCase$(Transport_Sha256File(tmp)) <> LCase$(Trim$(JsonGetString(json, "pptAddinSha256"))) Then
        Kill tmp
        Loader_Log "Updater_PpamDownload: SHA-256 nao confere; suplemento nao atualizado"
        Exit Function
    End If
    Updater_PpamDownload = True
    Exit Function
Fail:
    Loader_Log "Updater_PpamDownload: " & Err.Number & " " & Err.Description
    On Error GoTo -1
    On Error Resume Next
    If Len(Dir(tmp)) > 0 Then Kill tmp
End Function

Private Sub ClearPendingPpam()
    Settings_Delete KEY_PENDING_PPAM: Settings_Delete KEY_PENDING_PPAM_SHA: Settings_Delete KEY_PENDING_PPAM_VER
    Updater_CancelPpamRetry
End Sub

' (2) Troca atomica (mesma pasta: renomeacoes): ThinQQ.ppam -> .old, src -> ThinQQ.ppam, apaga .old. Arquivo em uso (a renomeacao
' falha: medido com o PowerPoint) -> False sem mudar nada. Falha no 2o passo -> o .old volta.
Private Function Updater_PpamSwap(src As String) As Boolean
    Dim dest As String, old As String
    dest = PpamPath(): old = dest & ".old"
    If Not PpamRecoverOld() Then Exit Function          ' nunca apaga a ultima copia (.old sem ThinQQ.ppam)
    On Error Resume Next
    If Len(Dir(old)) > 0 Then Kill old
    Err.Clear
    Name dest As old
    If Err.Number <> 0 Then Err.Clear: Exit Function            ' em uso
    Name src As dest
    If Err.Number <> 0 Then
        Err.Clear
        Name old As dest
        Exit Function
    End If
    Kill old
    Err.Clear
    Updater_PpamSwap = True
End Function

' (2) Depois do download conferido: o baixado vai para ThinQQ.ppam.new (mesma pasta do suplemento, para a troca ser uma renomeacao
' atomica) e troca agora, ou fica pendente (.new + chaves) com nova tentativa agendada. Devolve "replaced", "replaced-restart" (trocado com
' o PowerPoint aberto: vale ao reabrir) ou "staged". Erro = propaga (quem chama trata; o temporario e apagado).
Private Function Updater_PpamCommit(tmp As String, sha As String, ver As String) As String
    Dim nw As String, en As Long, ed As String
    On Error GoTo Fail
    nw = PpamPath() & ".new"
    ClearPendingPpam                                     ' o .new vai ser substituido: uma pendencia antiga deixa de valer
    If Len(Dir(nw)) > 0 Then Kill nw
    Name tmp As nw
    If Updater_PpamSwap(nw) Then
        Settings_Set KEY_PPAM_VERSION, ver
        ClearPendingPpam
        If Updater_PptRunning() Then Updater_PpamCommit = "replaced-restart" Else Updater_PpamCommit = "replaced"
        Loader_Log "Suplemento do PowerPoint atualizado para " & ver
        Exit Function
    End If
    Settings_Set KEY_PENDING_PPAM_SHA, LCase$(Trim$(sha)): Settings_Set KEY_PENDING_PPAM_VER, ver: Settings_Set KEY_PENDING_PPAM, "1"
    Updater_SchedulePpamRetry
    Loader_Log "Suplemento do PowerPoint " & ver & " pendente (arquivo em uso)"
    Updater_PpamCommit = "staged"
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    If Len(Dir(tmp)) > 0 Then Kill tmp
    If Settings_Get(KEY_PENDING_PPAM) <> "1" Then
        If Len(Dir(nw)) > 0 Then Kill nw
    End If
    On Error GoTo 0
    Err.Raise en, "Updater", ed
End Function

' So o suplemento (core em dia): baixa, confere e aplica. "" = nada feito (nao precisava ou falhou: registrado). Nunca propaga.
Public Function Updater_PpamUpdateNow(json As String) As String
    Dim tmp As String
    On Error GoTo Fail
    If Not Updater_PpamNeeded(json) Then Exit Function
    Loader_EnsureFolders
    tmp = Loader_TmpDir() & "\ThinQQ.ppam"
    If Not Updater_PpamDownload(json, tmp) Then Exit Function
    Updater_PpamUpdateNow = Updater_PpamCommit(tmp, JsonGetString(json, "pptAddinSha256"), JsonGetString(json, "pptAddinVersion"))
    Exit Function
Fail:
    Loader_Log "Updater_PpamUpdateNow: " & Err.Number & " " & Err.Description
    Updater_PpamUpdateNow = ""
End Function

' (4) Frase para o usuario (acrescentada ao aviso de atualizacao).
Public Function Updater_PpamStatusText(state As String) As String
    Select Case state
        Case "replaced": Updater_PpamStatusText = "Suplemento do PowerPoint atualizado."
        Case "replaced-restart": Updater_PpamStatusText = "Suplemento do PowerPoint atualizado: feche e abra o PowerPoint para usar a versão nova."
        Case "staged": Updater_PpamStatusText = "O suplemento do PowerPoint será atualizado automaticamente quando o PowerPoint for fechado (o Excel tenta de novo a cada 10 minutos e na próxima abertura)."
        Case "failed": Updater_PpamStatusText = "O suplemento do PowerPoint não foi atualizado agora; nova tentativa na próxima verificação."
    End Select
End Function

' (2) Aplicacao da pendencia: na abertura do Excel (Loader_Startup) e pela nova tentativa agendada. Reconfere o SHA-256 do .new (arquivo em
' pasta do usuario: nunca confia no que esta la); adulterado, ausente ou nao mais novo que o instalado = descarta. Ainda em uso = continua
' pendente (e agenda nova tentativa, sem empilhar). Antes, com ou sem pendencia: uma troca interrompida (ThinQQ.ppam ausente, .old
' presente) tem o .old devolvido ao lugar. Nunca propaga.
Public Sub Updater_ApplyPendingPpam()
    Dim nw As String, ver As String, ok As Boolean
    On Error GoTo Fail
    If PpamRootIsRealInTest() Then Exit Sub
    PpamRecoverOld
    If Settings_Get(KEY_PENDING_PPAM) <> "1" Then Exit Sub
    nw = PpamPath() & ".new"
    ver = Settings_Get(KEY_PENDING_PPAM_VER)
    If Not Updater_PpamInstalled() Then                  ' desinstalado depois de pendente: descarta (nunca instala, nunca tenta para sempre)
        Loader_Log "Updater_ApplyPendingPpam: suplemento nao instalado; pendencia descartada"
        If Len(Dir(nw)) > 0 Then Kill nw
        ClearPendingPpam
        Exit Sub
    End If
    ok = (Len(Dir(nw)) > 0)
    If ok Then ok = IsValidVersion(ver)
    If ok Then ok = (CompareVersions(ver, Settings_Get(KEY_PPAM_VERSION, "0.0.0")) > 0)
    If ok Then ok = (LCase$(Transport_Sha256File(nw)) = LCase$(Trim$(Settings_Get(KEY_PENDING_PPAM_SHA))))
    If Not ok Then
        Loader_Log "Updater_ApplyPendingPpam: .new ausente, com SHA-256 diferente ou nao mais novo que o instalado; descartado"
        If Len(Dir(nw)) > 0 Then Kill nw
        ClearPendingPpam
        Exit Sub
    End If
    If Updater_PpamSwap(nw) Then
        Settings_Set KEY_PPAM_VERSION, ver
        ClearPendingPpam
        Loader_Log "Suplemento do PowerPoint pendente aplicado: " & ver
    Else
        Updater_SchedulePpamRetry
    End If
    Exit Sub
Fail:
    Loader_Log "Updater_ApplyPendingPpam: " & Err.Number & " " & Err.Description
End Sub

' Alvo do Application.OnTime: a tentativa agendada disparou (nao esta mais agendada). Durante uma atualizacao em andamento, so reagenda.
Public Sub Updater_PpamRetryTick()
    On Error Resume Next
    mPpamRetryAt = 0
    If mBusy Then Updater_SchedulePpamRetry: Exit Sub
    Updater_ApplyPendingPpam
End Sub

' Nova tentativa em 10 min (so enquanto houver pendencia; o Excel fechado leva a tentativa para a proxima abertura). Uma so por vez.
Private Sub Updater_SchedulePpamRetry()
    On Error Resume Next
    If Settings_Get(KEY_PENDING_PPAM) <> "1" Then Exit Sub
    If mPpamRetryAt <> 0 Then Exit Sub                   ' ja agendada: nunca empilha
    mPpamRetryAt = Now + TimeSerial(0, PPAM_RETRY_MINUTES, 0)
    If BUILD_CONFIG = "Test" Then Updater_PpamRetryScheduleCount = Updater_PpamRetryScheduleCount + 1: Exit Sub
    Err.Clear
    Application.OnTime mPpamRetryAt, "'" & ThisWorkbook.Name & "'!Updater_PpamRetryTick"
    If Err.Number <> 0 Then
        Loader_Log "Updater_SchedulePpamRetry: " & Err.Number & " " & Err.Description
        mPpamRetryAt = 0
    End If
End Sub

' Cancela a nova tentativa agendada (Loader_Shutdown: um OnTime para o loader fechado o reabriria). Nunca propaga.
Public Sub Updater_CancelPpamRetry()
    On Error Resume Next
    If mPpamRetryAt = 0 Then Exit Sub
    If BUILD_CONFIG <> "Test" Then Application.OnTime mPpamRetryAt, "'" & ThisWorkbook.Name & "'!Updater_PpamRetryTick", , False
    mPpamRetryAt = 0
End Sub

Public Sub Updater_Check(silent As Boolean)
    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Fail
    Dim json As String, what As String, coreVer As String, installedCore As String, choice As String, ppamOnly As Boolean
    json = Transport_FetchText(Settings_UpdateSource())
    Settings_Set KEY_LAST_CHECK, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Settings_Delete KEY_LAST_FAILURE
    installedCore = Loader_CurrentCoreVersion()
    what = Updater_Evaluate(json, installedCore, APP_VERSION)
    coreVer = JsonGetString(json, "coreVersion")
    ' So o suplemento (core em dia, ou so o carregador muda, mesmo ignorado/adiado): aplica sem formulario e registra. Nunca passa a frente do
    ' core instalado; com core novo pendente de decisao, o suplemento so vai junto com ele (Updater_Apply).
    If what = "none" Or what = "loader" Then ppamOnly = Updater_PpamOnlyNeeded(json, installedCore)
    If ppamOnly Then
        mLastPpamNote = Updater_PpamUpdateNow(json)
        If Len(mLastPpamNote) = 0 Then mLastPpamNote = "failed"
        mLastPpamNote = Updater_PpamStatusText(mLastPpamNote)
        Loader_Log mLastPpamNote
        If Not silent Then Loader_Msg mLastPpamNote, vbInformation
    End If
    If what = "none" Then
        If Not silent And Not ppamOnly Then Loader_Msg "O ThinQQ já está na versão mais recente (" & installedCore & ").", vbInformation
        mBusy = False
        Exit Sub
    End If
    If silent And Settings_Get(KEY_SKIPPED) = Updater_SkipKey(json) Then mBusy = False: Exit Sub
    If what = "loader" Then
        choice = ShowUpdateForm("Nova versão do carregador do ThinQQ disponível", Updater_VersionLine(what, installedCore, json), JsonGetString(json, "notes"), silent)
    Else
        choice = ShowUpdateForm("Nova versão do ThinQQ disponível", Updater_VersionLine(what, installedCore, json), JsonGetString(json, "notes"), silent)
    End If
    Select Case choice
        Case "update"
            ' o flag continua ligado durante a instalacao (DoEvents no Transport)
            Updater_Apply json, what
        Case "skip": Settings_Set KEY_SKIPPED, Updater_SkipKey(json)
    End Select
    mBusy = False
    Exit Sub
Fail:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    mBusy = False
    On Error Resume Next
    Settings_Set KEY_LAST_FAILURE, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Loader_Log "Updater_Check(silent=" & silent & "): " & en & " " & ed
    If Not silent Then
        Loader_Msg "Não foi possível verificar atualizações." & vbLf & Updater_FriendlyError(en, ed) & vbLf & vbLf & _
                   "Fonte: " & Settings_UpdateSource(), vbExclamation
    End If
End Sub

' Baixa e instala. what = core | loader | both. Retorna True se tudo deu certo.
' Ordem do core: baixa para tmp -> confere SHA -> grava <core>.new na pasta da versao nova -> renomeia -> abre
' (Core_Ping). Se nao abrir: apaga SO a pasta da versao nova e reabre o core que estava carregado.
' Manifesto sem SHA e recusado, salvo AllowUnverified=1 no registro (decisao explicita do usuario).
Public Function Updater_Apply(json As String, what As String, Optional quiet As Boolean = False) As Boolean
    On Error GoTo Fail
    Dim coreVer As String, tmp As String, dest As String, destDir As String, stage As String, sha As String, actual As String
    Dim url As String, prevPath As String, wb As Workbook, allow As Boolean, swapStarted As Boolean, committed As Boolean
    Dim doCore As Boolean, doLoader As Boolean, ppamTmp As String, ppamOk As Boolean, ppamState As String, wantPpam As Boolean
    mLastPpamNote = ""
    doCore = (what = "core" Or what = "both")
    doLoader = (what = "loader" Or what = "both")
    allow = (Settings_Get(KEY_ALLOW_UNVERIFIED) = "1")
    ' Validacoes antes de qualquer download.
    If doCore Then
        coreVer = JsonGetString(json, "coreVersion")
        If Not IsValidVersion(coreVer) Then Err.Raise vbObjectError + 2000, "Updater", "Manifesto sem coreVersion válida."
        If Len(Trim$(JsonGetString(json, "coreSha256"))) = 0 And Not allow Then Err.Raise vbObjectError + 2004, "Updater", NO_SHA_MSG
        ' Nunca substitui o arquivo do core em uso: mesma versao ja carregada => recusa sem tocar em nada.
        Set wb = Loader_CoreWorkbook()
        If Not wb Is Nothing Then
            If StrComp(Left$(wb.FullName, InStrRev(wb.FullName, "\") - 1), Loader_VersionsDir() & "\" & coreVer, vbTextCompare) = 0 Then
                Err.Raise vbObjectError + 2005, "Updater", "Essa versão (" & coreVer & ") já está instalada e em uso."
            End If
        End If
    End If
    If doLoader Then
        If Len(Trim$(JsonGetString(json, "loaderSha256"))) = 0 And Not allow Then Err.Raise vbObjectError + 2004, "Updater", NO_SHA_MSG
    End If
    Loader_EnsureFolders
    Application.StatusBar = "ThinQQ: baixando atualização..."
    ' Suplemento do PowerPoint: baixado e conferido ANTES da troca do core; so e trocado depois do core trocado e carregado (o par nunca
    ' fica "suplemento novo + core velho"). Falha dele nunca desfaz o core: o aviso diz que fica para a proxima verificacao.
    wantPpam = doCore And Updater_PpamNeeded(json)
    If wantPpam Then ppamTmp = Loader_TmpDir() & "\ThinQQ.ppam": ppamOk = Updater_PpamDownload(json, ppamTmp)
    If doCore Then
        url = JsonGetString(json, "coreUrl")
        tmp = Loader_TmpDir() & "\" & Loader_CoreFileName()
        If Len(Dir(tmp)) > 0 Then Kill tmp
        Transport_DownloadFile url, tmp
        sha = LCase$(Trim$(JsonGetString(json, "coreSha256")))
        If Len(sha) > 0 Then
            If LCase$(Transport_Sha256File(tmp)) <> sha Then
                Kill tmp
                Err.Raise vbObjectError + 2001, "Updater", "Arquivo baixado corrompido (SHA-256 não confere)."
            End If
        Else
            Loader_Log "Updater_Apply: AllowUnverified=1; instalando core sem verificação de integridade."
        End If
        destDir = Loader_VersionsDir() & "\" & coreVer
        dest = destDir & "\" & Loader_CoreFileName()
        stage = dest & ".new"
        Set wb = Loader_CoreWorkbook()
        If Not wb Is Nothing Then prevPath = wb.FullName
        If Len(Dir(destDir, vbDirectory)) = 0 Then MkDir destDir
        If BUILD_CONFIG = "Test" Then
            Updater_PpamReadyAtCoreSwap = 0
            If wantPpam And ppamOk Then
                If Len(Dir(ppamTmp)) > 0 Then Updater_PpamReadyAtCoreSwap = 1
            End If
        End If
        swapStarted = True
        If Len(Dir(stage)) > 0 Then Kill stage
        FileCopy tmp, stage
        Kill tmp
        If Len(Dir(dest)) > 0 Then
            Kill dest
        End If
        Name stage As dest
        url = JsonGetString(json, "tutorialUrl")
        If Len(url) > 0 Then
            On Error Resume Next
            Transport_DownloadFile url, destDir & "\tutorial.html"
            If Err.Number <> 0 Then Loader_Log "Updater_Apply: tutorial não baixado (ignorado): " & Err.Description
            Err.Clear
            On Error GoTo Fail
        End If
        If Not Loader_OpenCore(dest) Then Err.Raise vbObjectError + 2003, "Updater", "A nova versão foi baixada, mas não pôde ser carregada. A versão anterior foi mantida."
        committed = True
        Settings_Set KEY_LAST_NOTIFIED_CORE, coreVer
        Settings_Delete KEY_SKIPPED
        Loader_InvalidateRibbon
        If wantPpam Then
            ppamState = "failed"
            If ppamOk Then
                On Error Resume Next
                ppamState = Updater_PpamCommit(ppamTmp, JsonGetString(json, "pptAddinSha256"), JsonGetString(json, "pptAddinVersion"))
                If Err.Number <> 0 Then
                    Loader_Log "Updater_Apply: suplemento do PowerPoint nao trocado: " & Err.Number & " " & Err.Description
                    ppamState = "failed"
                End If
                On Error GoTo -1
                On Error GoTo Fail
            End If
            mLastPpamNote = Updater_PpamStatusText(ppamState)
            Loader_Log mLastPpamNote
        End If
    End If
    If doLoader Then
        url = JsonGetString(json, "loaderUrl")
        tmp = Loader_TmpDir() & "\ThinQQ.xlam"
        If Len(Dir(tmp)) > 0 Then Kill tmp
        Transport_DownloadFile url, tmp
        sha = LCase$(Trim$(JsonGetString(json, "loaderSha256")))
        actual = LCase$(Transport_Sha256File(tmp))
        If Len(sha) > 0 Then
            If actual <> sha Then
                Kill tmp
                Err.Raise vbObjectError + 2002, "Updater", "Arquivo baixado corrompido (SHA-256 não confere)."
            End If
        Else
            Loader_Log "Updater_Apply: AllowUnverified=1; instalando loader sem verificação de integridade."
        End If
        ' O core recalcula este SHA antes de trocar o loader (o caminho vem do registro).
        Settings_Set KEY_PENDING_LOADER_SHA, actual
        Settings_Set KEY_PENDING_LOADER, tmp
        Settings_Delete KEY_SKIPPED
        ' A troca do loader é feita pelo core, fora da pilha de chamadas do loader.
        Updater_SchedulePendingLoader 1
    End If
    Application.StatusBar = False
    If Not quiet Then Loader_Msg "ThinQQ atualizado com sucesso." & IIf(Len(coreVer) > 0, " Versão " & coreVer & ".", "") & IIf(Len(mLastPpamNote) > 0, vbLf & mLastPpamNote, ""), vbInformation
    Updater_Apply = True
    Exit Function
Fail:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Application.StatusBar = False
    If swapStarted And Not committed Then RollbackCore dest, destDir, stage, tmp, prevPath
    If Len(ppamTmp) > 0 Then
        If Len(Dir(ppamTmp)) > 0 Then Kill ppamTmp          ' core nao trocado: o suplemento baixado nunca entra sozinho
    End If
    Loader_Log "Updater_Apply: " & en & " " & ed
    If Not quiet Then Loader_Msg "Falha ao atualizar: " & Updater_FriendlyError(en, ed), vbExclamation
    Updater_Apply = False
End Function

' Desfaz uma troca de core que nao concluiu: apaga so a pasta da versao nova (nunca a que esta em uso)
' e recarrega o core que estava aberto antes.
Private Sub RollbackCore(dest As String, destDir As String, stage As String, tmp As String, prevPath As String)
    Dim fso As Object, cur As Workbook
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set cur = Loader_CoreWorkbook()
    If Not cur Is Nothing Then
        If StrComp(cur.FullName, dest, vbTextCompare) = 0 Then Loader_CloseCore
    End If
    If Len(tmp) > 0 Then
        If Len(Dir(tmp)) > 0 Then Kill tmp
    End If
    If Len(Dir(stage)) > 0 Then Kill stage
    If StrComp(fso.GetParentFolderName(prevPath), destDir, vbTextCompare) <> 0 Then
        If fso.FolderExists(destDir) Then fso.DeleteFolder destDir, True
    End If
    If Len(prevPath) > 0 Then
        Set cur = Loader_CoreWorkbook()
        If cur Is Nothing Then
            Loader_OpenCore prevPath
        ElseIf StrComp(cur.FullName, prevPath, vbTextCompare) <> 0 Then
            Loader_OpenCore prevPath
        End If
    End If
    Loader_Log "Updater_Apply: troca do core desfeita; core anterior: " & prevPath
End Sub
