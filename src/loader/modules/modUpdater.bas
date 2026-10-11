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
    frm.Prepare title, versionLine, notes, showSkip
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

Public Sub Updater_Check(silent As Boolean)
    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Fail
    Dim json As String, what As String, coreVer As String, installedCore As String, choice As String
    json = Transport_FetchText(Settings_UpdateSource())
    Settings_Set KEY_LAST_CHECK, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Settings_Delete KEY_LAST_FAILURE
    installedCore = Loader_CurrentCoreVersion()
    what = Updater_Evaluate(json, installedCore, APP_VERSION)
    coreVer = JsonGetString(json, "coreVersion")
    If what = "none" Then
        If Not silent Then Loader_Msg "O ThinQQ já está na versão mais recente (" & installedCore & ").", vbInformation
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
    Dim doCore As Boolean, doLoader As Boolean
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
    If Not quiet Then Loader_Msg "ThinQQ atualizado com sucesso." & IIf(Len(coreVer) > 0, " Versão " & coreVer & ".", ""), vbInformation
    Updater_Apply = True
    Exit Function
Fail:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Application.StatusBar = False
    If swapStarted And Not committed Then RollbackCore dest, destDir, stage, tmp, prevPath
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
