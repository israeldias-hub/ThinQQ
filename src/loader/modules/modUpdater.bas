Attribute VB_Name = "modUpdater"
Option Explicit
' Verifica o manifesto (HTTP ou pasta), avisa o usuário e instala core/loader.
' Nenhuma mensagem sai daqui por MsgBox: tudo passa por Loader_Msg (em build de Teste apenas registra).
Private mBusy As Boolean                 ' protecao contra reentrada (o Transport usa DoEvents)
Public Updater_TestFormResult As String  ' gancho de teste: "update" | "skip" | "later" ("" = later no build de Teste)
Public Updater_FormShownCount As Long    ' quantas vezes o formulario foi solicitado (usado nos testes)

Public Function Updater_ShouldCheckNow(lastCheckText As String, nowValue As Date, intervalHours As Double) As Boolean
    Dim diff As Double
    If Len(lastCheckText) = 0 Then Updater_ShouldCheckNow = True: Exit Function
    If Not IsDate(lastCheckText) Then Updater_ShouldCheckNow = True: Exit Function
    diff = (nowValue - CDate(lastCheckText)) * 24
    If diff < 0 Then Updater_ShouldCheckNow = True: Exit Function    ' relogio voltou: verifica
    Updater_ShouldCheckNow = (diff >= intervalHours)
End Function

Public Function Updater_Evaluate(json As String, installedCore As String, installedLoader As String) As String
    Dim needCore As Boolean, needLoader As Boolean
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
    mBusy = value
End Sub

' Chamado por Application.OnTime: qualquer erro aqui abriria um dialogo modal, entao tudo e contido.
Public Sub Updater_StartupCheck()
    On Error GoTo Fail
    If Not Updater_ShouldCheckNow(Settings_Get(KEY_LAST_CHECK), Now, Val(Settings_Get(KEY_INTERVAL, "24"))) Then Exit Sub
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
    Dim frm As frmUpdate
    Updater_FormShownCount = Updater_FormShownCount + 1
    If BUILD_CONFIG = "Test" Then
        If Len(Updater_TestFormResult) > 0 Then ShowUpdateForm = Updater_TestFormResult Else ShowUpdateForm = "later"
        Exit Function
    End If
    Set frm = New frmUpdate
    frm.Prepare title, versionLine, notes, showSkip
    frm.Show
    ShowUpdateForm = frm.Result
    Unload frm
End Function

Public Sub Updater_Check(silent As Boolean)
    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Fail
    Dim json As String, what As String, coreVer As String, installedCore As String, choice As String
    json = Transport_FetchText(Settings_UpdateSource())
    Settings_Set KEY_LAST_CHECK, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    installedCore = Loader_CurrentCoreVersion()
    what = Updater_Evaluate(json, installedCore, APP_VERSION)
    coreVer = JsonGetString(json, "coreVersion")
    If what = "none" Then
        If Not silent Then Loader_Msg "O ThinQQ já está na versão mais recente (" & installedCore & ").", vbInformation
        mBusy = False
        Exit Sub
    End If
    If silent And Settings_Get(KEY_SKIPPED) = coreVer Then mBusy = False: Exit Sub
    choice = ShowUpdateForm("Nova versão do ThinQQ disponível", _
                "Instalada: " & installedCore & "   -   Nova: " & coreVer & "  (" & JsonGetString(json, "publishedAt") & ")", _
                JsonGetString(json, "notes"), silent)
    Select Case choice
        Case "update"
            ' o flag continua ligado durante a instalacao (DoEvents no Transport)
            Updater_Apply json, what
        Case "skip": Settings_Set KEY_SKIPPED, coreVer
    End Select
    mBusy = False
    Exit Sub
Fail:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    mBusy = False
    On Error Resume Next
    Loader_Log "Updater_Check(silent=" & silent & "): " & en & " " & ed
    If Not silent Then
        Loader_Msg "Não foi possível verificar atualizações." & vbLf & Updater_FriendlyError(en, ed) & vbLf & vbLf & _
                   "Fonte: " & Settings_UpdateSource(), vbExclamation
    End If
End Sub

' Baixa e instala. what = core | loader | both. Retorna True se tudo deu certo.
Public Function Updater_Apply(json As String, what As String, Optional quiet As Boolean = False) As Boolean
    On Error GoTo Fail
    Dim coreVer As String, tmp As String, dest As String, destDir As String, sha As String, url As String
    Loader_EnsureFolders
    Application.StatusBar = "ThinQQ: baixando atualização..."
    If what = "core" Or what = "both" Then
        coreVer = JsonGetString(json, "coreVersion")
        If Not IsValidVersion(coreVer) Then Err.Raise vbObjectError + 2000, "Updater", "Manifesto sem coreVersion válida."
        url = JsonGetString(json, "coreUrl")
        tmp = Loader_TmpDir() & "\" & Loader_CoreFileName()
        If Len(Dir(tmp)) > 0 Then Kill tmp
        Transport_DownloadFile url, tmp
        sha = LCase$(JsonGetString(json, "coreSha256"))
        If Len(sha) > 0 Then
            If Transport_Sha256File(tmp) <> sha Then
                Kill tmp
                Err.Raise vbObjectError + 2001, "Updater", "Arquivo baixado corrompido (SHA-256 não confere)."
            End If
        Else
            Loader_Log "Updater_Apply: manifesto sem coreSha256; instalando sem verificação de integridade."
        End If
        destDir = Loader_VersionsDir() & "\" & coreVer
        If Len(Dir(destDir, vbDirectory)) = 0 Then MkDir destDir
        dest = destDir & "\" & Loader_CoreFileName()
        If Len(Dir(dest)) > 0 Then Loader_CloseCore: Kill dest
        FileCopy tmp, dest
        Kill tmp
        url = JsonGetString(json, "tutorialUrl")
        If Len(url) > 0 Then
            On Error Resume Next
            Transport_DownloadFile url, destDir & "\tutorial.html"
            On Error GoTo Fail
        End If
        If Not Loader_OpenCore(dest) Then Err.Raise vbObjectError + 2003, "Updater", "Core instalado, mas não pôde ser carregado."
        Settings_Set KEY_LAST_NOTIFIED_CORE, coreVer
        Settings_Delete KEY_SKIPPED
        Loader_InvalidateRibbon
    End If
    If what = "loader" Or what = "both" Then
        url = JsonGetString(json, "loaderUrl")
        tmp = Loader_TmpDir() & "\ThinQQ.xlam"
        If Len(Dir(tmp)) > 0 Then Kill tmp
        Transport_DownloadFile url, tmp
        sha = LCase$(JsonGetString(json, "loaderSha256"))
        If Len(sha) > 0 Then
            If Transport_Sha256File(tmp) <> sha Then
                Kill tmp
                Err.Raise vbObjectError + 2002, "Updater", "Arquivo baixado corrompido (SHA-256 não confere)."
            End If
        Else
            Loader_Log "Updater_Apply: manifesto sem loaderSha256; instalando sem verificação de integridade."
        End If
        Settings_Set KEY_PENDING_LOADER, tmp
        ' A troca do loader é feita pelo core, fora da pilha de chamadas do loader.
        ' Sem core carregado nao ha quem aplique (e um OnTime para macro inexistente abriria dialogo): fica pendente.
        If Loader_CoreWorkbook() Is Nothing Then
            Loader_Log "Updater_Apply: core ausente; loader pendente em " & tmp
        Else
            Application.OnTime Now, "'" & Loader_CoreFileName() & "'!Core_ApplyPendingLoader"
        End If
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
    Loader_Log "Updater_Apply: " & en & " " & ed
    If Not quiet Then Loader_Msg "Falha ao atualizar: " & Updater_FriendlyError(en, ed), vbExclamation
    Updater_Apply = False
End Function
