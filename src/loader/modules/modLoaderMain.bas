Attribute VB_Name = "modLoaderMain"
Option Explicit
Public Const APP_NAME As String = "ThinQQ"
Private mRootOverride As String
Private mStartupCheckAt As Date
Private Const SWAP_FLAG_MAX_MINUTES As Double = 5
Public Loader_LastMsg As String      ' ultima mensagem enviada por Loader_Msg (usada nos testes)

Public Function Loader_Ping() As String
    Loader_Ping = "pong"
End Function

' Unico ponto de MsgBox do loader fora da faixa: em build de Teste apenas registra (nunca bloqueia).
Public Function Loader_Msg(prompt As String, Optional style As VbMsgBoxStyle = vbInformation) As VbMsgBoxResult
    Loader_LastMsg = prompt
    If BUILD_CONFIG = "Test" Then
        Loader_Log "Msg (teste): " & Replace(prompt, vbLf, " | ")
        Loader_Msg = vbOK
    Else
        Loader_Msg = MsgBox(prompt, style, APP_NAME)
    End If
End Function

' Gancho de teste: so existe no build de Teste. Em Release e inerte (um Application.Run de terceiros nao pode
' redirecionar a raiz de onde o loader carrega codigo).
Public Sub Loader_SetRootOverride(path As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mRootOverride = path
End Sub

Public Function Loader_Root() As String
    If Len(mRootOverride) > 0 Then Loader_Root = mRootOverride Else Loader_Root = Environ$("LOCALAPPDATA") & "\ThinQQ"
End Function

Public Function Loader_VersionsDir() As String
    Loader_VersionsDir = Loader_Root() & "\versions"
End Function

' Downloads e arquivos pendentes ficam em %TEMP%\ThinQQ, FORA da pasta confiavel (%LOCALAPPDATA%\ThinQQ e subpastas):
' um arquivo baixado e ainda nao verificado nunca esta em local confiavel. (Com a raiz de teste, fica sob ela.)
Public Function Loader_TmpDir() As String
    If Len(mRootOverride) > 0 Then
        Loader_TmpDir = mRootOverride & "\tmp"
    Else
        Loader_TmpDir = Environ$("TEMP") & "\ThinQQ"
    End If
End Function

Public Function Loader_CoreFileName() As String
    Loader_CoreFileName = "ThinQQ.core" & BUILD_SUFFIX & ".xlam"
End Function

' Pasta do log. Com a raiz de teste: <raiz>\logs. Build Test SEM raiz de teste: %TEMP%\ThinQQ-test-logs (os testes nao podem escrever no log
' REAL do usuario nem criar pastas em %LOCALAPPDATA%\ThinQQ). Demais builds: <raiz>\logs, como sempre.
Public Function Loader_LogDir() As String
    If Len(mRootOverride) = 0 And BUILD_CONFIG = "Test" Then
        Loader_LogDir = Environ$("TEMP") & "\ThinQQ-test-logs"
    Else
        Loader_LogDir = Loader_Root() & "\logs"
    End If
End Function

Public Sub Loader_EnsureFolders()
    EnsureDir Loader_Root()
    EnsureDir Loader_VersionsDir()
    EnsureDir Loader_TmpDir()
    EnsureDir Loader_LogDir()
End Sub

Private Sub EnsureDir(path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub
    Dim parent As String
    parent = fso.GetParentFolderName(path)
    If Len(parent) = 0 Then Exit Sub
    If Not fso.FolderExists(parent) Then EnsureDir parent
    fso.CreateFolder path
End Sub

Public Sub Loader_Log(msg As String)
    On Error Resume Next
    Dim f As Integer
    If Len(mRootOverride) = 0 And BUILD_CONFIG = "Test" Then
        EnsureDir Loader_LogDir()           ' build Test: so a pasta de log de %TEMP% (nada em %LOCALAPPDATA%\ThinQQ)
    Else
        Loader_EnsureFolders
    End If
    f = FreeFile
    Open Loader_LogDir() & "\thinqq.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & "loader" & vbTab & msg
    Close #f
End Sub

' Maior versão x.y.z em versionsDir que contenha o arquivo do core ("" se nenhuma).
Public Function Loader_FindNewestCoreVersion(versionsDir As String) As String
    Dim fso As Object, fld As Object, sub_ As Object, best As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(versionsDir) Then Exit Function
    Set fld = fso.GetFolder(versionsDir)
    For Each sub_ In fld.SubFolders
        If IsValidVersion(sub_.Name) Then
            If fso.FileExists(sub_.Path & "\" & Loader_CoreFileName()) Then
                If best = "" Then
                    best = sub_.Name
                ElseIf CompareVersions(sub_.Name, best) > 0 Then
                    best = sub_.Name
                End If
            End If
        End If
    Next
    Loader_FindNewestCoreVersion = best
End Function

' Versoes validas (com o arquivo do core) em ordem decrescente.
Public Function Loader_ListCoreVersions(versionsDir As String) As Collection
    Dim fso As Object, sub_ As Object, arr() As String, n As Long, i As Long, j As Long, t As String, res As Collection
    Set res = New Collection
    Set Loader_ListCoreVersions = res
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(versionsDir) Then Exit Function
    ReDim arr(0 To 0)
    For Each sub_ In fso.GetFolder(versionsDir).SubFolders
        If IsValidVersion(sub_.Name) Then
            If fso.FileExists(sub_.Path & "\" & Loader_CoreFileName()) Then
                ReDim Preserve arr(0 To n)
                arr(n) = sub_.Name: n = n + 1
            End If
        End If
    Next
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If CompareVersions(arr(j), arr(i)) > 0 Then t = arr(i): arr(i) = arr(j): arr(j) = t
        Next
    Next
    For i = 0 To n - 1
        res.Add arr(i)
    Next
End Function

' Abre o core mais novo que carregar de fato (health check incluso); se falhar, tenta a proxima versao.
' Retorna a versao aberta ("" se nenhuma abriu).
Public Function Loader_OpenBestCore(versionsDir As String) As String
    Dim v As Variant
    For Each v In Loader_ListCoreVersions(versionsDir)
        If Loader_OpenCore(versionsDir & "\" & CStr(v) & "\" & Loader_CoreFileName()) Then
            Loader_OpenBestCore = CStr(v)
            Exit Function
        End If
        Loader_Log "Loader_OpenBestCore: core " & CStr(v) & " nao carregou; tentando versao anterior"
    Next
End Function

Public Function Loader_CoreWorkbook() As Workbook
    On Error Resume Next
    Set Loader_CoreWorkbook = Application.Workbooks(Loader_CoreFileName())
    On Error GoTo 0
End Function

Public Sub Loader_CloseCore()
    Dim wb As Workbook
    Set wb = Loader_CoreWorkbook()
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Application.Run "'" & Loader_CoreFileName() & "'!Core_OnUnload"
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

' Abre (ou troca) o core. Retorna False se falhar.
Public Function Loader_OpenCore(path As String) As Boolean
    Dim wb As Workbook, prevAlerts As Boolean
    Set wb = Loader_CoreWorkbook()
    If Not wb Is Nothing Then
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            ' Core ja aberto neste caminho (ex.: apos a troca do loader): so reaponta o core para ESTE loader.
            On Error Resume Next
            Application.Run "'" & Loader_CoreFileName() & "'!Core_RegisterLoader", ThisWorkbook.Name, ThisWorkbook.FullName
            On Error GoTo 0
            Loader_OpenCore = True
            Exit Function
        End If
        Loader_CloseCore
    End If
    On Error GoTo Fail
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set wb = Application.Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    Application.DisplayAlerts = prevAlerts
    Dim pong As Variant
    On Error Resume Next
    pong = Application.Run("'" & wb.Name & "'!Core_Ping")
    On Error GoTo Fail
    If CStr(pong) <> "pong" Then
        Loader_Log "Loader_OpenCore: Core_Ping nao respondeu (macros desabilitadas?): " & path
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
        Loader_OpenCore = False
        Exit Function
    End If
    On Error Resume Next
    Application.Run "'" & Loader_CoreFileName() & "'!Core_RegisterLoader", ThisWorkbook.Name, ThisWorkbook.FullName
    On Error GoTo 0
    Loader_OpenCore = True
    Exit Function
Fail:
    Application.DisplayAlerts = prevAlerts
    Loader_Log "Loader_OpenCore falhou: " & path & " - " & Err.Description
    Loader_OpenCore = False
End Function

' Encaminha uma chamada ao core. Retorna Empty se o core não estiver carregado.
Public Function Loader_RunCore(procName As String, ParamArray args() As Variant) As Variant
    Dim macro As String
    If Loader_CoreWorkbook() Is Nothing Then Exit Function
    macro = "'" & Loader_CoreFileName() & "'!" & procName
    Dim n As Long
    n = UBound(args) - LBound(args) + 1
    If n > 3 Then Err.Raise vbObjectError + 513, "Loader_RunCore", "Loader_RunCore: máximo de 3 argumentos"
    On Error GoTo RunFail
    Select Case n
        Case 0: Loader_RunCore = Application.Run(macro)
        Case 1: Loader_RunCore = Application.Run(macro, args(0))
        Case 2: Loader_RunCore = Application.Run(macro, args(0), args(1))
        Case 3: Loader_RunCore = Application.Run(macro, args(0), args(1), args(2))
    End Select
    Exit Function
RunFail:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    Loader_RunCore = Empty
    Loader_Log "Loader_RunCore(" & procName & ") erro " & en & ": " & ed
End Function

Public Function Loader_CurrentCoreVersion() As String
    Dim v As Variant
    v = Loader_RunCore("Core_Version")
    If IsEmpty(v) Then Loader_CurrentCoreVersion = "0.0.0" Else Loader_CurrentCoreVersion = CStr(v)
End Function

Public Sub Loader_Startup()
    On Error GoTo Fail
    Dim ver As String
    Loader_EnsureFolders
    If Loader_FindNewestCoreVersion(Loader_VersionsDir()) = "" Then
        Loader_Log "core não encontrado em " & Loader_VersionsDir()
        If MsgBox("O ThinQQ está instalado, mas o módulo principal (core) não foi encontrado." & vbLf & _
                  "Deseja baixá-lo agora pela fonte de atualização?", vbYesNo + vbQuestion, APP_NAME) = vbYes Then
            Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!Updater_CheckInteractive"
        End If
        Exit Sub
    End If
    ver = Loader_OpenBestCore(Loader_VersionsDir())
    If ver = "" Then
        MsgBox "Não foi possível carregar o ThinQQ core. Veja o log em " & Loader_LogDir() & ".", vbExclamation, APP_NAME
        Exit Sub
    End If
    Updater_SchedulePendingLoader 4
    mStartupCheckAt = Now + TimeSerial(0, 0, 8)
    Application.OnTime mStartupCheckAt, "'" & ThisWorkbook.Name & "'!Updater_StartupCheck"
    Exit Sub
Fail:
    Loader_Log "Loader_Startup: " & Err.Description
End Sub

' True enquanto o core esta trocando o arquivo do loader (flag recente, < 5 min). Flag velha e ignorada e limpa.
Public Function Loader_SwapInProgress() As Boolean
    Dim t As String, age As Double
    On Error GoTo Fail
    t = Settings_Get(KEY_SWAP_IN_PROGRESS)
    If Len(t) = 0 Then Exit Function
    If IsDate(t) Then
        age = (Now - CDate(t)) * 1440
        If age >= 0 And age < SWAP_FLAG_MAX_MINUTES Then Loader_SwapInProgress = True: Exit Function
    End If
    Settings_Delete KEY_SWAP_IN_PROGRESS      ' obsoleta ou invalida
    Exit Function
Fail:
    Loader_SwapInProgress = False
End Function

' Decisao de fechar o core ao encerrar o loader: nao fecha durante a troca do loader (o core esta executando a troca).
Public Function Loader_ShouldCloseCoreOnShutdown() As Boolean
    Loader_ShouldCloseCoreOnShutdown = Not Loader_SwapInProgress()
End Function

Public Sub Loader_Shutdown()
    On Error Resume Next
    If mStartupCheckAt <> 0 Then
        Application.OnTime mStartupCheckAt, "'" & ThisWorkbook.Name & "'!Updater_StartupCheck", , False
        mStartupCheckAt = 0
    End If
    If Loader_ShouldCloseCoreOnShutdown() Then Loader_CloseCore
End Sub
