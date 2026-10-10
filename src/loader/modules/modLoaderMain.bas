Attribute VB_Name = "modLoaderMain"
Option Explicit
Public Const APP_NAME As String = "ThinQQ"
Private mRootOverride As String
Private mStartupCheckAt As Date
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

Public Sub Loader_SetRootOverride(path As String)
    mRootOverride = path
End Sub

Public Function Loader_Root() As String
    If Len(mRootOverride) > 0 Then Loader_Root = mRootOverride Else Loader_Root = Environ$("LOCALAPPDATA") & "\ThinQQ"
End Function

Public Function Loader_VersionsDir() As String
    Loader_VersionsDir = Loader_Root() & "\versions"
End Function

Public Function Loader_TmpDir() As String
    Loader_TmpDir = Loader_Root() & "\tmp"
End Function

Public Function Loader_CoreFileName() As String
    Loader_CoreFileName = "ThinQQ.core" & BUILD_SUFFIX & ".xlam"
End Function

Public Sub Loader_EnsureFolders()
    EnsureDir Loader_Root()
    EnsureDir Loader_VersionsDir()
    EnsureDir Loader_TmpDir()
    EnsureDir Loader_Root() & "\logs"
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
    Loader_EnsureFolders
    f = FreeFile
    Open Loader_Root() & "\logs\thinqq.log" For Append As #f
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
    Dim wb As Workbook
    Set wb = Loader_CoreWorkbook()
    If Not wb Is Nothing Then
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then Loader_OpenCore = True: Exit Function
        Loader_CloseCore
    End If
    On Error GoTo Fail
    Set wb = Application.Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
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
    ver = Loader_FindNewestCoreVersion(Loader_VersionsDir())
    If ver = "" Then
        Loader_Log "core não encontrado em " & Loader_VersionsDir()
        If MsgBox("O ThinQQ está instalado, mas o módulo principal (core) não foi encontrado." & vbLf & _
                  "Deseja baixá-lo agora pela fonte de atualização?", vbYesNo + vbQuestion, APP_NAME) = vbYes Then
            Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!Updater_CheckInteractive"
        End If
        Exit Sub
    End If
    If Not Loader_OpenCore(Loader_VersionsDir() & "\" & ver & "\" & Loader_CoreFileName()) Then
        MsgBox "Não foi possível carregar o ThinQQ core " & ver & ". Veja o log em " & Loader_Root() & "\logs.", vbExclamation, APP_NAME
        Exit Sub
    End If
    mStartupCheckAt = Now + TimeSerial(0, 0, 8)
    Application.OnTime mStartupCheckAt, "'" & ThisWorkbook.Name & "'!Updater_StartupCheck"
    Exit Sub
Fail:
    Loader_Log "Loader_Startup: " & Err.Description
End Sub

Public Sub Loader_Shutdown()
    On Error Resume Next
    If mStartupCheckAt <> 0 Then
        Application.OnTime mStartupCheckAt, "'" & ThisWorkbook.Name & "'!Updater_StartupCheck", , False
        mStartupCheckAt = 0
    End If
    Loader_CloseCore
End Sub
