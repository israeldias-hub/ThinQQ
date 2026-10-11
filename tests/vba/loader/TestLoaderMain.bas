Attribute VB_Name = "TestLoaderMain"
Option Explicit

Private Sub MakeVersion(root As String, ver As String)
    MkDir root & "\" & ver
    Dim f As Integer
    f = FreeFile: Open root & "\" & ver & "\" & Loader_CoreFileName() For Output As #f: Print #f, "x": Close #f
End Sub

Public Sub Test_CoreFileName_UsesSuffix()
    AssertEqual "ThinQQ.core.test.xlam", Loader_CoreFileName()
End Sub

Public Sub Test_FindNewestCoreVersion()
    Dim d As String
    d = TempDir()
    MakeVersion d, "1.0.0"
    MakeVersion d, "1.10.0"
    MakeVersion d, "1.2.0"
    MkDir d & "\abc"                      ' pasta inválida: ignorada
    MkDir d & "\2.0.0"                    ' sem arquivo do core: ignorada
    AssertEqual "1.10.0", Loader_FindNewestCoreVersion(d)
End Sub

Public Sub Test_FindNewestCoreVersion_EmptyOrMissing()
    AssertEqual "", Loader_FindNewestCoreVersion(TempDir())
    AssertEqual "", Loader_FindNewestCoreVersion("C:\nao\existe")
End Sub

Public Sub Test_RootOverride_And_EnsureFolders()
    Dim d As String
    d = TempDir()
    Loader_SetRootOverride d & "\root"
    Loader_EnsureFolders
    AssertTrue Len(Dir(d & "\root\versions", vbDirectory)) > 0
    AssertTrue Len(Dir(d & "\root\tmp", vbDirectory)) > 0
    AssertEqual d & "\root\versions", Loader_VersionsDir()
    Loader_SetRootOverride ""
    AssertEqual Environ$("LOCALAPPDATA") & "\ThinQQ", Loader_Root()
End Sub

' M4b: downloads fora da pasta confiavel (%LOCALAPPDATA%\ThinQQ e subpastas).
Public Sub Test_TmpDir_IsOutsideTrustedRoot()
    Dim t As String
    Loader_SetRootOverride ""
    t = Loader_TmpDir()
    AssertEqual Environ$("TEMP") & "\ThinQQ", t
    AssertFalse LCase$(Left$(t, Len(Loader_Root()))) = LCase$(Loader_Root()), "tmp nao pode ficar dentro da raiz confiavel"
End Sub

Public Sub Test_RunCore_Ping()
    ' O run-tests.ps1 abre dist\ThinQQ.core.test.xlam antes dos testes do loader.
    AssertTrue Not Loader_CoreWorkbook() Is Nothing, "core de teste não está aberto"
    AssertEqual "pong", Loader_RunCore("Core_Ping")
    AssertTrue IsValidVersion(Loader_CurrentCoreVersion())
End Sub

Public Sub Test_OpenCore_SamePath_IsNoop()
    Dim p As String
    p = Loader_CoreWorkbook().FullName
    AssertTrue Loader_OpenCore(p)
    AssertEqual p, Loader_CoreWorkbook().FullName
End Sub

Public Sub Test_RunCore_TooManyArgs_Raises()
    Dim n As Long
    On Error Resume Next
    Loader_RunCore "Core_Ping", 1, 2, 3, 4
    n = Err.Number
    On Error GoTo 0
    AssertEqual vbObjectError + 513, n
End Sub

Public Sub Test_RunCore_MissingProc_ReturnsEmpty_NoRaise()
    Dim v As Variant, n As Long
    Loader_SetRootOverride TempDir() & "\r"
    On Error Resume Next
    v = Loader_RunCore("Core_NaoExiste")
    n = Err.Number
    On Error GoTo 0
    Loader_SetRootOverride ""
    AssertEqual 0, n
    AssertTrue IsEmpty(v), "esperado Empty"
End Sub

Public Sub Test_RunCore_MissingProc_Logs()
    Dim d As String, f As Integer, txt As String
    d = TempDir() & "\r"
    Loader_SetRootOverride d
    Loader_RunCore "Core_NaoExiste"
    Loader_SetRootOverride ""
    f = FreeFile: Open d & "\logs\thinqq.log" For Input As #f
    txt = Input$(LOF(f), f): Close #f
    AssertTrue InStr(txt, "Core_NaoExiste") > 0, "log sem a entrada"
End Sub

' Build Test sem raiz de teste: o log vai para %TEMP%\ThinQQ-test-logs, nunca para o log REAL do usuario (%LOCALAPPDATA%\ThinQQ\logs);
' com a raiz de teste continua em <raiz>\logs (Test_RunCore_MissingProc_Logs). Release nao muda (a raiz de teste e inerte).
Public Sub Test_Loader_Log_TestBuildWritesToTempFolder()
    Dim marker As String, p As String, f As Integer, txt As String
    Loader_SetRootOverride ""
    AssertEqual Environ$("TEMP") & "\ThinQQ-test-logs", Loader_LogDir()
    AssertFalse LCase$(Left$(Loader_LogDir(), Len(Loader_Root()))) = LCase$(Loader_Root()), "o log de teste nao pode ficar dentro da pasta real do usuario"
    Randomize
    marker = "tqq-loader-log-probe-" & Format$(Now, "hhnnss") & "-" & CStr(Int(Rnd * 100000))
    Loader_Log marker
    p = Loader_LogDir() & "\thinqq.log"
    AssertTrue Len(Dir(p)) > 0, "log de teste criado em " & p
    f = FreeFile
    Open p For Binary Access Read As #f
    txt = Space$(LOF(f))
    Get #f, , txt
    Close #f
    AssertTrue InStr(txt, marker) > 0, "a linha foi para o log de teste"
End Sub

' Com a raiz de teste o log fica sob ela.
Public Sub Test_Loader_LogDir_WithRootOverride_IsUnderRoot()
    Dim d As String
    d = TempDir() & "\r"
    Loader_SetRootOverride d
    AssertEqual d & "\logs", Loader_LogDir()
    Loader_SetRootOverride ""
End Sub

' Build Test SEM raiz de teste: Loader_EnsureFolders (chamada por Loader_Startup e por Updater_Apply) nao pode criar as pastas REAIS do usuario
' em %LOCALAPPDATA%\ThinQQ (a mesma regra do log, Test_Loader_Log_TestBuildWritesToTempFolder); cria so a pasta de log e a de downloads, em %TEMP%.
' Nao ha como redirecionar LOCALAPPDATA (Environ$ nao enxerga mudancas feitas no processo), e numa maquina que ja usou o ThinQQ a pasta real existe:
' por isso o teste tem duas partes. (1) Comportamento: se a raiz real NAO existia antes, nao pode existir depois (vale em maquina limpa/CI);
' em qualquer maquina as pastas de %TEMP% passam a existir. (2) Guarda no codigo: o procedimento le o build e a raiz de teste ANTES de qualquer
' EnsureDir da raiz (le o proprio projeto VBA, como Test_Ribbon_AllCallbacks_HaveErrorContainment).
Public Sub Test_EnsureFolders_TestBuildWithoutRootOverride_DoesNotCreateLocalAppDataFolders()
    Dim fso As Object, realRoot As String, existed As Boolean, cm As Object, code As String, pGuard As Long, pRoot As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    Loader_SetRootOverride ""
    realRoot = Loader_Root()
    AssertEqual Environ$("LOCALAPPDATA") & "\ThinQQ", realRoot, "sem raiz de teste a raiz e a do usuario"
    existed = fso.FolderExists(realRoot)
    Loader_EnsureFolders
    If Not existed Then AssertFalse fso.FolderExists(realRoot), "EnsureFolders criou a raiz REAL do usuario em build Test"
    AssertTrue fso.FolderExists(Loader_LogDir()), "a pasta de log de teste (em %TEMP%) existe"
    AssertTrue fso.FolderExists(Loader_TmpDir()), "a pasta de downloads (em %TEMP%) existe"
    Set cm = ThisWorkbook.VBProject.VBComponents("modLoaderMain").CodeModule
    code = cm.Lines(cm.ProcStartLine("Loader_EnsureFolders", 0), cm.ProcCountLines("Loader_EnsureFolders", 0))
    pGuard = InStr(code, "BUILD_CONFIG = ""Test""")
    pRoot = InStr(code, "EnsureDir Loader_Root()")
    AssertTrue pGuard > 0 And pRoot > 0, "Loader_EnsureFolders: guarda do build Test e criacao da raiz nao encontradas"
    AssertTrue pGuard < pRoot, "a guarda do build Test tem de vir ANTES de criar a raiz"
End Sub

' Loader_Startup sem core instalado: a pergunta "baixar agora?" passa por Loader_Msg (no build Test so registra e devolve vbOK: nunca bloqueia
' e nao agenda a atualizacao). Antes era um MsgBox direto, que travaria o Excel invisivel dos testes (e fugia da regra "MsgBox so em Loader_Msg").
Public Sub Test_Startup_NoCoreInstalled_AsksThroughLoaderMsg_NoBlocking()
    Dim d As String
    d = TempDir() & "\r"
    Loader_SetRootOverride d
    Loader_LastMsg = ""
    Loader_Startup
    Loader_SetRootOverride ""
    AssertTrue InStr(Loader_LastMsg, "módulo principal (core) não foi encontrado") > 0, "a pergunta nao passou por Loader_Msg: <" & Loader_LastMsg & ">"
End Sub
