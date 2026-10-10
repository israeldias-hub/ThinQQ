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
