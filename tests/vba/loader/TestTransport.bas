Attribute VB_Name = "TestTransport"
Option Explicit

Private Function HttpBase() As String
    Dim p As String
    p = RepoPath() & "\tests\results\http-base.txt"
    If Len(Dir(p)) = 0 Then Err.Raise vbObjectError + 900, "TestTransport", "servidor HTTP de fixtures ausente (Node não encontrado?)"
    HttpBase = Transport_ReadTextFile(p)
End Function

Public Sub Test_Transport_IsHttp()
    AssertTrue Transport_IsHttp("https://x/y.json")
    AssertTrue Transport_IsHttp("HTTP://x")
    AssertFalse Transport_IsHttp("\servidor\pasta\manifest.json")
    AssertFalse Transport_IsHttp("C:\a.json")
End Sub

Public Sub Test_Transport_ExpandPath()
    AssertEqual Environ$("TEMP") & "\x\y", Transport_ExpandPath("%TEMP%\x\y")
    AssertEqual "C:\sem\variavel", Transport_ExpandPath("C:\sem\variavel")
End Sub

Public Sub Test_Transport_File_RoundTrip_Utf8()
    Dim d As String, p As String
    d = TempDir(): p = d & "\t.txt"
    Transport_WriteTextFile p, "ação & café"
    AssertEqual "ação & café", Transport_ReadTextFile(p)
    AssertEqual "ação & café", Transport_FetchText(p)
End Sub

Public Sub Test_Transport_DownloadFile_FromPath()
    Dim d As String
    d = TempDir()
    Transport_WriteTextFile d & "\src.txt", "abc"
    Transport_DownloadFile d & "\src.txt", d & "\dst.txt"
    AssertEqual "abc", Transport_ReadTextFile(d & "\dst.txt")
End Sub

Public Sub Test_Transport_Sha256_KnownVector()
    Dim d As String, f As Integer, b() As Byte
    d = TempDir()
    b = StrConv("abc", vbFromUnicode)
    f = FreeFile: Open d & "\abc.bin" For Binary As #f: Put #f, , b: Close #f
    AssertEqual "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", Transport_Sha256File(d & "\abc.bin")
    f = FreeFile: Open d & "\empty.bin" For Binary As #f: Close #f
    AssertEqual "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855", Transport_Sha256File(d & "\empty.bin")
End Sub

Public Sub Test_Transport_FetchText_MissingFile_Raises()
    Dim raised As Boolean
    On Error Resume Next
    Transport_FetchText "C:\nao\existe\manifest.json"
    raised = (Err.Number <> 0)
    On Error GoTo 0
    AssertTrue raised, "deveria lançar erro para arquivo inexistente"
End Sub

Public Sub Test_Transport_Http_FetchText()
    AssertEqual "olá ThinQQ", Transport_FetchText(HttpBase() & "http/hello.txt")
End Sub

Public Sub Test_Transport_Http_404_Raises()
    Dim raised As Boolean
    On Error Resume Next
    Transport_FetchText HttpBase() & "http/nao-existe.txt"
    raised = (Err.Number <> 0)
    On Error GoTo 0
    AssertTrue raised, "deveria lançar erro em HTTP 404"
End Sub

Public Sub Test_Transport_Http_DownloadFile()
    Dim d As String
    d = TempDir()
    Transport_DownloadFile HttpBase() & "http/hello.txt", d & "\hello.txt"
    AssertEqual "olá ThinQQ", Transport_ReadTextFile(d & "\hello.txt")
End Sub
