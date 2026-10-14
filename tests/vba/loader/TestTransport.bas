Attribute VB_Name = "TestTransport"
Option Explicit

Private Function HttpBase() As String
    Dim p As String
    p = RepoPath() & "\tests\results\http-base.txt"
    If Len(Dir(p)) = 0 Then Err.Raise vbObjectError + 900, "TestTransport", "servidor HTTP de fixtures ausente (Node não encontrado?)"
    HttpBase = Transport_ReadTextFile(p)
End Function

' Core de teste publicado pelo run-tests em tests\fixtures\release (servido pelo node em /release/).
Private Function FixtureCore() As String
    FixtureCore = RepoPath() & "\tests\fixtures\release\ThinQQ.core.test.xlam"
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
    Dim n As Long
    On Error Resume Next
    Transport_FetchText "C:\nao\existe\manifest.json"
    n = Err.Number
    On Error GoTo 0
    AssertEqual CStr(vbObjectError + 1002), CStr(n), "erro esperado para arquivo inexistente"
End Sub

Public Sub Test_Transport_Sha256_SpecialPath()
    Dim d As String, f As Integer, b() As Byte
    d = TempDir() & "\a b & é"
    MkDir d
    b = StrConv("abc", vbFromUnicode)
    f = FreeFile: Open d & "\abc.bin" For Binary As #f: Put #f, , b: Close #f
    AssertEqual "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", Transport_Sha256File(d & "\abc.bin")
End Sub

Public Sub Test_Transport_Sha256_MissingFile_Raises()
    Dim n As Long, r As String
    On Error Resume Next
    r = Transport_Sha256File(TempDir() & "\nao-existe.bin")
    n = Err.Number
    On Error GoTo 0
    AssertTrue n <> 0, "deveria lançar erro para arquivo inexistente"
    AssertEqual "", r
End Sub

Public Sub Test_Transport_Http_FetchText()
    AssertEqual "olá ThinQQ", Transport_FetchText(HttpBase() & "http/hello.txt")
End Sub

Public Sub Test_Transport_Http_404_Raises()
    Dim n As Long, u As String
    u = HttpBase() & "http/nao-existe.txt"
    On Error Resume Next
    Transport_FetchText u
    n = Err.Number
    On Error GoTo 0
    AssertEqual CStr(vbObjectError + 1001), CStr(n), "erro esperado em HTTP 404"
End Sub

Public Sub Test_Transport_Http_DownloadFile()
    Dim d As String
    d = TempDir()
    Transport_DownloadFile HttpBase() & "http/hello.txt", d & "\hello.txt"
    AssertEqual "olá ThinQQ", Transport_ReadTextFile(d & "\hello.txt")
End Sub

' O GitHub responde 302 de github.com para OUTRO dominio (release-assets.githubusercontent.com); o MSXML2.XMLHTTP bloqueia
' esse redirecionamento (HTTP 0). /redir-host/ do servidor local faz o mesmo entre 127.0.0.1 e localhost: o arquivo
' tem de chegar inteiro (tamanho e SHA-256 do original).
Public Sub Test_Transport_Http_CrossHostRedirect_DownloadsWholeFile()
    Dim d As String
    d = TempDir()
    Transport_DownloadFile HttpBase() & "redir-host/release/ThinQQ.core.test.xlam", d & "\core.xlam"
    AssertEqual CStr(FileLen(FixtureCore())), CStr(FileLen(d & "\core.xlam")), "tamanho do arquivo baixado"
    AssertEqual Transport_Sha256File(FixtureCore()), Transport_Sha256File(d & "\core.xlam"), "SHA-256 do arquivo baixado"
    AssertEqual "ServerXMLHTTP", Transport_LastPath(), "o arquivo deve chegar pela reserva"
End Sub

' Texto (manifesto) pelo mesmo caminho de reserva.
Public Sub Test_Transport_Http_CrossHostRedirect_FetchText()
    AssertEqual "olá ThinQQ", Transport_FetchText(HttpBase() & "redir-host/http/hello.txt")
    AssertEqual "ServerXMLHTTP", Transport_LastPath()
End Sub

' Redirecionamento no MESMO host: o XMLHTTP segue sozinho (sem reserva).
Public Sub Test_Transport_Http_SameHostRedirect_Downloads()
    AssertEqual "olá ThinQQ", Transport_FetchText(HttpBase() & "redir-same/http/hello.txt")
    AssertEqual "XMLHTTP", Transport_LastPath()
End Sub

' Sem redirecionamento: primeira tentativa, como sempre.
Public Sub Test_Transport_Http_Direct_UsesXmlHttp()
    Dim d As String
    d = TempDir()
    Transport_DownloadFile HttpBase() & "release/ThinQQ.core.test.xlam", d & "\core.xlam"
    AssertEqual Transport_Sha256File(FixtureCore()), Transport_Sha256File(d & "\core.xlam")
    AssertEqual "XMLHTTP", Transport_LastPath()
End Sub

' Resposta HTTP real (404) e erro do servidor: NAO dispara a reserva e mantem a mensagem de hoje.
Public Sub Test_Transport_Http_404_NoFallback_SameMessage()
    Dim n As Long, ds As String, u As String
    u = HttpBase() & "http/nao-existe.txt"
    On Error Resume Next
    Transport_FetchText u
    n = Err.Number: ds = Err.Description
    On Error GoTo 0
    AssertEqual CStr(vbObjectError + 1001), CStr(n)
    AssertEqual "HTTP 404 ao acessar " & u, ds
    AssertEqual "XMLHTTP", Transport_LastPath(), "404 nao pode passar pela reserva"
End Sub

' As duas tentativas falhando (porta fechada): erro 1005 com as duas causas, sem o cache-buster na URL.
Public Sub Test_Transport_Http_BothFail_MessageHasBothCauses()
    Dim n As Long, ds As String, u As String
    u = "http://127.0.0.1:1/manifest.json"
    On Error Resume Next
    Transport_FetchText u
    n = Err.Number: ds = Err.Description
    On Error GoTo 0
    AssertEqual CStr(vbObjectError + 1005), CStr(n), ds
    AssertTrue InStr(ds, "1ª tentativa (XMLHTTP): ") > 0, ds
    AssertTrue InStr(ds, "2ª tentativa (ServerXMLHTTP): ") > 0, ds
    AssertTrue InStr(ds, "12029") > 0, "causa da 1a tentativa (WinINet 12029): " & ds
    AssertTrue InStr(ds, u) > 0, ds
    AssertTrue InStr(ds, "_ts=") = 0, "sem cache-buster na mensagem: " & ds
    AssertEqual "", Transport_LastPath()
    AssertTrue InStr(Updater_FriendlyError(n, ds), ds) > 0, "a mensagem ao usuario traz as duas causas"
End Sub

' Proxy estatico do usuario para a reserva: funcao pura (valores de Internet Settings passados pelo teste; o registro real nao e lido).
Public Sub Test_Transport_ProxyFor_Values()
    AssertEqual "proxy.corp:8080", Transport_ProxyFor("https://github.com/a", 1, "proxy.corp:8080")
    AssertEqual "proxy.corp:8080", Transport_ProxyFor("http://x/a", 1, " proxy.corp:8080 ")
    AssertEqual "", Transport_ProxyFor("https://github.com/a", 0, "proxy.corp:8080"), "ProxyEnable=0: direto"
    AssertEqual "", Transport_ProxyFor("https://github.com/a", 1, ""), "sem servidor: direto"
    AssertEqual "", Transport_ProxyFor("https://github.com/a", 1, "   "), "so espacos: direto"
    AssertEqual "p2:443", Transport_ProxyFor("https://github.com/a", 1, "http=p1:80;https=p2:443;ftp=p3:21")
    AssertEqual "p1:80", Transport_ProxyFor("http://x/a", 1, "http=p1:80;https=p2:443;ftp=p3:21")
    AssertEqual "P2:443", Transport_ProxyFor("HTTPS://github.com/a", 1, "HTTP=p1:80; HTTPS=P2:443")
    AssertEqual "", Transport_ProxyFor("https://github.com/a", 1, "http=p1:80;ftp=p3:21"), "lista por protocolo sem https: direto"
    AssertEqual "geral:3128", Transport_ProxyFor("https://github.com/a", 1, "ftp=p3:21;geral:3128"), "entrada geral na lista"
End Sub
