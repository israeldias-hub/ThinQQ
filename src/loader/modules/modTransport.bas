Attribute VB_Name = "modTransport"
Option Explicit
Private Const HTTP_TIMEOUT_TEXT_SEC As Long = 20
Private Const HTTP_TIMEOUT_FILE_SEC As Long = 120
Private Const INET_SETTINGS_KEY As String = "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\"
Private mLastPath As String
' Acesso a manifesto/arquivos por HTTP(S) (MSXML2.XMLHTTP, respeita proxy do usuário; reserva MSXML2.ServerXMLHTTP quando
' o XMLHTTP falha por transporte, ex.: redirecionamento do GitHub para outro domínio) ou por caminho (UNC/local).

Public Function Transport_IsHttp(source As String) As Boolean
    Transport_IsHttp = (LCase$(Left$(source, 7)) = "http://") Or (LCase$(Left$(source, 8)) = "https://")
End Function

' Expande %VARIAVEIS% de ambiente (ex.: %USERPROFILE%\OneDrive\ThinQQ\manifest.json).
Public Function Transport_ExpandPath(p As String) As String
    Dim out As String, i As Long, j As Long, name As String
    out = p
    i = InStr(1, out, "%")
    Do While i > 0
        j = InStr(i + 1, out, "%")
        If j = 0 Then Exit Do
        name = Mid$(out, i + 1, j - i - 1)
        out = Left$(out, i - 1) & Environ$(name) & Mid$(out, j + 1)
        i = InStr(i + Len(Environ$(name)), out, "%")
    Loop
    Transport_ExpandPath = out
End Function

Public Function Transport_FetchText(source As String) As String
    If Transport_IsHttp(source) Then
        Transport_FetchText = BytesToUtf8(HttpGetBytes(source, HTTP_TIMEOUT_TEXT_SEC))
    Else
        Transport_FetchText = Transport_ReadTextFile(Transport_ExpandPath(source))
    End If
End Function

Public Sub Transport_DownloadFile(source As String, destPath As String)
    If Transport_IsHttp(source) Then
        SaveBytes HttpGetBytes(source, HTTP_TIMEOUT_FILE_SEC), destPath
    Else
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile Transport_ExpandPath(source), destPath, True
    End If
End Sub

Public Function Transport_ReadTextFile(path As String) As String
    Dim stm As Object
    If Len(Dir(path)) = 0 Then Err.Raise vbObjectError + 1002, "Transport", "Arquivo não encontrado: " & path
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.LoadFromFile path
    Transport_ReadTextFile = stm.ReadText(-1)
    stm.Close
End Function

Public Sub Transport_WriteTextFile(path As String, text As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2    ' adSaveCreateOverWrite
    stm.Close
End Sub

' SHA-256 do arquivo (hex minusculo) em VBA puro (modSha256): sem processo filho. Levanta erro se o arquivo nao
' puder ser lido; nunca devolve texto vazio.
Public Function Transport_Sha256File(path As String) As String
    Transport_Sha256File = Sha256_File(path)
End Function

' Caminho que obteve a resposta HTTP da ultima requisicao: "XMLHTTP" | "ServerXMLHTTP" ("" = nenhum respondeu).
Public Function Transport_LastPath() As String
    Transport_LastPath = mLastPath
End Function

' Proxy estatico para a reserva, a partir dos valores de Internet Settings do usuario (funcao pura: quem chama le o registro).
' proxyServer: "host:porta" (vale para todos os protocolos) ou "http=h1:p1;https=h2:p2;...": usa a entrada do protocolo da URL,
' senao uma entrada geral (sem "=") da lista; nada disso => "" (conexao direta, como o WinINet faz).
Public Function Transport_ProxyFor(url As String, proxyEnable As Long, proxyServer As String) As String
    Dim scheme As String, parts() As String, i As Long, item As String, general As String
    If proxyEnable = 0 Then Exit Function
    If Len(Trim$(proxyServer)) = 0 Then Exit Function
    If InStr(proxyServer, "=") = 0 Then Transport_ProxyFor = Trim$(proxyServer): Exit Function
    If InStr(url, ":") > 0 Then scheme = LCase$(Left$(url, InStr(url, ":") - 1))
    parts = Split(proxyServer, ";")
    For i = LBound(parts) To UBound(parts)
        item = Trim$(parts(i))
        If InStr(item, "=") = 0 Then
            If Len(general) = 0 Then general = item
        ElseIf LCase$(Trim$(Left$(item, InStr(item, "=") - 1))) = scheme Then
            Transport_ProxyFor = Trim$(Mid$(item, InStr(item, "=") + 1))
            Exit Function
        End If
    Next
    Transport_ProxyFor = general
End Function

' GET assincrono com prazo: um servidor travado nao congela o Excel.
' 1a tentativa: MSXML2.XMLHTTP (WinINet: proxy/PAC do usuario). Ela bloqueia redirecionamento para OUTRO dominio (o GitHub
' responde 302 de github.com para release-assets.githubusercontent.com): termina com HTTP 0 ou "Acesso negado". Em falha de
' TRANSPORTE (excecao, status 0, 12xxx) repete a mesma requisicao com MSXML2.ServerXMLHTTP (WinHTTP, segue o redirecionamento)
' com o proxy estatico do usuario. Respostas HTTP reais (404, 500...) nao disparam a reserva.
Private Function HttpGetBytes(url As String, timeoutSec As Long) As Byte()
    Dim sep As String, reqUrl As String, st As Long, body As Variant, cause1 As String, cause2 As String
    mLastPath = ""
    sep = IIf(InStr(url, "?") > 0, "&", "?")
    reqUrl = url & sep & "_ts=" & Format$(CDbl(Now) * 86400, "0")
    If HttpAttempt(False, reqUrl, url, timeoutSec, st, body, cause1) Then
        mLastPath = "XMLHTTP"
    Else
        Loader_Log "Transport: 1a tentativa (XMLHTTP) falhou (" & cause1 & "); tentando a reserva (ServerXMLHTTP): " & url
        If Not HttpAttempt(True, reqUrl, url, timeoutSec, st, body, cause2) Then
            Loader_Log "Transport: reserva (ServerXMLHTTP) falhou (" & cause2 & "): " & url
            Err.Raise vbObjectError + 1005, "Transport", "Sem conexão ao acessar " & url & " (1ª tentativa (XMLHTTP): " & cause1 & "; 2ª tentativa (ServerXMLHTTP): " & cause2 & ")."
        End If
        mLastPath = "ServerXMLHTTP"
    End If
    If st <> 200 Then Err.Raise vbObjectError + 1001, "Transport", "HTTP " & st & " ao acessar " & url
    Loader_Log "Transport: baixado por " & mLastPath & ": " & url
    HttpGetBytes = body
End Function

' Uma tentativa de GET. True = houve resposta HTTP (status em st; corpo em body quando 200). False = falha de transporte
' (excecao no Open/Send/espera, status 0 ou 12xxx do WinINet), com a causa em cause. Prazo esgotado levanta 1004 (sem reserva).
Private Function HttpAttempt(server As Boolean, reqUrl As String, url As String, timeoutSec As Long, st As Long, body As Variant, cause As String) As Boolean
    Dim http As Object, deadline As Double, timedOut As Boolean
    st = 0: cause = ""
    On Error GoTo NetFail
    If server Then
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.setTimeouts timeoutSec * 1000&, timeoutSec * 1000&, timeoutSec * 1000&, timeoutSec * 1000&
        ApplyUserProxy http, url
    Else
        Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    End If
    http.Open "GET", reqUrl, True
    http.setRequestHeader "Cache-Control", "no-cache"
    http.Send
    deadline = CDbl(Now) * 86400 + timeoutSec
    Do While http.readyState <> 4
        ' ServerXMLHTTP assincrono so revela erro de conexao por waitForResponse (sem ele fica em readyState 1 ate o prazo).
        If server Then
            If http.waitForResponse(0.1) Then Exit Do
        End If
        If CDbl(Now) * 86400 > deadline Then timedOut = True: Exit Do
        DoEvents
    Loop
    If timedOut Then
        On Error Resume Next
        http.Abort
        On Error GoTo 0
        Err.Raise vbObjectError + 1004, "Transport", "Tempo esgotado ao acessar " & url
    End If
    st = http.Status
    If st = 0 Or (st >= 12000 And st < 13000) Then
        cause = NetStatusText(st)
        Exit Function
    End If
    If st = 200 Then body = http.responseBody
    HttpAttempt = True
    Exit Function
NetFail:
    cause = Trim$(Replace(Replace(Err.Description, vbCr, " "), vbLf, " "))
    If Len(cause) = 0 Then cause = "erro " & Err.Number
    st = 0
    HttpAttempt = False
End Function

Private Function NetStatusText(st As Long) As String
    Select Case st
        Case 0: NetStatusText = "sem resposta (HTTP 0; redirecionamento para outro domínio bloqueado?)"
        Case 12002: NetStatusText = "erro de rede 12002 (tempo esgotado)"
        Case 12007: NetStatusText = "erro de rede 12007 (nome do servidor não encontrado)"
        Case 12029: NetStatusText = "erro de rede 12029 (não foi possível conectar)"
        Case 12030, 12031: NetStatusText = "erro de rede " & st & " (conexão interrompida)"
        Case 12175: NetStatusText = "erro de rede 12175 (falha de segurança/certificado)"
        Case Else: NetStatusText = "erro de rede " & st
    End Select
End Function

' Proxy da reserva: o estatico do usuario (HKCU Internet Settings) com setProxy 2; sem proxy, conexao direta.
' Script de configuracao automatica (AutoConfigURL/PAC) nao e suportado pela reserva: so registra no log.
' Loopback (127.0.0.1, localhost, [::1]) e sempre direto, sem ler o registro (o WinINet tambem nunca manda o loopback ao proxy).
Private Sub ApplyUserProxy(http As Object, url As String)
    Dim enable As Long, srv As String, bypass As String, pac As String, px As String
    If IsLoopbackUrl(url) Then Exit Sub
    enable = Val(RegReadText(INET_SETTINGS_KEY & "ProxyEnable"))
    srv = RegReadText(INET_SETTINGS_KEY & "ProxyServer")
    bypass = Trim$(RegReadText(INET_SETTINGS_KEY & "ProxyOverride"))
    pac = RegReadText(INET_SETTINGS_KEY & "AutoConfigURL")
    If Len(pac) > 0 Then Loader_Log "Transport: configuração automática de proxy (script PAC) não é usada pela reserva (ServerXMLHTTP)."
    px = Transport_ProxyFor(url, enable, srv)
    If Len(px) = 0 Then Exit Sub
    If Len(bypass) > 0 Then http.setProxy 2, px, bypass Else http.setProxy 2, px
    Loader_Log "Transport: reserva usa o proxy configurado do usuário: " & px
End Sub

Private Function IsLoopbackUrl(url As String) As Boolean
    Dim h As String, p As Long
    h = LCase$(Mid$(url, InStr(url, "//") + 2))
    p = InStr(h, "/")
    If p > 0 Then h = Left$(h, p - 1)
    If Left$(h, 1) = "[" Then
        h = Left$(h, InStr(h & "]", "]"))
    ElseIf InStr(h, ":") > 0 Then
        h = Left$(h, InStr(h, ":") - 1)
    End If
    IsLoopbackUrl = (h = "localhost" Or h = "127.0.0.1" Or h = "[::1]")
End Function

' Valor do registro como texto ("" se ausente ou ilegivel). Somente leitura.
Private Function RegReadText(key As String) As String
    On Error Resume Next
    RegReadText = CStr(CreateObject("WScript.Shell").RegRead(key))
End Function

Private Sub SaveBytes(data() As Byte, destPath As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open
    stm.Write data
    stm.SaveToFile destPath, 2
    stm.Close
End Sub

Private Function BytesToUtf8(data() As Byte) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open
    stm.Write data
    stm.Position = 0: stm.Type = 2: stm.Charset = "utf-8"
    BytesToUtf8 = stm.ReadText(-1)
    stm.Close
End Function
