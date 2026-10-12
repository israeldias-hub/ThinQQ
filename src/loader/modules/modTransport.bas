Attribute VB_Name = "modTransport"
Option Explicit
Private Const HTTP_TIMEOUT_TEXT_SEC As Long = 20
Private Const HTTP_TIMEOUT_FILE_SEC As Long = 120
' Acesso a manifesto/arquivos por HTTP(S) (MSXML2.XMLHTTP, respeita proxy do usuário) ou por caminho (UNC/local).

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

' GET assincrono com prazo: um servidor travado nao congela o Excel.
Private Function HttpGetBytes(url As String, timeoutSec As Long) As Byte()
    Dim http As Object, sep As String, deadline As Double
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    sep = IIf(InStr(url, "?") > 0, "&", "?")
    http.Open "GET", url & sep & "_ts=" & Format$(CDbl(Now) * 86400, "0"), True
    http.setRequestHeader "Cache-Control", "no-cache"
    http.Send
    deadline = CDbl(Now) * 86400 + timeoutSec
    Do While http.readyState <> 4
        If CDbl(Now) * 86400 > deadline Then
            On Error Resume Next
            http.Abort
            On Error GoTo 0
            Err.Raise vbObjectError + 1004, "Transport", "Tempo esgotado ao acessar " & url
        End If
        DoEvents
    Loop
    If http.Status <> 200 Then Err.Raise vbObjectError + 1001, "Transport", "HTTP " & http.Status & " ao acessar " & url
    HttpGetBytes = http.responseBody
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
