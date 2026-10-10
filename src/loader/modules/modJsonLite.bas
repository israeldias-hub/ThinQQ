Attribute VB_Name = "modJsonLite"
Option Explicit
' Leitor mínimo de JSON "plano": chaves de primeiro nível com valores string/número/booleano.

Public Function JsonGetString(json As String, key As String) As String
    Dim p As Long, q As Long, needle As String
    needle = """" & key & """"
    p = InStr(1, json, needle, vbBinaryCompare)
    If p = 0 Then Exit Function
    p = InStr(p + Len(needle), json, ":")
    If p = 0 Then Exit Function
    p = p + 1
    Do While p <= Len(json)
        If InStr(1, " " & vbTab & vbCr & vbLf, Mid$(json, p, 1)) = 0 Then Exit Do
        p = p + 1
    Loop
    If p > Len(json) Then Exit Function
    If Mid$(json, p, 1) = """" Then
        JsonGetString = JsonReadString(json, p + 1)
    Else
        q = p
        Do While q <= Len(json)
            If InStr(1, ",}]" & vbCr & vbLf & " " & vbTab, Mid$(json, q, 1)) > 0 Then Exit Do
            q = q + 1
        Loop
        JsonGetString = Mid$(json, p, q - p)
    End If
End Function

Public Function JsonGetNumber(json As String, key As String) As Double
    Dim s As String
    s = JsonGetString(json, key)
    If Len(s) = 0 Then Exit Function
    JsonGetNumber = Val(s)   ' Val usa ponto decimal, independente do locale
End Function

' Lê uma string JSON a partir da posição após a aspa de abertura, desfazendo escapes.
Private Function JsonReadString(json As String, startPos As Long) As String
    Dim i As Long, ch As String, out As String, esc As String
    i = startPos
    Do While i <= Len(json)
        ch = Mid$(json, i, 1)
        If ch = """" Then Exit Do
        If ch = "\" Then
            esc = Mid$(json, i + 1, 1)
            Select Case esc
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case "b": out = out & Chr$(8)
                Case "f": out = out & Chr$(12)
                Case "u": out = out & ChrW$(CLng("&H" & Mid$(json, i + 2, 4))): i = i + 4
                Case Else: out = out & esc    ' \" \ \/
            End Select
            i = i + 2
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    JsonReadString = out
End Function
