Attribute VB_Name = "modVersion"
Option Explicit

Public Function IsValidVersion(v As String) As Boolean
    Dim parts() As String, i As Long, k As Long
    If Len(v) = 0 Then Exit Function
    parts = Split(v, ".")
    If UBound(parts) <> 2 Then Exit Function
    For i = 0 To 2
        If Len(parts(i)) = 0 Or Len(parts(i)) > 6 Then Exit Function
        For k = 1 To Len(parts(i))
            If InStr("0123456789", Mid$(parts(i), k, 1)) = 0 Then Exit Function
        Next
    Next
    IsValidVersion = True
End Function

' 1 se a > b, -1 se a < b, 0 se iguais. Versão inválida vale 0.0.0.
Public Function CompareVersions(a As String, b As String) As Long
    Dim pa() As String, pb() As String, i As Long, na As Long, nb As Long
    pa = Split(IIf(IsValidVersion(a), a, "0.0.0"), ".")
    pb = Split(IIf(IsValidVersion(b), b, "0.0.0"), ".")
    For i = 0 To 2
        na = CLng(pa(i)): nb = CLng(pb(i))
        If na > nb Then CompareVersions = 1: Exit Function
        If na < nb Then CompareVersions = -1: Exit Function
    Next
    CompareVersions = 0
End Function
