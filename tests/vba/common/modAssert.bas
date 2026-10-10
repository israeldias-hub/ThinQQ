Attribute VB_Name = "modAssert"
Option Explicit
Public Const ASSERT_ERR As Long = vbObjectError + 513

Public Sub AssertTrue(cond As Boolean, Optional msg As String = "esperado True")
    If Not cond Then Err.Raise ASSERT_ERR, "Assert", msg
End Sub

Public Sub AssertFalse(cond As Boolean, Optional msg As String = "esperado False")
    If cond Then Err.Raise ASSERT_ERR, "Assert", msg
End Sub

Public Sub AssertEqual(expected As Variant, actual As Variant, Optional msg As String = "")
    Dim same As Boolean
    If IsNumeric(expected) And IsNumeric(actual) Then
        same = (CDbl(expected) = CDbl(actual))
    Else
        same = (CStr(expected) = CStr(actual))
    End If
    If Not same Then Err.Raise ASSERT_ERR, "Assert", msg & " esperado=<" & CStr(expected) & "> obtido=<" & CStr(actual) & ">"
End Sub

Public Sub AssertAlmost(expected As Double, actual As Double, Optional tol As Double = 0.000001, Optional msg As String = "")
    If Abs(expected - actual) > tol Then Err.Raise ASSERT_ERR, "Assert", msg & " esperado~" & expected & " obtido=" & actual
End Sub

Public Function NewTempWorkbook() As Workbook
    Set NewTempWorkbook = Application.Workbooks.Add
End Function

Public Sub CloseTemp(wb As Workbook)
    On Error Resume Next
    wb.Close SaveChanges:=False
End Sub

' Pasta temporária nova e vazia (o chamador apaga se quiser).
Public Function TempDir() As String
    Dim p As String
    Randomize
    p = Environ$("TEMP") & "\thinqq_test_" & Format$(Now, "hhnnss") & "_" & CLng(Rnd * 1000000)
    MkDir p
    TempDir = p
End Function

' Raiz do repositório (o .xlam de teste fica em <repo>\dist).
Public Function RepoPath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    RepoPath = fso.GetParentFolderName(ThisWorkbook.Path)
End Function
