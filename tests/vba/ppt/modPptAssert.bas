Attribute VB_Name = "modPptAssert"
Option Explicit
' Asserts do projeto de teste do PowerPoint (o modAssert comum depende do Excel).
Public Sub AssertTrue(cond As Boolean, Optional msg As String = "esperado True")
    If Not cond Then Err.Raise vbObjectError + 100, "Assert", msg
End Sub
Public Sub AssertFalse(cond As Boolean, Optional msg As String = "esperado False")
    If cond Then Err.Raise vbObjectError + 100, "Assert", msg
End Sub
Public Sub AssertEqual(expected As Variant, actual As Variant, Optional msg As String = "")
    If expected <> actual Then Err.Raise vbObjectError + 100, "Assert", "esperado [" & expected & "] obtido [" & actual & "] " & msg
End Sub
' Raiz do repositorio (o .pptm de teste fica em <repo>\dist): usada pelo modFormCheck comum (no Excel, a do modAssert).
Public Function RepoPath() As String
    Dim i As Long
    For i = 1 To Application.Presentations.Count
        If LCase$(Application.Presentations(i).Name) Like "thinqq.ppt.test*.pptm" Then
            RepoPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(Application.Presentations(i).Path)
            Exit Function
        End If
    Next
End Function
