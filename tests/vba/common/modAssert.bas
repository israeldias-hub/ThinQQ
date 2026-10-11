Attribute VB_Name = "modAssert"
Option Explicit
Public Const ASSERT_ERR As Long = vbObjectError + 513
Private mTempDirs As Collection   ' pastas criadas por TempDir; o runner apaga após cada teste

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

' Pasta temporaria nova e vazia (registrada; o modTestRunner apaga apos cada teste).
Public Function TempDir() As String
    Dim p As String
    Randomize
    p = Environ$("TEMP") & "\thinqq_test_" & Format$(Now, "hhnnss") & "_" & CLng(Rnd * 1000000)
    MkDir p
    If mTempDirs Is Nothing Then Set mTempDirs = New Collection
    mTempDirs.Add p
    TempDir = p
End Function

' Forca o desenho do grafico: o Excel so atualiza Left/Top/Width/Height de um DataLabel (e a area de plotagem que a seta de CAGR reduz)
' depois de RENDERIZAR o grafico; antes disso a leitura devolve os valores antigos (medido: iguais aos de dentro apos ligar os rotulos
' fora, ate um Export). Toda leitura de geometria de rotulo depois de mudar o layout passa por aqui.
Public Sub ForceLayout(cht As Chart)
    cht.Export TempDir() & "\layout.png"
End Sub

' Raiz do repositório (o .xlam de teste fica em <repo>\dist).
Public Function RepoPath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    RepoPath = fso.GetParentFolderName(ThisWorkbook.Path)
End Function

' Apaga as pastas criadas por TempDir (erros ignorados).
Public Sub CleanupTempDirs()
    Dim fso As Object, p As Variant
    If mTempDirs Is Nothing Then Exit Sub
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each p In mTempDirs
        If Left$(fso.GetFileName(CStr(p)), 12) = "thinqq_test_" Then fso.DeleteFolder CStr(p), True
    Next
    Set mTempDirs = New Collection
End Sub
