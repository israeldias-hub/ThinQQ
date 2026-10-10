Attribute VB_Name = "modTestRunner"
Option Explicit
' Executa todo Sub público "Test_*" de todo módulo "Test*" e grava saída no formato TAP.
Private mPassed As Long, mFailed As Long, mLog As String

Public Function RunAllTests(outputPath As String) As Long
    Dim comp As Object, cm As Object, line As Long, procName As String, kind As Long, lastProc As String
    mPassed = 0: mFailed = 0: mLog = ""
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 1 And Left$(comp.Name, 4) = "Test" Then    ' 1 = vbext_ct_StdModule
            Set cm = comp.CodeModule
            lastProc = ""
            For line = cm.CountOfDeclarationLines + 1 To cm.CountOfLines
                procName = cm.ProcOfLine(line, kind)
                If procName <> lastProc Then
                    lastProc = procName
                    If Left$(procName, 5) = "Test_" Then RunOne comp.Name, procName
                End If
            Next
        End If
    Next
    mLog = mLog & "1.." & (mPassed + mFailed) & vbLf & "# passed " & mPassed & " failed " & mFailed & vbLf
    WriteText outputPath, mLog
    RunAllTests = mFailed
End Function

Private Sub RunOne(moduleName As String, procName As String)
    On Error GoTo Fail
    Application.Run "'" & ThisWorkbook.Name & "'!" & moduleName & "." & procName
    mPassed = mPassed + 1
    mLog = mLog & "ok " & (mPassed + mFailed) & " - " & moduleName & "." & procName & vbLf
    Exit Sub
Fail:
    mFailed = mFailed + 1
    mLog = mLog & "not ok " & (mPassed + mFailed) & " - " & moduleName & "." & procName & " # " & Replace(Replace(Err.Description, vbCr, " "), vbLf, " ") & vbLf
End Sub

Private Sub WriteText(path As String, text As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, text
    Close #f
End Sub
