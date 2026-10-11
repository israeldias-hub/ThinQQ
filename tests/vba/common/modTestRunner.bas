Attribute VB_Name = "modTestRunner"
Option Explicit
' Executa todo Sub público "Test_*" de todo módulo "Test*" e grava saída no formato TAP.
Private mPassed As Long, mFailed As Long, mLog As String

Public Function RunAllTests(outputPath As String) As Long
    Dim comp As Object, cm As Object, line As Long, procName As String, kind As Long, lastProc As String, filt As String
    filt = Environ$("TQQ_TEST_FILTER")
    mPassed = 0: mFailed = 0: mLog = ""
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 1 And Left$(comp.Name, 4) = "Test" Then    ' 1 = vbext_ct_StdModule
            Set cm = comp.CodeModule
            lastProc = ""
            For line = cm.CountOfDeclarationLines + 1 To cm.CountOfLines
                procName = cm.ProcOfLine(line, kind)
                If procName <> lastProc Then
                    lastProc = procName
                    If Left$(procName, 5) = "Test_" Then
                        If FilterMatches(filt, comp.Name & "." & procName) Then RunOne comp.Name, procName
                    End If
                End If
            Next
        End If
    Next
    mLog = mLog & "1.." & (mPassed + mFailed) & vbLf & "# passed " & mPassed & " failed " & mFailed & vbLf
    WriteText outputPath, mLog
    RunAllTests = mFailed
End Function

' Filtro de iteracao rapida (variavel TQQ_TEST_FILTER): trechos separados por "|"; vazio = todos.
Private Function FilterMatches(filt As String, fullName As String) As Boolean
    Dim parts() As String, i As Long
    If Len(filt) = 0 Then FilterMatches = True: Exit Function
    parts = Split(filt, "|")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            If InStr(1, fullName, parts(i), vbTextCompare) > 0 Then FilterMatches = True: Exit Function
        End If
    Next
End Function

Private Function OpenBooks() As String
    Dim w As Workbook, s As String
    For Each w In Application.Workbooks: s = s & "|" & w.Name: Next
    OpenBooks = s & "|"
End Function

' Fecha pastas abertas durante o teste (vazadas por falha de assert).
Private Sub CloseLeaked(baseline As String)
    Dim i As Long, w As Workbook, prev As Boolean
    prev = Application.DisplayAlerts: Application.DisplayAlerts = False
    On Error Resume Next
    For i = Application.Workbooks.Count To 1 Step -1
        Set w = Application.Workbooks(i)
        If InStr(baseline, "|" & w.Name & "|") = 0 Then w.Close SaveChanges:=False
    Next
    Application.DisplayAlerts = prev
End Sub

Private Sub RunOne(moduleName As String, procName As String)
    Dim baseline As String
    baseline = OpenBooks()
    RunOne2 moduleName, procName
    CloseLeaked baseline
    CleanupTempDirs
End Sub

Private Sub RunOne2(moduleName As String, procName As String)
    On Error GoTo Fail
    ' Application.Run abre dialogo modal em erro nao tratado no Excel invisivel (o On Error nao captura);
    ' por isso chama via modTestDispatch (gerado pelo build com chamadas diretas).
    modTestDispatch.Dispatch moduleName & "." & procName
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
