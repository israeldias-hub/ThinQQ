Attribute VB_Name = "modPptTestRunner"
Option Explicit
' Corredor do projeto de teste do PowerPoint: roda modTestDispatch.AllTests() (lista gerada pelo build) e grava TAP. Filtro: TQQ_TEST_FILTER.
Private mPassed As Long, mFailed As Long, mLog As String

Public Function RunAllPptTests(outputPath As String) As Long
    Dim names() As String, i As Long, filt As String
    filt = Environ$("TQQ_TEST_FILTER")
    mPassed = 0: mFailed = 0: mLog = ""
    names = Split(modTestDispatch.AllTests(), "|")
    For i = LBound(names) To UBound(names)
        If Len(names(i)) > 0 Then
            If Len(filt) = 0 Or InStr(1, names(i), filt, vbTextCompare) > 0 Then RunOne names(i)
        End If
    Next
    mLog = mLog & "1.." & (mPassed + mFailed) & vbLf & "# passed " & mPassed & " failed " & mFailed & vbLf
    WriteText outputPath, mLog
    RunAllPptTests = mFailed
End Function

Private Sub RunOne(qualified As String)
    On Error GoTo Fail
    modTestDispatch.Dispatch qualified
    mPassed = mPassed + 1
    mLog = mLog & "ok " & (mPassed + mFailed) & " - " & qualified & vbLf
    Exit Sub
Fail:
    mFailed = mFailed + 1
    mLog = mLog & "not ok " & (mPassed + mFailed) & " - " & qualified & " # " & Replace(Err.Description, vbLf, " ") & vbLf
End Sub

Private Sub WriteText(p As String, s As String)
    Dim f As Integer
    f = FreeFile
    Open p For Output As #f
    Print #f, s;
    Close #f
End Sub
