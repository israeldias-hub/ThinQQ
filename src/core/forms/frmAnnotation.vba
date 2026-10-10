Option Explicit
Public Result As String
Public FromIdx As Long
Public ToIdx As Long

Public Sub Prepare(categories As Variant, defaultFrom As Long, defaultTo As Long)
    Dim i As Long
    cboFrom.Clear: cboTo.Clear
    For i = LBound(categories) To UBound(categories)
        cboFrom.AddItem CStr(categories(i))
        cboTo.AddItem CStr(categories(i))
    Next
    cboFrom.ListIndex = defaultFrom - 1
    cboTo.ListIndex = defaultTo - 1
    Result = "cancel"
End Sub

Private Sub btnOk_Click()
    If cboFrom.ListIndex < 0 Or cboTo.ListIndex < 0 Then Exit Sub
    If cboFrom.ListIndex >= cboTo.ListIndex Then
        MsgBox "A categoria final deve vir depois da inicial.", vbExclamation, "ThinQQ"
        Exit Sub
    End If
    FromIdx = cboFrom.ListIndex + 1
    ToIdx = cboTo.ListIndex + 1
    Result = "ok"
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Result = "cancel"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: Result = "cancel": Me.Hide
End Sub
