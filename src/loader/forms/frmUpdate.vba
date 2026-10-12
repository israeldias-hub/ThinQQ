Option Explicit
Public Result As String

Public Sub Prepare(title As String, versionLine As String, notes As String, Optional showSkip As Boolean = True)
    lblTitle.Caption = title
    lblVersion.Caption = versionLine
    txtNotes.Text = notes
    btnSkip.Visible = showSkip
    Result = "later"
End Sub

Private Sub btnUpdate_Click()
    Result = "update"
    Me.Hide
End Sub

Private Sub btnLater_Click()
    Result = "later"
    Me.Hide
End Sub

Private Sub btnSkip_Click()
    Result = "skip"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Result = "later"
        Me.Hide
    End If
End Sub
