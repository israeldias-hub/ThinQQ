Option Explicit
Public Result As String

Public Sub Prepare(title As String, versionLine As String, notes As String, Optional showSkip As Boolean = True, Optional showPpam As Boolean = False)
    lblTitle.Caption = title
    lblVersion.Caption = versionLine
    txtNotes.Text = notes
    btnSkip.Visible = showSkip
    chkPpam.Visible = showPpam
    chkPpam.Value = (Settings_Get(KEY_PPAM_AUTO, "1") <> "0")
    Result = "later"
End Sub

Private Sub btnUpdate_Click()
    Result = "update"
    If chkPpam.Visible Then Settings_Set KEY_PPAM_AUTO, IIf(chkPpam.Value, "1", "0")
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
