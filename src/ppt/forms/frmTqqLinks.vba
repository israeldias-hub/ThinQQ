Option Explicit
' Vinculos do ThinQQ (1.11.0; o Data Links do think-cell): formulario NAO modal. A logica e testavel em modPptLinks; aqui so a tela.
' Segue a apresentacao ativa e a selecao (PptLinks_OnSelChange -> SyncView): so etiquetas lidas em processo, nunca o Excel; a lista so e
' redesenhada quando algo visivel mudou (PptLinks_Signature: sem piscar). Todo evento contido (nunca propaga erro ao PowerPoint).
Private mPres As Object
Private mPresKey As String
Private mRows As Collection
Private mShown As Collection
Private mSig As String
Private mLoading As Boolean
Private mActing As Boolean

Public Property Get Pres() As Object
    Set Pres = mPres
End Property

Public Property Get PresKey() As String
    PresKey = mPresKey
End Property

Public Sub Prepare(p As Object)
    Dim avail As Boolean
    On Error Resume Next
    Set mPres = p
    mPresKey = PptLinks_PresKey(p)
    mSig = vbNullChar
    mLoading = True
    cboFilter.Clear
    cboFilter.AddItem "Todos"
    cboFilter.AddItem "Dados novos"
    cboFilter.AddItem "Atualizados"
    cboFilter.AddItem "Automáticos"
    cboFilter.AddItem "Com problema"
    cboFilter.ListIndex = 0
    cboOpen.Clear
    cboOpen.AddItem "Não perguntar"
    cboOpen.AddItem "Perguntar"
    cboOpen.AddItem "Atualizar sempre"
    cboOpen.ListIndex = PptLinks_OpenModeIndex(PptTqq_OpenMode())
    lblStatus.Caption = ""
    chkEngine.Value = PptTqq_CoreFlag("autoon", avail)
    chkEngine.Enabled = avail
    chkNewAuto.Value = PptTqq_CoreFlag("autonew", avail)
    chkNewAuto.Enabled = avail
    If Not avail Then lblStatus.Caption = "Abra o Excel com o ThinQQ para ver o estado e mudar as opções de acompanhamento."
    mLoading = False
    Reload
End Sub

' Relê as etiquetas e redesenha a lista so se algo visivel mudou (mantem as linhas selecionadas).
Public Sub Reload()
    On Error Resume Next
    If mPres Is Nothing Then Exit Sub
    Set mRows = PptLinks_Collect(mPres)
    If PptLinks_Signature(mRows) <> mSig Then Fill
End Sub

' Evento de selecao/slide/janela do PowerPoint: acompanha a apresentacao ativa, rele as etiquetas e marca na lista os elementos
' selecionados no slide. Nunca chama o Excel; ignorado durante uma acao do proprio dialogo.
Public Sub SyncView()
    Dim p As Object, keys As Object
    If mActing Or mLoading Then Exit Sub
    On Error Resume Next
    Set p = PptLinks_ActivePres()
    If p Is Nothing Then Exit Sub
    If PptLinks_PresKey(p) <> mPresKey Then
        Set mPres = p
        mPresKey = PptLinks_PresKey(p)
        mSig = vbNullChar
        lblStatus.Caption = ""
    End If
    Reload
    Set keys = PptLinks_SelectedKeys()
    If Not keys Is Nothing Then
        If keys.Count > 0 Then MarkKeys keys
    End If
End Sub

Private Sub Fill()
    Dim r As Variant, i As Long, keep As Object
    On Error Resume Next
    Set keep = ListedKeys(True)
    Set mShown = PptLinks_Filter(mRows, cboFilter.ListIndex)
    lstLinks.Clear
    For Each r In mShown
        lstLinks.AddItem CStr(r(0))
        i = lstLinks.ListCount - 1
        lstLinks.List(i, 1) = PptLinks_KindLabel(CStr(r(3))) & " " & CStr(r(2))
        lstLinks.List(i, 2) = PptLinks_ModeLabel(CStr(r(4)))
        lstLinks.List(i, 3) = PptLinks_StateLabel(CStr(r(5)))
        lstLinks.List(i, 4) = PptLinks_FileName(CStr(r(6)))
    Next
    If keep.Count > 0 Then MarkKeys keep
    lblHead.Caption = PptLinks_HeadText(mRows)
    mSig = PptLinks_Signature(mRows)
End Sub

' Chaves (SlideID|nome) das linhas listadas: so as selecionadas (onlySelected) ou todas.
Private Function ListedKeys(onlySelected As Boolean) As Object
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    Set ListedKeys = d
    On Error Resume Next
    If mShown Is Nothing Then Exit Function
    For i = 0 To lstLinks.ListCount - 1
        If i + 1 > mShown.Count Then Exit For
        If lstLinks.Selected(i) Or Not onlySelected Then d(PptLinks_RowKey(mShown(i + 1))) = True
    Next
End Function

' Seleciona na lista exatamente as linhas das chaves dadas (as outras saem da selecao).
Private Sub MarkKeys(keys As Object)
    Dim i As Long, want As Boolean
    On Error Resume Next
    If mShown Is Nothing Then Exit Sub
    For i = 0 To lstLinks.ListCount - 1
        If i + 1 > mShown.Count Then Exit For
        want = keys.Exists(PptLinks_RowKey(mShown(i + 1)))
        If lstLinks.Selected(i) <> want Then lstLinks.Selected(i) = want
    Next
End Sub

Private Function SelectedRows() As Collection
    Dim c As New Collection, i As Long
    Set SelectedRows = c
    On Error Resume Next
    If mShown Is Nothing Then Exit Function
    For i = 0 To lstLinks.ListCount - 1
        If i + 1 > mShown.Count Then Exit For
        If lstLinks.Selected(i) Then c.Add mShown(i + 1)
    Next
End Function

' Uma acao (resultado na linha de status, sem caixa quando deu certo). Reentrada ignorada.
Private Sub Act(what As String, rows As Collection)
    Dim msg As String
    If mActing Then Exit Sub
    mActing = True
    On Error Resume Next
    lblStatus.Caption = "Aguarde..."
    Me.Repaint
    msg = PptLinks_Do(mPres, what, rows)
    lblStatus.Caption = msg
    mActing = False
    Reload
End Sub

Private Sub btnUpdate_Click()
    Act "update", SelectedRows()
End Sub

Private Sub btnUpdateNew_Click()
    On Error Resume Next
    Reload
    Act "update", PptLinks_Filter(mRows, 1)
End Sub

Private Sub btnAuto_Click()
    Act "auto", SelectedRows()
End Sub

Private Sub btnManual_Click()
    Act "manual", SelectedRows()
End Sub

Private Sub btnGoto_Click()
    Act "goto", SelectedRows()
End Sub

Private Sub btnEdit_Click()
    Act "edit", SelectedRows()
End Sub

Private Sub btnCheck_Click()
    Act "check", mRows
End Sub

Private Sub btnUnlink_Click()
    Act "unlink", SelectedRows()
End Sub

Private Sub btnAll_Click()
    Dim i As Long
    On Error Resume Next
    For i = 0 To lstLinks.ListCount - 1
        lstLinks.Selected(i) = True
    Next
End Sub

Private Sub btnClose_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cboFilter_Change()
    On Error Resume Next
    If Not mLoading Then mSig = vbNullChar: Fill
End Sub

Private Sub cboOpen_Change()
    On Error Resume Next
    If Not mLoading Then
        PptTqq_SetOpenMode PptLinks_OpenModeOf(cboOpen.ListIndex)
        lblStatus.Caption = "Ao abrir a apresentação: " & cboOpen.Text & "."
    End If
End Sub

Private Sub chkEngine_Click()
    On Error Resume Next
    If Not mLoading Then lblStatus.Caption = PptTqq_SetCoreFlag("autoon", CBool(chkEngine.Value))
End Sub

Private Sub chkNewAuto_Click()
    On Error Resume Next
    If Not mLoading Then lblStatus.Caption = PptTqq_SetCoreFlag("autonew", CBool(chkNewAuto.Value))
End Sub

Private Sub lstLinks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Act "goto", SelectedRows()
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    PptLinks_FormClosed
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    PptLinks_FormClosed
    Set mPres = Nothing
End Sub
