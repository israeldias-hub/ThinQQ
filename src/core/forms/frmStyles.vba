Option Explicit
' Estilos de grafico (1.12.0): lista, edicao por papel (valor digitado ou escolhido; validado por StyleDef_Validate), salvar no deposito do
' usuario, importar/exportar, aplicar e padrao para novos. Logica pura em modStyleDef/modStyleStore/modStyleUi; aqui so a tela.
Private mEdit As Object                 ' Dictionary do estilo em edicao
Private mKeys As Variant
Private mLoading As Boolean
Private mShown As String                ' estilo mostrado (o item da lista que esta em edicao)
Private mDirty As Boolean               ' valores mudados com Usar valor e ainda nao salvos
Private mReverting As Boolean           ' voltando a selecao da lista (o usuario nao quis descartar)

Public Sub Prepare()
    mLoading = True
    mKeys = StyleDef_Keys()
    cboScope.Clear
    cboScope.AddItem "Gráfico selecionado"
    cboScope.AddItem "Planilha"
    cboScope.AddItem "Pasta"
    cboScope.ListIndex = 0
    LoadNames Core_GetStyleName()
    mLoading = False
End Sub

Private Sub LoadNames(selectName As String)
    Dim n As Variant, i As Long
    lstStyles.Clear
    For Each n In Split(StyleStore_Names(ActiveWorkbook), "|")
        lstStyles.AddItem CStr(n)
        If StrComp(CStr(n), selectName, vbTextCompare) = 0 Then i = lstStyles.ListCount - 1
    Next
    If lstStyles.ListCount > 0 Then lstStyles.ListIndex = i
End Sub

Private Sub ShowStyle(name As String)
    Dim ign As Long, k As Variant, i As Long, wasLoading As Boolean
    wasLoading = mLoading
    mLoading = True
    Set mEdit = StyleDef_Parse(StyleStore_Get(name, ActiveWorkbook, True), ign)
    If mEdit Is Nothing Then Set mEdit = StyleDef_Parse(STYLEDEF_HEADER, ign)
    txtName.Text = name
    chkDefault.Value = (StrComp(StyleStore_DefaultName(), name, vbTextCompare) = 0)
    lstRoles.Clear
    For Each k In mKeys
        If CStr(k) <> "nome" Then
            lstRoles.AddItem StyleUi_RoleLabel(CStr(k))
            i = lstRoles.ListCount - 1
            lstRoles.List(i, 1) = StyleDef_Get(mEdit, CStr(k))
        End If
    Next
    btnDelete.Enabled = Not StyleDef_IsBuiltin(name)
    mShown = name
    mDirty = False
    mLoading = wasLoading
End Sub

' True se o usuario ja tem um estilo com esse nome que NAO e o mostrado: gravar por cima pergunta antes (revisao final 1.12.0, M2).
Private Function ConfirmOverwrite(nm As String) As Boolean
    ConfirmOverwrite = True
    If Len(StyleStore_UserText(nm)) = 0 Then Exit Function
    If StrComp(nm, mShown, vbTextCompare) = 0 Then Exit Function
    ConfirmOverwrite = Core_Confirm("Já existe um estilo seu chamado '" & nm & "'. Substituir?")
End Function

Private Function ListIndexOf(nm As String) As Long
    Dim i As Long
    ListIndexOf = -1
    For i = 0 To lstStyles.ListCount - 1
        If StrComp(CStr(lstStyles.List(i)), nm, vbTextCompare) = 0 Then
            ListIndexOf = i
            Exit Function
        End If
    Next
End Function

Private Function CurrentKey() As String
    If lstRoles.ListIndex < 0 Then Exit Function
    CurrentKey = CStr(mKeys(lstRoles.ListIndex + 1))          ' mKeys(0) = "nome" (fora da lista)
End Function

Private Sub lstStyles_Click()
    If mReverting Then Exit Sub
    If lstStyles.ListIndex < 0 Then Exit Sub
    If mDirty And Not mLoading Then
        If StrComp(CStr(lstStyles.List(lstStyles.ListIndex)), mShown, vbTextCompare) <> 0 Then
            If Not Core_Confirm("Descartar as mudanças não salvas em '" & mShown & "'?") Then
                mReverting = True
                lstStyles.ListIndex = ListIndexOf(mShown)
                mReverting = False
                Exit Sub
            End If
        End If
    End If
    ShowStyle CStr(lstStyles.List(lstStyles.ListIndex))
End Sub

Private Sub lstRoles_Click()
    Dim k As String, items As Variant, i As Long
    k = CurrentKey()
    If Len(k) = 0 Then Exit Sub
    txtValue.Text = StyleDef_Get(mEdit, k)
    cboPick.Clear
    items = StyleUi_PickItems(k)
    If IsArray(items) Then
        For i = LBound(items) To UBound(items)
            cboPick.AddItem CStr(items(i))
        Next
    End If
    Swatch
End Sub

Private Sub cboPick_Change()
    If cboPick.ListIndex >= 0 Then txtValue.Text = StyleUi_PickValue(CurrentKey(), cboPick.ListIndex)
    Swatch
End Sub

Private Sub Swatch()
    Dim c As Long
    c = StyleDef_ColorOf(txtValue.Text)
    If c >= 0 Then lblSwatch.BackColor = c Else lblSwatch.BackColor = RGB(240, 240, 240)
End Sub

Private Sub txtValue_Change()
    Swatch
End Sub

Private Sub btnSetValue_Click()
    Dim k As String, v As String
    k = CurrentKey()
    If Len(k) = 0 Then
        lblStatus.Caption = "Escolha um papel na lista."
        Exit Sub
    End If
    If Len(Trim$(txtValue.Text)) = 0 Then
        If mEdit.Exists(k) Then mEdit.Remove k
        lstRoles.List(lstRoles.ListIndex, 1) = ""
        mDirty = True
        lblStatus.Caption = "Valor removido (o ThinQQ usa o padrão)."
        Exit Sub
    End If
    v = StyleDef_Validate(k, txtValue.Text)
    If Len(v) = 0 Then
        lblStatus.Caption = "Valor inválido para '" & StyleUi_RoleLabel(k) & "'."
        Exit Sub
    End If
    mEdit(k) = v
    lstRoles.List(lstRoles.ListIndex, 1) = v
    mDirty = True
    lblStatus.Caption = ""
End Sub

Private Sub btnSave_Click()
    Dim r As String
    If Not ConfirmOverwrite(StyleDef_SafeName(txtName.Text)) Then Exit Sub
    r = StyleStore_SaveUser(txtName.Text, StyleDef_Serialize(mEdit))
    If Left$(r, 1) = "!" Then
        lblStatus.Caption = Mid$(r, 2)
        Exit Sub
    End If
    If chkDefault.Value Then StyleStore_SetDefaultName StyleDef_SafeName(txtName.Text)
    mDirty = False
    lblStatus.Caption = "Estilo salvo: " & StyleDef_SafeName(txtName.Text) & "."
    LoadNames StyleDef_SafeName(txtName.Text)
End Sub

Private Sub btnDuplicate_Click()
    txtName.Text = txtName.Text & " (cópia)"
    btnDelete.Enabled = True
    lblStatus.Caption = "Mude o nome se quiser e clique em Salvar."
End Sub

Private Sub btnNewFromChart_Click()
    Dim r As String, nm As String
    nm = StyleDef_SafeName(txtName.Text)
    If Len(nm) = 0 Or StyleDef_IsBuiltin(nm) Then nm = "Meu estilo"
    If Len(StyleStore_UserText(nm)) > 0 Then
        If Not Core_Confirm("Já existe um estilo seu chamado '" & nm & "'. Substituir?") Then Exit Sub
    End If
    mDirty = False
    r = Core_StyleFromChart(nm)
    If Len(r) > 0 Then
        lblStatus.Caption = r
        Exit Sub
    End If
    lblStatus.Caption = "Estilo criado a partir do gráfico selecionado: " & nm & "."
    LoadNames nm
End Sub

' Exclui o estilo SELECIONADO na lista (nao o texto digitado em Nome), depois de perguntar. Nunca apaga embutidos nem o que nao existe.
Private Sub btnDelete_Click()
    Dim r As String, nm As String
    If lstStyles.ListIndex < 0 Then Exit Sub
    nm = CStr(lstStyles.List(lstStyles.ListIndex))
    If Len(StyleStore_UserText(nm)) = 0 Then
        lblStatus.Caption = IIf(StyleDef_IsBuiltin(nm), "Estilos embutidos não podem ser excluídos.", "'" & nm & "' não é um estilo seu (vem da pasta): nada a excluir.")
        Exit Sub
    End If
    If Not Core_Confirm("Excluir o estilo '" & nm & "'? Não dá para desfazer.") Then Exit Sub
    r = StyleStore_DeleteUser(nm)
    If Left$(r, 1) = "!" Then
        lblStatus.Caption = Mid$(r, 2)
        Exit Sub
    End If
    mDirty = False
    lblStatus.Caption = "Estilo excluído: " & nm & "."
    LoadNames ""
End Sub

Private Sub btnImport_Click()
    Core_StyleImport
    LoadNames txtName.Text
End Sub

Private Sub btnExport_Click()
    Core_StyleExport txtName.Text
End Sub

Private Sub btnApply_Click()
    Dim scope As String
    scope = Choose(cboScope.ListIndex + 1, "chart", "sheet", "book")
    Core_StyleApply StyleDef_SafeName(txtName.Text), scope
    If chkDefault.Value Then StyleStore_SetDefaultName StyleDef_SafeName(txtName.Text)
    lblStatus.Caption = "Aplicado."
End Sub

Private Sub chkDefault_Click()
    If mLoading Then Exit Sub
    If chkDefault.Value Then
        StyleStore_SetDefaultName StyleDef_SafeName(txtName.Text)
    ElseIf StrComp(StyleStore_DefaultName(), StyleDef_SafeName(txtName.Text), vbTextCompare) = 0 Then
        StyleStore_SetDefaultName ""
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
