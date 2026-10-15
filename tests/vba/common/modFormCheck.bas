Attribute VB_Name = "modFormCheck"
Option Explicit
' Confere que cada UserForm do projeto (src\<projeto>\forms\*.json) sai do build com a legenda do JSON como TITULO DA JANELA
' (propriedade Caption do proprio formulario). Nenhum formulario e exibido: so UserForms.Add, leitura de Caption e Unload.

' Texto UTF-8 de um arquivo.
Private Function ReadUtf8(path As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    ReadUtf8 = stm.ReadText
    stm.Close
End Function

' Valor de um campo de texto de primeiro nivel do JSON do formulario ("name" ou "caption"; ocorrencia mais antiga do texto).
Private Function JsonText(json As String, field As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & field & """\s*:\s*""([^""]*)"""
    Set m = re.Execute(json)
    If m.Count > 0 Then JsonText = m(0).SubMatches(0)
End Function

' Percorre src\<projeto>\forms\*.json, cria cada formulario em memoria e compara f.Caption com o "caption" do JSON.
' Devolve quantos formularios foram conferidos (o chamador exige pelo menos o esperado).
Public Function CheckFormCaptions(project As String) As Long
    Dim fso As Object, fld As String, f As Object, json As String, nm As String, cap As String, frm As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fld = RepoPath() & "\src\" & project & "\forms"
    AssertTrue fso.FolderExists(fld), "pasta de formularios nao encontrada: " & fld
    For Each f In fso.GetFolder(fld).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "json" Then
            json = ReadUtf8(f.Path)
            nm = JsonText(json, "name")
            cap = JsonText(json, "caption")
            AssertTrue Len(nm) > 0 And Len(cap) > 0, f.Name & ": JSON sem name/caption"
            Set frm = VBA.UserForms.Add(nm)
            AssertEqual cap, frm.Caption, nm & ": titulo da janela diferente do caption do JSON (Caption do formulario)"
            Unload frm
            CheckFormCaptions = CheckFormCaptions + 1
        End If
    Next
End Function
