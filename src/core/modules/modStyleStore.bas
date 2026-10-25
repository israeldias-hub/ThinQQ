Attribute VB_Name = "modStyleStore"
Option Explicit
' Depositos dos estilos (1.12.0; spec 2026-10-23 secao 8.5): embutidos (StyleDef_Builtin), do usuario (%LOCALAPPDATA%\ThinQQ\styles\<nome>
' .tqqestilo, UTF-8) e da pasta (CustomXMLParts no namespace STYLESTORE_NS: quem recebe a pasta ve o mesmo visual; medido PX1c). No build Test o
' deposito do usuario so existe com o gancho StyleStore_SetStoreDirForTest (nunca a pasta real) e o padrao para novos graficos fica em memoria.
Public Const STYLESTORE_NS As String = "urn:thinqq:styles:1"
Private Const STYLESTORE_SECTION As String = "Estilos"
Private mTestDir As String
Private mTestDefault As String

Public Sub StyleStore_SetStoreDirForTest(dirPath As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestDir = dirPath
End Sub

Public Sub StyleStore_ResetTestHooks()
    mTestDir = ""
    mTestDefault = ""
End Sub

Public Function StyleStore_UserDir() As String
    If BUILD_CONFIG = "Test" Then
        StyleStore_UserDir = mTestDir
        Exit Function
    End If
    StyleStore_UserDir = Environ$("LOCALAPPDATA") & "\ThinQQ\styles"
End Function

' Cria a pasta do deposito do usuario se faltar (ThinQQ\styles). "" ok ou a mensagem. Nunca propaga.
Private Function StyleStore_EnsureDir() As String
    Dim d As String, parent As String
    On Error GoTo Fail
    d = StyleStore_UserDir()
    If Len(d) = 0 Then
        StyleStore_EnsureDir = "O depósito de estilos do usuário não está disponível."
        Exit Function
    End If
    If Len(Dir(d, vbDirectory)) > 0 Then Exit Function
    parent = Left$(d, InStrRev(d, "\") - 1)
    If Len(Dir(parent, vbDirectory)) = 0 Then MkDir parent
    MkDir d
    Exit Function
Fail:
    StyleStore_EnsureDir = "Não foi possível criar a pasta de estilos: " & Err.Description
End Function

Public Function StyleStore_ReadFile(path As String, ByRef msg As String) As String
    Dim stm As Object
    On Error GoTo Fail
    msg = ""
    If FileLen(path) > STYLEDEF_MAX_BYTES * 3 Then
        msg = "O arquivo é grande demais para um estilo do ThinQQ."
        Exit Function
    End If
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    StyleStore_ReadFile = stm.ReadText
    stm.Close
    If Left$(StyleStore_ReadFile, 1) = ChrW(65279) Then StyleStore_ReadFile = Mid$(StyleStore_ReadFile, 2)   ' BOM
    Exit Function
Fail:
    msg = "Não foi possível ler o arquivo: " & Err.Description
End Function

Public Function StyleStore_WriteFile(path As String, text As String) As String
    Dim stm As Object
    On Error GoTo Fail
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2                                          ' adSaveCreateOverWrite
    stm.Close
    Exit Function
Fail:
    StyleStore_WriteFile = "Não foi possível gravar o arquivo: " & Err.Description
End Function

Private Function StyleStore_UserPath(name As String) As String
    If Len(StyleStore_UserDir()) = 0 Or Len(StyleDef_SafeName(name)) = 0 Then Exit Function
    StyleStore_UserPath = StyleStore_UserDir() & "\" & StyleDef_SafeName(name) & STYLEDEF_EXT
End Function

Public Function StyleStore_UserText(name As String) As String
    Dim p As String, msg As String
    On Error Resume Next
    p = StyleStore_UserPath(name)
    If Len(p) = 0 Then Exit Function
    If Len(Dir(p)) = 0 Then Exit Function
    StyleStore_UserText = StyleStore_ReadFile(p, msg)
    Err.Clear
End Function

Public Function StyleStore_UserNames() As Collection
    Dim out As New Collection, f As String, d As String
    Set StyleStore_UserNames = out
    On Error Resume Next
    d = StyleStore_UserDir()
    If Len(d) = 0 Then Exit Function
    f = Dir(d & "\*" & STYLEDEF_EXT)
    Do While Len(f) > 0
        out.Add Left$(f, Len(f) - Len(STYLEDEF_EXT))
        f = Dir()
    Loop
    Err.Clear
End Function

' Grava o estilo do usuario (texto canonico, validado). Nome de embutido e recusado. "" ok ou "!" + mensagem.
Public Function StyleStore_SaveUser(name As String, text As String) As String
    Dim d As Object, ign As Long, msg As String, nm As String
    nm = StyleDef_SafeName(name)
    If Len(nm) = 0 Then
        StyleStore_SaveUser = "!Dê um nome ao estilo."
        Exit Function
    End If
    If StyleDef_IsBuiltin(nm) Then
        StyleStore_SaveUser = "!'" & nm & "' é um estilo embutido do ThinQQ: use outro nome (Duplicar)."
        Exit Function
    End If
    Set d = StyleDef_Parse(text, ign)
    If d Is Nothing Then
        StyleStore_SaveUser = "!O estilo não é válido."
        Exit Function
    End If
    d("nome") = nm
    msg = StyleStore_EnsureDir()
    If Len(msg) > 0 Then
        StyleStore_SaveUser = "!" & msg
        Exit Function
    End If
    msg = StyleStore_WriteFile(StyleStore_UserPath(nm), StyleDef_Serialize(d))
    If Len(msg) > 0 Then StyleStore_SaveUser = "!" & msg
End Function

Public Function StyleStore_DeleteUser(name As String) As String
    Dim p As String
    On Error GoTo Fail
    If StyleDef_IsBuiltin(name) Then
        StyleStore_DeleteUser = "!Estilos embutidos não podem ser excluídos."
        Exit Function
    End If
    p = StyleStore_UserPath(name)
    If Len(p) > 0 Then
        If Len(Dir(p)) > 0 Then Kill p
    End If
    Exit Function
Fail:
    StyleStore_DeleteUser = "!Não foi possível excluir: " & Err.Description
End Function

Private Function StyleStore_Esc(s As String) As String
    StyleStore_Esc = Replace(Replace(Replace(Replace(s, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), """", "&quot;")
End Function

' Estilos da pasta: Dictionary (texto) nome -> texto, lidos das partes no namespace (filhos <style name="...">texto</style>). Nunca propaga.
Public Function StyleStore_BookAll(wb As Workbook) As Object
    Dim d As Object, p As Object, nd As Object, a As Object, nm As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set StyleStore_BookAll = d
    On Error Resume Next
    For Each p In wb.CustomXMLParts.SelectByNamespace(STYLESTORE_NS)
        For Each nd In p.DocumentElement.ChildNodes
            If nd.BaseName = "style" Then
                nm = ""
                For Each a In nd.Attributes
                    If a.BaseName = "name" Then nm = StyleDef_SafeName(CStr(a.NodeValue))
                Next
                If Len(nm) > 0 And Len(nd.text) <= STYLEDEF_MAX_BYTES Then d(nm) = CStr(nd.text)
            End If
        Next
    Next
    Err.Clear
End Function

Public Function StyleStore_BookText(wb As Workbook, name As String) As String
    Dim d As Object
    Set d = StyleStore_BookAll(wb)
    If d.Exists(name) Then StyleStore_BookText = CStr(d(name))
End Function

' Grava/atualiza o estilo na pasta (uma unica parte no namespace, reescrita). Propaga erro (quem chama trata).
Public Sub StyleStore_SaveBook(wb As Workbook, name As String, text As String)
    Dim all As Object, k As Variant, xml As String, p As Object, old As New Collection
    Set all = StyleStore_BookAll(wb)
    all(StyleDef_SafeName(name)) = text
    xml = "<tqqStyles xmlns=""" & STYLESTORE_NS & """>"
    For Each k In all.Keys
        xml = xml & "<style name=""" & StyleStore_Esc(CStr(k)) & """>" & StyleStore_Esc(CStr(all(k))) & "</style>"
    Next
    xml = xml & "</tqqStyles>"
    For Each p In wb.CustomXMLParts.SelectByNamespace(STYLESTORE_NS)
        old.Add p
    Next
    For Each p In old
        p.Delete
    Next
    wb.CustomXMLParts.Add xml
End Sub

' A pasta passa a levar o estilo (so grava se o texto mudou). Nunca propaga (registrado).
Public Sub StyleStore_EnsureInBook(wb As Workbook, name As String, text As String)
    On Error GoTo Fail
    If Replace(StyleStore_BookText(wb, name), vbCrLf, vbLf) = Replace(text, vbCrLf, vbLf) Then Exit Sub   ' o XML devolve LF: nada a regravar
    StyleStore_SaveBook wb, name, text
    Exit Sub
Fail:
    Errors_Log "StyleStore_EnsureInBook" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Texto do estilo: preferUser (escolha explicita do usuario) = usuario, embutido, pasta; senao (reaplicar) = pasta, usuario, embutido.
Public Function StyleStore_Get(name As String, Optional wb As Workbook, Optional preferUser As Boolean = False) As String
    Dim t As String
    If preferUser Then
        t = StyleStore_UserText(name)
        If Len(t) = 0 Then t = StyleDef_Builtin(name)
        If Len(t) = 0 And Not wb Is Nothing Then t = StyleStore_BookText(wb, name)
    Else
        If Not wb Is Nothing Then t = StyleStore_BookText(wb, name)
        If Len(t) = 0 Then t = StyleStore_UserText(name)
        If Len(t) = 0 Then t = StyleDef_Builtin(name)
    End If
    StyleStore_Get = t
End Function

' Nomes: embutidos, do usuario (ordem do disco) e os so da pasta; sem repeticao; "|". Nunca propaga.
Public Function StyleStore_Names(Optional wb As Workbook) As String
    Dim seen As Object, s As String, n As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    On Error Resume Next
    For Each n In StyleDef_BuiltinNames()
        If Not seen.Exists(CStr(n)) Then
            seen.Add CStr(n), True
            s = s & IIf(Len(s) > 0, "|", "") & CStr(n)
        End If
    Next
    For Each n In StyleStore_UserNames()
        If Not seen.Exists(CStr(n)) Then
            seen.Add CStr(n), True
            s = s & IIf(Len(s) > 0, "|", "") & CStr(n)
        End If
    Next
    If Not wb Is Nothing Then
        For Each n In StyleStore_BookAll(wb).Keys
            If Not seen.Exists(CStr(n)) Then
                seen.Add CStr(n), True
                s = s & IIf(Len(s) > 0, "|", "") & CStr(n)
            End If
        Next
    End If
    StyleStore_Names = s
    Err.Clear
End Function

' Importa um arquivo de estilo (validado): grava no deposito do usuario com o nome do arquivo quando o "nome=" nao serve. Devolve o texto
' canonico gravado ou "!" + mensagem; name = nome gravado; ignored = linhas ignoradas.
Public Function StyleStore_Import(path As String, ByRef name As String, ByRef ignored As Long) As String
    Dim t As String, msg As String, d As Object, base As String
    t = StyleStore_ReadFile(path, msg)
    If Len(msg) > 0 Then
        StyleStore_Import = "!" & msg
        Exit Function
    End If
    Set d = StyleDef_Parse(t, ignored)
    If d Is Nothing Then
        StyleStore_Import = "!O arquivo não é um estilo do ThinQQ válido (falta a linha " & STYLEDEF_HEADER & ", ou é grande demais)."
        Exit Function
    End If
    name = StyleDef_Get(d, "nome")
    If Len(name) = 0 Then
        base = Mid$(path, InStrRev(path, "\") + 1)
        If LCase$(Right$(base, Len(STYLEDEF_EXT))) = STYLEDEF_EXT Then base = Left$(base, Len(base) - Len(STYLEDEF_EXT))
        name = StyleDef_SafeName(base)
    End If
    If StyleDef_IsBuiltin(name) Then name = name & " (importado)"
    d("nome") = name
    t = StyleDef_Serialize(d)
    msg = StyleStore_SaveUser(name, t)
    If Len(msg) > 0 Then
        StyleStore_Import = msg
        Exit Function
    End If
    StyleStore_Import = t
End Function

Public Function StyleStore_Export(name As String, path As String, Optional wb As Workbook) As String
    Dim t As String, d As Object, ign As Long
    t = StyleStore_Get(name, wb, True)
    Set d = StyleDef_Parse(t, ign)
    If d Is Nothing Then
        StyleStore_Export = "!Estilo não encontrado: " & name
        Exit Function
    End If
    t = StyleStore_WriteFile(path, StyleDef_Serialize(d))
    If Len(t) > 0 Then StyleStore_Export = "!" & t
End Function

Public Function StyleStore_DefaultName() As String
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then
        StyleStore_DefaultName = mTestDefault
        Exit Function
    End If
    StyleStore_DefaultName = GetSetting(APP_NAME, STYLESTORE_SECTION, "Padrao", "")
End Function

Public Sub StyleStore_SetDefaultName(n As String)
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then
        mTestDefault = n
        Exit Sub
    End If
    SaveSetting APP_NAME, STYLESTORE_SECTION, "Padrao", n
End Sub
