Attribute VB_Name = "modStyleStore"
Option Explicit
' Depositos dos estilos (1.12.0; spec 2026-10-23 secao 8.5): embutidos (StyleDef_Builtin), do usuario (%LOCALAPPDATA%\ThinQQ\styles\<nome>
' .tqqestilo, UTF-8) e da pasta (CustomXMLParts no namespace STYLESTORE_NS: quem recebe a pasta ve o mesmo visual; medido PX1c). No build Test o
' deposito do usuario so existe com o gancho StyleStore_SetStoreDirForTest (nunca a pasta real) e o padrao para novos graficos fica em memoria.
Public Const STYLESTORE_NS As String = "urn:thinqq:styles:1"
Private Const STYLESTORE_SECTION As String = "Estilos"
' Tetos da pasta (revisao final 1.12.0, I2: uma pasta recebida, hostil ou so grande, nunca para o Excel na reaplicacao automatica)
Public Const STYLESTORE_BOOK_MAX As Long = 100              ' estilos lidos/gravados por pasta
Public Const STYLESTORE_BOOK_MAXCHARS As Long = 1048576     ' caracteres somados dos textos da pasta (1 MB)
Public Const STYLESTORE_NAMES_MAX As Long = 200             ' nomes entregues a faixa (o dropDown do Office aceita ate 1000)
Public Const STYLESTORE_MSG_BOOK_FULL As String = "A pasta já tem estilos demais."
Private mTestDir As String
Private mTestDefault As String
Private mBookCapLogged As Boolean

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
    If Mid$(Environ$("LOCALAPPDATA"), 2, 1) <> ":" Then Exit Function      ' sem LOCALAPPDATA (ou relativo): nunca a raiz da unidade
    StyleStore_UserDir = Environ$("LOCALAPPDATA") & "\ThinQQ\styles"
End Function

' Nome de dispositivo do Windows (CON, PRN, AUX, NUL, COM1..COM9, LPT1..LPT9): "<nome>.tqqestilo" nao e um arquivo comum. Pura.
Public Function StyleStore_IsDeviceName(name As String) As Boolean
    Dim u As String
    u = UCase$(Trim$(name))
    Select Case u
        Case "CON", "PRN", "AUX", "NUL"
            StyleStore_IsDeviceName = True
        Case Else
            If Len(u) = 4 Then
                If (Left$(u, 3) = "COM" Or Left$(u, 3) = "LPT") And Right$(u, 1) Like "[0-9]" Then StyleStore_IsDeviceName = True
            End If
    End Select
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

' Grava em <path>.tmp e so entao troca pelo arquivo (uma queda no meio deixa o arquivo antigo inteiro, nunca um meio-arquivo que o parser
' recusaria). "" ok ou a mensagem. Nunca propaga.
Public Function StyleStore_WriteFile(path As String, text As String) As String
    Dim stm As Object, tmp As String
    On Error GoTo Fail
    tmp = path & ".tmp"
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile tmp, 2                                           ' adSaveCreateOverWrite
    stm.Close
    Set stm = Nothing
    If Len(Dir(path)) > 0 Then Kill path
    Name tmp As path
    Exit Function
Fail:
    StyleStore_WriteFile = "Não foi possível gravar o arquivo: " & Err.Description
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
    If Len(tmp) > 0 Then
        If Len(Dir(tmp)) > 0 Then Kill tmp
    End If
End Function

Private Function StyleStore_UserPath(name As String) As String
    If Len(StyleStore_UserDir()) = 0 Or Len(StyleDef_SafeName(name)) = 0 Then Exit Function
    If StyleStore_IsDeviceName(StyleDef_SafeName(name)) Then Exit Function
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
    Dim out As New Collection, f As String, d As String, n As String
    Set StyleStore_UserNames = out
    On Error Resume Next
    d = StyleStore_UserDir()
    If Len(d) = 0 Then Exit Function
    f = Dir(d & "\*" & STYLEDEF_EXT)
    Do While Len(f) > 0
        n = Left$(f, Len(f) - Len(STYLEDEF_EXT))
        ' so nomes que o ThinQQ gravaria (arquivo posto a mao com ; = & ou com nome de embutido nao entra na lista)
        If StyleDef_SafeName(n) = n And Not StyleDef_IsBuiltin(n) And Not StyleStore_IsDeviceName(n) And Len(n) > 0 Then out.Add n
        If out.Count >= STYLESTORE_NAMES_MAX Then Exit Do
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
    If StyleStore_IsDeviceName(nm) Then
        StyleStore_SaveUser = "!'" & nm & "' é um nome reservado do Windows: use outro nome."
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

' Estilos da pasta: Dictionary (texto) nome -> texto, lidos das partes no namespace (filhos <style name="...">texto</style>). Para de ler no
' teto (STYLESTORE_BOOK_MAX estilos ou STYLESTORE_BOOK_MAXCHARS caracteres; registrado uma vez por sessao): uma pasta com milhares de estilos
' nunca para o Excel. Nunca propaga.
Public Function StyleStore_BookAll(wb As Workbook) As Object
    Dim d As Object, p As Object, nd As Object, a As Object, nm As String, t As String, total As Long, nodes As Long, capped As Boolean, parts As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set StyleStore_BookAll = d
    On Error Resume Next
    For Each p In wb.CustomXMLParts.SelectByNamespace(STYLESTORE_NS)
        parts = parts + 1
        If parts > STYLESTORE_BOOK_MAX Then capped = True: Exit For
        For Each nd In p.DocumentElement.ChildNodes
            nodes = nodes + 1
            If d.Count >= STYLESTORE_BOOK_MAX Or nodes > 4 * STYLESTORE_BOOK_MAX Then capped = True: Exit For
            If nd.BaseName = "style" Then
                nm = ""
                For Each a In nd.Attributes
                    If a.BaseName = "name" Then nm = StyleDef_SafeName(CStr(a.NodeValue))
                Next
                t = ""
                t = CStr(nd.text)
                If Len(nm) > 0 And Len(t) <= STYLEDEF_MAX_BYTES Then
                    If total + Len(t) > STYLESTORE_BOOK_MAXCHARS Then capped = True: Exit For
                    total = total + Len(t)
                    d(nm) = t
                End If
            End If
            Err.Clear
        Next
        If capped Then Exit For
    Next
    If capped And Not mBookCapLogged Then
        mBookCapLogged = True
        Errors_Log "StyleStore_BookAll: pasta com estilos demais (lidos so os primeiros " & d.Count & ")" & vbTab & wb.name
    End If
    Err.Clear
End Function

Public Function StyleStore_BookText(wb As Workbook, name As String) As String
    Dim d As Object
    Set d = StyleStore_BookAll(wb)
    If d.Exists(name) Then StyleStore_BookText = CStr(d(name))
End Function

' Grava/atualiza o estilo na pasta (uma unica parte no namespace, reescrita). Propaga erro (quem chama trata).
' Recusa (erro vbObjectError + 7140, STYLESTORE_MSG_BOOK_FULL) quando passaria do teto da pasta; o XML e montado por Join (linear).
Public Sub StyleStore_SaveBook(wb As Workbook, name As String, text As String)
    Dim all As Object, k As Variant, xml As String, p As Object, old As New Collection, parts() As String, i As Long, total As Long
    Set all = StyleStore_BookAll(wb)
    If Not all.Exists(StyleDef_SafeName(name)) And all.Count >= STYLESTORE_BOOK_MAX Then Err.Raise vbObjectError + 7140, "Style", STYLESTORE_MSG_BOOK_FULL
    all(StyleDef_SafeName(name)) = text
    For Each k In all.Keys
        total = total + Len(CStr(all(k)))
    Next
    If total > STYLESTORE_BOOK_MAXCHARS Then Err.Raise vbObjectError + 7140, "Style", STYLESTORE_MSG_BOOK_FULL
    ReDim parts(0 To all.Count + 1)
    parts(0) = "<tqqStyles xmlns=""" & STYLESTORE_NS & """>"
    i = 1
    For Each k In all.Keys
        parts(i) = "<style name=""" & StyleStore_Esc(CStr(k)) & """>" & StyleStore_Esc(CStr(all(k))) & "</style>"
        i = i + 1
    Next
    parts(i) = "</tqqStyles>"
    xml = Join(parts, "")
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

' Texto do estilo: o nome de um embutido e SEMPRE o embutido (uma pasta ou um arquivo posto a mao com "Cascata MC" nao o sombreiam); senao
' preferUser (escolha explicita do usuario) = usuario, pasta; reaplicar = pasta, usuario.
Public Function StyleStore_Get(name As String, Optional wb As Workbook, Optional preferUser As Boolean = False) As String
    Dim t As String
    If StyleDef_IsBuiltin(name) Then
        StyleStore_Get = StyleDef_Builtin(name)
        Exit Function
    End If
    If preferUser Then
        t = StyleStore_UserText(name)
        If Len(t) = 0 And Not wb Is Nothing Then t = StyleStore_BookText(wb, name)
    Else
        If Not wb Is Nothing Then t = StyleStore_BookText(wb, name)
        If Len(t) = 0 Then t = StyleStore_UserText(name)
    End If
    StyleStore_Get = t
End Function

' Nomes: embutidos, do usuario (ordem do disco) e os so da pasta; sem repeticao; no maximo STYLESTORE_NAMES_MAX; "|". Nunca propaga.
Public Function StyleStore_Names(Optional wb As Workbook) As String
    Dim seen As Object, n As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    On Error Resume Next
    For Each n In StyleDef_BuiltinNames()
        If Not seen.Exists(CStr(n)) Then seen.Add CStr(n), True
    Next
    For Each n In StyleStore_UserNames()
        If seen.Count >= STYLESTORE_NAMES_MAX Then Exit For
        If Not seen.Exists(CStr(n)) Then seen.Add CStr(n), True
    Next
    If Not wb Is Nothing And seen.Count < STYLESTORE_NAMES_MAX Then
        For Each n In StyleStore_BookAll(wb).Keys
            If seen.Count >= STYLESTORE_NAMES_MAX Then Exit For
            If Not seen.Exists(CStr(n)) Then seen.Add CStr(n), True
        Next
    End If
    StyleStore_Names = Join(seen.Keys, "|")
    Err.Clear
End Function

' Pergunta de Importar quando o usuario ja tem um estilo com o nome do arquivo (revisao final 1.12.0, I1). Pura.
Public Function StyleStore_ReplaceQuestion(name As String) As String
    StyleStore_ReplaceQuestion = "Você já tem um estilo chamado '" & name & "'. Substituir pelo do arquivo?"
End Function

' Importa um arquivo de estilo (validado): grava no deposito do usuario com o nome do arquivo quando o "nome=" nao serve. Devolve o texto
' canonico gravado, "!" + mensagem, ou "?" + texto canonico (NADA gravado) quando o usuario ja tem um estilo com esse nome e allowOverwrite e
' False (quem chama pergunta e grava com StyleStore_SaveUser); name = nome decidido; ignored = linhas ignoradas.
Public Function StyleStore_Import(path As String, ByRef name As String, ByRef ignored As Long, Optional allowOverwrite As Boolean = False) As String
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
    If Not allowOverwrite And Len(StyleStore_UserText(name)) > 0 Then
        StyleStore_Import = "?" & t
        Exit Function
    End If
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
