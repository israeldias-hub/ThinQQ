Attribute VB_Name = "TestStyles"
Option Explicit
' Estilos de grafico (1.12.0; spec 2026-10-23 secao 8.5): modelo puro, depositos, aplicar por papel e sobrevivencia.

Public Sub Test_StyleDef_Validate_Rules()
    AssertEqual "#70AD47", StyleDef_Validate("aumento", " #70ad47 ")
    AssertEqual "", StyleDef_Validate("aumento", "#70AD4"), "hex curto"
    AssertEqual "", StyleDef_Validate("aumento", "70AD47"), "sem #"
    AssertEqual "", StyleDef_Validate("aumento", "#GGGGGG")
    AssertEqual "auto", StyleDef_Validate("rotulo.cor", "AUTO")
    AssertEqual "#FFFFFF", StyleDef_Validate("rotulo.cor", "#ffffff")
    AssertEqual "nenhuma", StyleDef_Validate("primeira.borda", "nenhuma")
    AssertEqual "diagonal_larga_para_cima", StyleDef_Validate("primeira.padrao", "Diagonal_Larga_Para_Cima")
    AssertEqual "", StyleDef_Validate("primeira.padrao", "nenhum,xadrez_pequeno"), "virgula nunca casa dois itens"
    AssertEqual "", StyleDef_Validate("primeira.padrao", "qualquer")
    AssertEqual "9", StyleDef_Validate("fonte.tamanho", "9")
    AssertEqual "", StyleDef_Validate("fonte.tamanho", "5")
    AssertEqual "", StyleDef_Validate("fonte.tamanho", "9.5")
    AssertEqual "", StyleDef_Validate("espaco", "-1")
    AssertEqual "500", StyleDef_Validate("espaco", "500")
    AssertEqual "0.75", StyleDef_Validate("conector.espessura", "0,75")
    AssertEqual "", StyleDef_Validate("conector.espessura", "3.5")
    AssertEqual "", StyleDef_Validate("conector.espessura", "1.2.3")
    AssertEqual "pontos", StyleDef_Validate("conector.traco", "Pontos")
    AssertEqual "", StyleDef_Validate("conector.traco", "zigue")
    AssertEqual "1", StyleDef_Validate("rotulo.negrito", "1")
    AssertEqual "", StyleDef_Validate("rotulo.negrito", "sim")
    AssertEqual "Calibri Light", StyleDef_Validate("fonte", "Calibri Light")
    AssertEqual "", StyleDef_Validate("fonte", "Calibri;=x")
    AssertEqual "", StyleDef_Validate("desconhecida", "#000000")
    AssertEqual "", StyleDef_Validate("serie9", "#000000"), "so serie1..serie8"
End Sub

Public Sub Test_StyleDef_SafeName()
    AssertEqual "Cascata MC", StyleDef_SafeName("  Cascata MC  ")
    AssertEqual "ab", StyleDef_SafeName("a;=|""\/:*?<>&b")
    AssertEqual "Padrão ThinQQ", StyleDef_SafeName("Padrão ThinQQ")
    AssertEqual 40, Len(StyleDef_SafeName(String$(60, "x")))
    AssertEqual "", StyleDef_SafeName(vbTab & vbCr)
End Sub

Public Sub Test_StyleDef_ParseSerialize_RoundTrip()
    Dim d As Object, ign As Long, t As String, d2 As Object
    t = ChrW(65279) & STYLEDEF_HEADER & vbLf & "# comentario" & vbCrLf & "nome=Meu" & vbLf & "AUMENTO=#70ad47" & vbLf & "reducao=vermelho" & vbLf & "lixo" & vbLf & "outra=1" & vbLf & vbLf & "espaco=40"
    Set d = StyleDef_Parse(t, ign)
    AssertTrue Not d Is Nothing
    AssertEqual 3, ign, "reducao invalida, linha sem '=', chave desconhecida"
    AssertEqual "#70AD47", StyleDef_Get(d, "aumento")
    AssertEqual "40", StyleDef_Get(d, "espaco")
    AssertEqual "", StyleDef_Get(d, "reducao")
    t = StyleDef_Serialize(d)
    AssertEqual STYLEDEF_HEADER & vbCrLf & "nome=Meu" & vbCrLf & "aumento=#70AD47" & vbCrLf & "espaco=40", t, "ordem canonica"
    Set d2 = StyleDef_Parse(t, ign)
    AssertEqual 0, ign
    AssertEqual t, StyleDef_Serialize(d2), "ida e volta"
    AssertEqual 8, Len(StyleDef_Version(t))
    AssertTrue StyleDef_Version(t) <> StyleDef_Version(t & vbCrLf & "total=#000000")
End Sub

Public Sub Test_StyleDef_Parse_RefusesHostile()
    Dim ign As Long, i As Long, s As String
    AssertTrue StyleDef_Parse("", ign) Is Nothing
    AssertTrue StyleDef_Parse("nome=x", ign) Is Nothing, "sem cabecalho"
    AssertTrue StyleDef_Parse(STYLEDEF_HEADER & vbLf & String$(STYLEDEF_MAX_BYTES, "a"), ign) Is Nothing, "grande demais"
    s = STYLEDEF_HEADER
    For i = 1 To STYLEDEF_MAX_LINES + 1
        s = s & vbLf & "x=1"
    Next
    AssertTrue StyleDef_Parse(s, ign) Is Nothing, "linhas demais"
    AssertTrue Not StyleDef_Parse(STYLEDEF_HEADER & vbLf & "nome==<script>" & vbLf & "fonte=a&b", ign) Is Nothing
    AssertEqual 2, ign, "nome e fonte invalidos ignorados, nada executado"
End Sub

Public Sub Test_StyleDef_ColorsPatternsAndBuiltins()
    Dim n As Variant, d As Object, ign As Long
    AssertEqual RGB(112, 173, 71), StyleDef_ColorOf("#70AD47")
    AssertEqual -1, StyleDef_ColorOf("x")
    AssertEqual "#70AD47", StyleDef_HexOf(RGB(112, 173, 71))
    AssertEqual msoPatternWideUpwardDiagonal, StyleDef_PatternOf("diagonal_larga_para_cima")
    AssertEqual 0, StyleDef_PatternOf("nenhum")
    AssertEqual "diagonal_escura_para_cima", StyleDef_PatternName(msoPatternDarkUpwardDiagonal)
    AssertEqual "", StyleDef_PatternName(999)
    AssertEqual 5, UBound(StyleDef_BuiltinNames()) + 1
    For Each n In StyleDef_BuiltinNames()
        Set d = StyleDef_Parse(StyleDef_Builtin(CStr(n)), ign)
        AssertTrue Not d Is Nothing, CStr(n)
        AssertEqual 0, ign, "embutido sem linha invalida: " & CStr(n)
        AssertEqual CStr(n), StyleDef_Get(d, "nome")
        AssertTrue StyleDef_IsBuiltin(CStr(n))
    Next
    AssertEqual "padrao", StyleDef_BuiltinPalette("Padrão ThinQQ")
    AssertEqual "", StyleDef_BuiltinPalette("Cascata MC")
    Set d = StyleDef_Parse(StyleDef_Builtin("Cascata MC"), ign)
    AssertEqual "#70AD47", StyleDef_Get(d, "aumento")
    AssertEqual "#E8173C", StyleDef_Get(d, "reducao")
    AssertEqual "#D9D9D9", StyleDef_Get(d, "total")
    AssertEqual "diagonal_larga_para_cima", StyleDef_Get(d, "primeira.padrao")
    AssertEqual 3, StyleDef_SeriesCount(d)
    AssertFalse StyleDef_IsBuiltin("Meu")
End Sub

Private Function StyleTestDir() As String
    Dim d As String
    d = TempDir() & "\styles"
    If Len(Dir(d, vbDirectory)) = 0 Then MkDir d
    StyleStore_SetStoreDirForTest d
    StyleTestDir = d
End Function

Public Sub Test_StyleStore_UserSaveGetDelete_TestDirOnly()
    Dim d As String, t As String
    StyleStore_ResetTestHooks
    AssertEqual "", StyleStore_UserDir(), "build Test sem gancho: nunca a pasta real do usuario"
    AssertTrue Left$(StyleStore_SaveUser("X", STYLEDEF_HEADER), 1) = "!", "sem deposito: recusa"
    d = StyleTestDir()
    t = STYLEDEF_HEADER & vbCrLf & "nome=Meu Estilo" & vbCrLf & "aumento=#00FF00"
    AssertEqual "", StyleStore_SaveUser("Meu Estilo", t)
    AssertTrue Len(Dir(d & "\Meu Estilo" & STYLEDEF_EXT)) > 0
    AssertEqual t, StyleStore_UserText("meu estilo"), "sem diferenciar maiusculas"
    AssertTrue InStr("|" & StyleStore_Names() & "|", "|Meu Estilo|") > 0
    AssertTrue InStr("|" & StyleStore_Names() & "|", "|Cascata MC|") > 0, "embutidos na lista"
    AssertTrue Left$(StyleStore_SaveUser("Cascata MC", t), 1) = "!", "nome de embutido recusado"
    AssertEqual "", StyleStore_DeleteUser("Meu Estilo")
    AssertEqual "", StyleStore_UserText("Meu Estilo")
    StyleStore_ResetTestHooks
End Sub

Public Sub Test_StyleStore_Book_CustomXml_SaveReopen()
    Dim wb As Workbook, p As String, t As String
    Set wb = NewTempWorkbook()
    t = STYLEDEF_HEADER & vbCrLf & "nome=Da Pasta" & vbCrLf & "total=#123456"
    StyleStore_SaveBook wb, "Da Pasta", t
    StyleStore_SaveBook wb, "Outro <&>", STYLEDEF_HEADER
    AssertEqual t, Replace(StyleStore_BookText(wb, "da pasta"), vbLf, vbCrLf), "texto de volta (o XML normaliza CRLF)"
    AssertEqual 1, wb.CustomXMLParts.SelectByNamespace(STYLESTORE_NS).Count, "uma parte so"
    p = TempDir() & "\style-book.xlsx"
    wb.SaveAs p, xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Set wb = Workbooks.Open(p)
    AssertTrue InStr(StyleStore_BookText(wb, "Da Pasta"), "total=#123456") > 0, "sobrevive a salvar e reabrir"
    AssertEqual 2, StyleStore_BookAll(wb).Count
    AssertTrue InStr(StyleStore_Get("Da Pasta", wb), "#123456") > 0, "a pasta acha o estilo que o usuario nao tem"
    CloseTemp wb
    Kill p
End Sub

Public Sub Test_StyleStore_ImportExport_Validates()
    Dim d As String, p As String, nm As String, ign As Long, r As String, f As Integer
    StyleStore_ResetTestHooks
    d = StyleTestDir()
    p = d & "\entrada.tqqestilo"
    f = FreeFile
    Open p For Output As #f
    Print #f, STYLEDEF_HEADER
    Print #f, "nome=Recebido;x"
    Print #f, "aumento=#ABCDEF"
    Print #f, "comando=executar o calc"
    Close #f
    r = StyleStore_Import(p, nm, ign)
    AssertTrue Left$(r, 1) <> "!", r
    AssertEqual "entrada", nm, "nome invalido no arquivo: vale o nome do arquivo"
    AssertEqual 2, ign
    AssertTrue InStr(StyleStore_UserText("entrada"), "aumento=#ABCDEF") > 0
    AssertEqual "", StyleStore_Export("entrada", d & "\saida.tqqestilo")
    AssertTrue InStr(StyleStore_ReadFile(d & "\saida.tqqestilo", r), "aumento=#ABCDEF") > 0
    f = FreeFile
    Open d & "\ruim.tqqestilo" For Output As #f
    Print #f, "<xml>nada</xml>"
    Close #f
    AssertTrue Left$(StyleStore_Import(d & "\ruim.tqqestilo", nm, ign), 1) = "!", "sem cabecalho"
    StyleStore_ResetTestHooks
End Sub

Public Sub Test_StyleStore_Default_TestMemory()
    StyleStore_ResetTestHooks
    AssertEqual "", StyleStore_DefaultName()
    StyleStore_SetDefaultName "Cascata MC"
    AssertEqual "Cascata MC", StyleStore_DefaultName()
    StyleStore_ResetTestHooks
    AssertEqual "", StyleStore_DefaultName()
End Sub
