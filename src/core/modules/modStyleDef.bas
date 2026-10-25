Attribute VB_Name = "modStyleDef"
Option Explicit
' Estilos de grafico (1.12.0; spec 2026-10-23 secao 8.5): modelo PURO. Texto "chave=valor" (uma linha por chave, a primeira util =
' STYLEDEF_HEADER; linhas vazias e comecando por "#" ignoradas; BOM ignorado), validado chave a chave contra listas fechadas: valor invalido ou
' chave desconhecida sao ignorados e contados. Nada e executado (so Split, Like e comparacoes). Teto: STYLEDEF_MAX_BYTES caracteres e
' STYLEDEF_MAX_LINES linhas (texto maior e recusado inteiro).
Public Const STYLEDEF_HEADER As String = "ThinQQ-Estilo=1"
Public Const STYLEDEF_MAX_BYTES As Long = 16384
Public Const STYLEDEF_MAX_LINES As Long = 200
Public Const STYLEDEF_EXT As String = ".tqqestilo"
Private Const STYLEDEF_KEYLIST As String = ",nome,fonte,fonte.tamanho,serie1,serie2,serie3,serie4,serie5,serie6,serie7,serie8,total,aumento,reducao,primeira,primeira.padrao,primeira.padrao.cor,primeira.borda,ultima,ultima.padrao,ultima.padrao.cor,ultima.borda,rotulo.cor,rotulo.negrito,rotulo.tamanho,conector.cor,conector.traco,conector.espessura,seta.cor,seta.rotulo.fundo,seta.rotulo.borda,eixo.categoria.linha,espaco,"
Private Const STYLEDEF_PATTERNS As String = ",nenhum,diagonal_clara_para_cima,diagonal_escura_para_cima,diagonal_larga_para_cima,diagonal_clara_para_baixo,diagonal_escura_para_baixo,diagonal_larga_para_baixo,horizontal_escura,vertical_escura,xadrez_pequeno,pontilhado_50,grade_pequena,"

Public Function StyleDef_Keys() As Variant
    StyleDef_Keys = Split(Mid$(STYLEDEF_KEYLIST, 2, Len(STYLEDEF_KEYLIST) - 2), ",")
End Function

Public Function StyleDef_Hex(v As String) As String
    Dim s As String
    s = UCase$(Trim$(v))
    If Len(s) = 7 And Left$(s, 1) = "#" Then
        If Not Mid$(s, 2) Like "*[!0-9A-F]*" Then StyleDef_Hex = s
    End If
End Function

Public Function StyleDef_ColorOf(hexValue As String) As Long
    Dim s As String
    s = StyleDef_Hex(hexValue)
    If Len(s) = 0 Then StyleDef_ColorOf = -1: Exit Function
    StyleDef_ColorOf = RGB(CLng("&H" & Mid$(s, 2, 2)), CLng("&H" & Mid$(s, 4, 2)), CLng("&H" & Mid$(s, 6, 2)))
End Function

Public Function StyleDef_HexOf(rgbValue As Long) As String
    StyleDef_HexOf = "#" & Right$("0" & Hex$(rgbValue And &HFF&), 2) & Right$("0" & Hex$((rgbValue \ &H100&) And &HFF&), 2) & Right$("0" & Hex$((rgbValue \ &H10000) And &HFF&), 2)
End Function

Private Function StyleDef_IntIn(v As String, lo As Long, hi As Long) As Boolean
    If Len(v) = 0 Or Len(v) > 4 Then Exit Function
    If v Like "*[!0-9]*" Then Exit Function
    StyleDef_IntIn = (CLng(v) >= lo And CLng(v) <= hi)
End Function

' Numero com ponto ou virgula (no maximo um separador e 6 caracteres); -1 se invalido. Pura.
Private Function StyleDef_Num(v As String) As Double
    Dim s As String
    s = Replace(v, ",", ".")
    StyleDef_Num = -1
    If Len(s) = 0 Or Len(s) > 6 Then Exit Function
    If s Like "*[!0-9.]*" Then Exit Function
    If Len(s) - Len(Replace(s, ".", "")) > 1 Then Exit Function
    StyleDef_Num = Val(s)
End Function

' Nome de estilo seguro (arquivo, opcao do grafico, XML): sem ; = | " \ / : * ? < > & e sem controles; espacos das pontas fora; no maximo 40. Pura.
Public Function StyleDef_SafeName(s As String) As String
    Dim i As Long, ch As String, t As String, c As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1): c = AscW(ch)
        If (c >= 32 Or c < 0) And c <> 127 And InStr(";=|""\/:*?<>&", ch) = 0 Then t = t & ch
    Next
    t = Trim$(t)
    If Len(t) > 40 Then t = Trim$(Left$(t, 40))
    StyleDef_SafeName = t
End Function

' Valor normalizado da chave ou "" (invalido/chave desconhecida). Pura.
Public Function StyleDef_Validate(key As String, value As String) As String
    Dim v As String, d As Double
    v = Trim$(value)
    If InStr(key, ",") > 0 Or InStr(STYLEDEF_KEYLIST, "," & key & ",") = 0 Or Len(key) = 0 Then Exit Function
    Select Case key
        Case "nome"
            If Len(v) > 0 And StyleDef_SafeName(v) = v Then StyleDef_Validate = v
        Case "fonte"
            If Len(v) > 0 And Len(v) <= 40 Then
                If Not v Like "*[!A-Za-z0-9 -]*" Then StyleDef_Validate = v
            End If
        Case "fonte.tamanho", "rotulo.tamanho"
            If StyleDef_IntIn(v, 6, 24) Then StyleDef_Validate = CStr(CLng(v))
        Case "espaco"
            If StyleDef_IntIn(v, 0, 500) Then StyleDef_Validate = CStr(CLng(v))
        Case "rotulo.negrito", "eixo.categoria.linha"
            If v = "0" Or v = "1" Then StyleDef_Validate = v
        Case "primeira.padrao", "ultima.padrao"
            If Len(v) > 0 And InStr(v, ",") = 0 Then
                If InStr(STYLEDEF_PATTERNS, "," & LCase$(v) & ",") > 0 Then StyleDef_Validate = LCase$(v)
            End If
        Case "conector.traco"
            Select Case LCase$(v)
                Case "solido", "pontos", "tracos": StyleDef_Validate = LCase$(v)
            End Select
        Case "conector.espessura"
            d = StyleDef_Num(v)
            If d >= 0.25 And d <= 3 Then StyleDef_Validate = Replace(Format$(d, "0.00"), ",", ".")
        Case "rotulo.cor"
            If LCase$(v) = "auto" Then StyleDef_Validate = "auto" Else StyleDef_Validate = StyleDef_Hex(v)
        Case "primeira.borda", "ultima.borda", "seta.rotulo.borda"
            If LCase$(v) = "nenhuma" Then StyleDef_Validate = "nenhuma" Else StyleDef_Validate = StyleDef_Hex(v)
        Case Else
            StyleDef_Validate = StyleDef_Hex(v)                ' as chaves restantes da lista sao cores
    End Select
End Function

Public Function StyleDef_Get(d As Object, key As String) As String
    If d Is Nothing Then Exit Function
    If d.Exists(key) Then StyleDef_Get = CStr(d(key))
End Function

' Quantas cores de serie seguidas (serie1, serie2, ...) o estilo define. Pura.
Public Function StyleDef_SeriesCount(d As Object) As Long
    Dim i As Long
    For i = 1 To 8
        If Len(StyleDef_Get(d, "serie" & i)) = 0 Then Exit For
        StyleDef_SeriesCount = i
    Next
End Function

' Texto -> Dictionary (chave -> valor normalizado) das chaves validas; ignored = linhas desconhecidas ou invalidas. Nothing se o texto e
' recusado inteiro (vazio, maior que o teto, linhas demais, sem o cabecalho como 1a linha util). Pura.
Public Function StyleDef_Parse(text As String, ByRef ignored As Long) As Object
    Dim d As Object, lines() As String, i As Long, ln As String, p As Long, k As String, v As String, gotHeader As Boolean
    ignored = 0
    If Len(text) = 0 Or Len(text) > STYLEDEF_MAX_BYTES Then Exit Function
    lines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    If UBound(lines) + 1 > STYLEDEF_MAX_LINES Then Exit Function
    Set d = CreateObject("Scripting.Dictionary")
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Left$(ln, 1) = ChrW(65279) Then ln = Mid$(ln, 2)
        If Len(ln) > 0 And Left$(ln, 1) <> "#" Then
            If Not gotHeader Then
                If ln <> STYLEDEF_HEADER Then Exit Function
                gotHeader = True
            Else
                p = InStr(ln, "=")
                If p > 1 Then
                    k = LCase$(Trim$(Left$(ln, p - 1)))
                    v = StyleDef_Validate(k, Mid$(ln, p + 1))
                    If Len(v) > 0 Then d(k) = v Else ignored = ignored + 1
                Else
                    ignored = ignored + 1
                End If
            End If
        End If
    Next
    If gotHeader Then Set StyleDef_Parse = d
End Function

' Dictionary -> texto canonico: cabecalho e as chaves validas na ordem de STYLEDEF_KEYLIST, linhas por vbCrLf. Pura.
Public Function StyleDef_Serialize(d As Object) As String
    Dim k As Variant, s As String, v As String
    s = STYLEDEF_HEADER
    For Each k In StyleDef_Keys()
        If d.Exists(CStr(k)) Then
            v = StyleDef_Validate(CStr(k), CStr(d(CStr(k))))
            If Len(v) > 0 Then s = s & vbCrLf & CStr(k) & "=" & v
        End If
    Next
    StyleDef_Serialize = s
End Function

' Versao curta do estilo: 8 primeiros hex do SHA-256 do texto canonico. Pura.
Public Function StyleDef_Version(text As String) As String
    Dim d As Object, ign As Long, s As String, b() As Byte
    Set d = StyleDef_Parse(text, ign)
    If d Is Nothing Then s = text Else s = StyleDef_Serialize(d)
    b = StrConv(s, vbFromUnicode)
    StyleDef_Version = Left$(Sha256_Bytes(b, UBound(b) - LBound(b) + 1), 8)
End Function

Public Function StyleDef_PatternOf(nm As String) As Long
    Select Case LCase$(nm)
        Case "diagonal_clara_para_cima": StyleDef_PatternOf = msoPatternLightUpwardDiagonal
        Case "diagonal_escura_para_cima": StyleDef_PatternOf = msoPatternDarkUpwardDiagonal
        Case "diagonal_larga_para_cima": StyleDef_PatternOf = msoPatternWideUpwardDiagonal
        Case "diagonal_clara_para_baixo": StyleDef_PatternOf = msoPatternLightDownwardDiagonal
        Case "diagonal_escura_para_baixo": StyleDef_PatternOf = msoPatternDarkDownwardDiagonal
        Case "diagonal_larga_para_baixo": StyleDef_PatternOf = msoPatternWideDownwardDiagonal
        Case "horizontal_escura": StyleDef_PatternOf = msoPatternDarkHorizontal
        Case "vertical_escura": StyleDef_PatternOf = msoPatternDarkVertical
        Case "xadrez_pequeno": StyleDef_PatternOf = msoPatternSmallCheckerBoard
        Case "pontilhado_50": StyleDef_PatternOf = msoPattern50Percent
        Case "grade_pequena": StyleDef_PatternOf = msoPatternSmallGrid
    End Select
End Function

Public Function StyleDef_PatternName(pattern As Long) As String
    Dim k As Variant
    For Each k In Split(Mid$(STYLEDEF_PATTERNS, 2, Len(STYLEDEF_PATTERNS) - 2), ",")
        If StyleDef_PatternOf(CStr(k)) = pattern And pattern <> 0 Then StyleDef_PatternName = CStr(k): Exit Function
    Next
End Function

Public Function StyleDef_BuiltinNames() As Variant
    StyleDef_BuiltinNames = Array("Padrão ThinQQ", "Tons de cinza", "Cores do tema", "Cascata MC", "Cascata DDE")
End Function

Public Function StyleDef_IsBuiltin(name As String) As Boolean
    Dim n As Variant
    For Each n In StyleDef_BuiltinNames()
        If StrComp(CStr(n), name, vbTextCompare) = 0 Then StyleDef_IsBuiltin = True: Exit Function
    Next
End Function

' Paleta antiga (opcao "palette") dos embutidos que sao so paleta; "" para os demais. Pura.
Public Function StyleDef_BuiltinPalette(name As String) As String
    Select Case LCase$(name)
        Case LCase$("Padrão ThinQQ"): StyleDef_BuiltinPalette = "padrao"
        Case "tons de cinza": StyleDef_BuiltinPalette = "cinza"
        Case "cores do tema": StyleDef_BuiltinPalette = "tema"
    End Select
End Function

' Texto dos embutidos (validos por construcao: Test_StyleDef_ColorsPatternsAndBuiltins). "" para nome desconhecido. Pura.
Public Function StyleDef_Builtin(name As String) As String
    Dim s As String, nl As String
    nl = vbCrLf
    Select Case LCase$(name)
        Case LCase$("Padrão ThinQQ"): s = "nome=Padrão ThinQQ"
        Case "tons de cinza": s = "nome=Tons de cinza"
        Case "cores do tema": s = "nome=Cores do tema"
        Case "cascata mc"
            s = "nome=Cascata MC" & nl & "serie1=#70AD47" & nl & "serie2=#D9D9D9" & nl & "serie3=#E8173C" & nl & "total=#D9D9D9" & nl & "aumento=#70AD47" & nl & "reducao=#E8173C"
            s = s & nl & "primeira=#FFFFFF" & nl & "primeira.padrao=diagonal_larga_para_cima" & nl & "primeira.padrao.cor=#000000" & nl & "primeira.borda=#000000"
            s = s & nl & "ultima=#D9D9D9" & nl & "rotulo.cor=auto" & nl & "rotulo.negrito=1" & nl & "conector.cor=#7F7F7F" & nl & "conector.traco=pontos"
            s = s & nl & "conector.espessura=0.75" & nl & "seta.cor=#404040" & nl & "seta.rotulo.fundo=#FFFFFF" & nl & "seta.rotulo.borda=#1F3864" & nl & "eixo.categoria.linha=1"
        Case "cascata dde"
            s = "nome=Cascata DDE" & nl & "serie1=#4472C4" & nl & "serie2=#ED7D31" & nl & "serie3=#70AD47" & nl & "total=#4472C4" & nl & "aumento=#ED7D31" & nl & "reducao=#70AD47"
            s = s & nl & "primeira=#FFFFFF" & nl & "primeira.padrao=diagonal_escura_para_cima" & nl & "primeira.padrao.cor=#4472C4" & nl & "primeira.borda=#4472C4"
            s = s & nl & "ultima=#4472C4" & nl & "rotulo.cor=auto" & nl & "conector.cor=#A6A6A6" & nl & "conector.traco=solido" & nl & "conector.espessura=0.75"
        Case Else: Exit Function
    End Select
    StyleDef_Builtin = STYLEDEF_HEADER & nl & s
End Function
