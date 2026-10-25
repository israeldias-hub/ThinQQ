Attribute VB_Name = "modStyleUi"
Option Explicit
' Textos e listas da janela Estilos (1.12.0). Puras.
Private Const STYLEUI_COLORS As String = "Preto|#000000;Branco|#FFFFFF;Cinza-escuro|#404040;Cinza|#7F7F7F;Cinza-médio|#A6A6A6;Cinza-claro|#D9D9D9;Azul-marinho|#1F3864;Azul|#4472C4;Azul-claro|#9DC3E6;Verde|#70AD47;Verde-escuro|#548235;Verde-claro|#A9D18E;Laranja|#ED7D31;Amarelo|#FFC000;Vermelho|#E8173C;Vinho|#C00000;Roxo|#7030A0;Turquesa|#009688;Âmbar|#EDA224;Salmão|#F4B183"

Public Function StyleUi_IsColorKey(key As String) As Boolean
    Select Case key
        Case "nome", "fonte", "fonte.tamanho", "rotulo.tamanho", "espaco", "rotulo.negrito", "eixo.categoria.linha", "primeira.padrao", "ultima.padrao", "conector.traco", "conector.espessura"
        Case Else: StyleUi_IsColorKey = (Len(StyleDef_Validate(key, "#000000")) > 0)
    End Select
End Function

Public Function StyleUi_RoleLabel(key As String) As String
    Select Case key
        Case "nome": StyleUi_RoleLabel = "Nome"
        Case "fonte": StyleUi_RoleLabel = "Fonte"
        Case "fonte.tamanho": StyleUi_RoleLabel = "Tamanho da fonte"
        Case "total": StyleUi_RoleLabel = "Totais"
        Case "aumento": StyleUi_RoleLabel = "Aumentos (cascata)"
        Case "reducao": StyleUi_RoleLabel = "Reduções (cascata)"
        Case "primeira": StyleUi_RoleLabel = "Primeira barra"
        Case "primeira.padrao": StyleUi_RoleLabel = "Primeira barra: hachura"
        Case "primeira.padrao.cor": StyleUi_RoleLabel = "Primeira barra: cor da hachura"
        Case "primeira.borda": StyleUi_RoleLabel = "Primeira barra: borda"
        Case "ultima": StyleUi_RoleLabel = "Última barra"
        Case "ultima.padrao": StyleUi_RoleLabel = "Última barra: hachura"
        Case "ultima.padrao.cor": StyleUi_RoleLabel = "Última barra: cor da hachura"
        Case "ultima.borda": StyleUi_RoleLabel = "Última barra: borda"
        Case "rotulo.cor": StyleUi_RoleLabel = "Rótulos: cor (auto = contraste)"
        Case "rotulo.negrito": StyleUi_RoleLabel = "Rótulos: negrito"
        Case "rotulo.tamanho": StyleUi_RoleLabel = "Rótulos: tamanho"
        Case "conector.cor": StyleUi_RoleLabel = "Conectores: cor"
        Case "conector.traco": StyleUi_RoleLabel = "Conectores: traço"
        Case "conector.espessura": StyleUi_RoleLabel = "Conectores: espessura (pt)"
        Case "seta.cor": StyleUi_RoleLabel = "Setas: cor"
        Case "seta.rotulo.fundo": StyleUi_RoleLabel = "Setas: fundo do rótulo"
        Case "seta.rotulo.borda": StyleUi_RoleLabel = "Setas: borda do rótulo"
        Case "eixo.categoria.linha": StyleUi_RoleLabel = "Linha do eixo de categorias"
        Case "espaco": StyleUi_RoleLabel = "Espaço entre colunas (%)"
        Case Else
            If Left$(key, 5) = "serie" Then StyleUi_RoleLabel = "Série " & Mid$(key, 6)
    End Select
End Function

' Rotulos da lista de escolha do papel (Empty = so digitar). Pura.
Public Function StyleUi_PickItems(key As String) As Variant
    Dim a() As String, p() As String, i As Long
    Select Case key
        Case "primeira.padrao", "ultima.padrao"
            StyleUi_PickItems = Array("nenhum", "diagonal_clara_para_cima", "diagonal_escura_para_cima", "diagonal_larga_para_cima", "diagonal_clara_para_baixo", "diagonal_escura_para_baixo", "diagonal_larga_para_baixo", "horizontal_escura", "vertical_escura", "xadrez_pequeno", "pontilhado_50", "grade_pequena")
        Case "conector.traco": StyleUi_PickItems = Array("solido", "pontos", "tracos")
        Case "rotulo.negrito", "eixo.categoria.linha": StyleUi_PickItems = Array("0 (não)", "1 (sim)")
        Case Else
            If StyleUi_IsColorKey(key) Then
                p = Split(STYLEUI_COLORS, ";")
                ReDim a(0 To UBound(p))
                For i = 0 To UBound(p)
                    a(i) = Replace(p(i), "|", " ")
                Next
                StyleUi_PickItems = a
            End If
    End Select
End Function

' Valor (o que vai para o estilo) do item index da lista. Pura.
Public Function StyleUi_PickValue(key As String, index As Long) As String
    Dim items As Variant, s As String
    items = StyleUi_PickItems(key)
    If Not IsArray(items) Then Exit Function
    If index < LBound(items) Or index > UBound(items) Then Exit Function
    s = CStr(items(index))
    If StyleUi_IsColorKey(key) Then
        StyleUi_PickValue = Mid$(s, InStrRev(s, " ") + 1)
    ElseIf key = "rotulo.negrito" Or key = "eixo.categoria.linha" Then
        StyleUi_PickValue = Left$(s, 1)
    Else
        StyleUi_PickValue = s
    End If
End Function
