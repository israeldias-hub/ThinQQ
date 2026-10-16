Attribute VB_Name = "modPptText"
Option Explicit
' Texto vinculado a celula (1.9.0; "Text Linked to Excel" do think-cell; spec 2026-10-15 secao 5.3). UM campo por forma: o texto e
' prefixo + valor exibido da celula + sufixo. A celula e lembrada por um nome definido OCULTO na pasta de origem (acompanha linhas inseridas)
' e por um endereco de reserva. Atualizar troca so os caracteres do campo (a formatacao do resto fica); edicao do usuario fora do campo nunca
' e sobrescrita. A pasta de origem e aberta so pelo motor (PptSync_OpenBook: somente leitura, macros desligadas).
' SEGURANCA: o nome definido (uma edicao da pasta do usuario) so e criado pela acao explicita "vincular" (PptText_Link, numa pasta ABERTA);
' a sincronizacao (PptText_Sync) so LE a pasta: nunca grava nela nem a salva. O valor que vai para o slide e saneado (sem caracteres de
' controle) e cortado em PPTTEXT_MAX_LEN caracteres.
' ETIQUETAS: o .pptx salvo perde Chr(30)/Chr(31) e outros caracteres de controle, troca Tab por espaco e CR por LF (medido na 1.6.0). Prefixo,
' sufixo e ultimo valor sao gravados codificados (PptText_TagEnc: "\" "\r" "\n" "\t" "\xHH"); o endereco de reserva e planilha & "|" & endereco
' (o endereco nunca tem "|": a divisao e pelo ULTIMO "|", entao o nome da planilha pode ter "|").
Public Const PPTTEXT_TAG As String = "TQQ_TEXT"
Public Const PPTLINK_TAG_NAME As String = "TQQ_TNAME"
Public Const PPTLINK_TAG_REF As String = "TQQ_TREF"
Public Const PPTTEXT_TAG_PRE As String = "TQQ_TPRE"
Public Const PPTTEXT_TAG_SUF As String = "TQQ_TSUF"
Public Const PPTTEXT_TAG_LAST As String = "TQQ_TLAST"
Public Const PPTTEXT_MAX_LEN As Long = 255                   ' valor da celula no slide: no maximo isto (o resto e cortado)
Public Const PPTTEXT_MSG_UNSAVED As String = "Salve a pasta de trabalho antes de vincular: o vínculo aponta para o arquivo salvo."
Public Const PPTTEXT_MSG_NOTBOOK As String = "Salve a pasta como pasta de trabalho do Excel (.xlsx, .xlsm, .xlsb ou .xls) antes de vincular."
Public Const PPTTEXT_MSG_NOTEXT As String = "Selecione uma caixa de texto (ou parte do texto dela) no PowerPoint."
Public Const PPTTEXT_MSG_NOCELL As String = "Selecione no Excel a célula com o valor e tente de novo."
Public Const PPTTEXT_MSG_CUT As String = "Texto vinculado. O valor da célula tem mais de 255 caracteres: o slide mostra só os primeiros 255."
Public Const PPTTEXT_MSG_EDITDATA As String = "Editar dados vale para gráficos do ThinQQ. Para mudar o texto vinculado, altere a célula no Excel e atualize."
Public Const PPTTEXT_MSG_NOTCHART As String = "Este comando vale só para gráficos do ThinQQ (não para texto vinculado)."
Private Const PPTLINK_REF_SEP As String = "|"

' Onde esta o campo no texto atual: texto = prefixo + ultimo + sufixo, ou ainda comeca pelo prefixo e termina pelo sufixo (o usuario editou so
' o campo). False = edicao incompativel (nao tocar). Pura.
Public Function PptText_Locate(cur As String, pre As String, suf As String, last As String, ByRef startPos As Long, ByRef length As Long) As Boolean
    If cur = pre & last & suf Then
        startPos = Len(pre) + 1: length = Len(last): PptText_Locate = True
        Exit Function
    End If
    If Len(cur) < Len(pre) + Len(suf) Then Exit Function
    If Left$(cur, Len(pre)) <> pre Then Exit Function
    If Right$(cur, Len(suf)) <> suf Then Exit Function
    startPos = Len(pre) + 1: length = Len(cur) - Len(pre) - Len(suf)
    PptText_Locate = True
End Function

' Texto exibido da celula (formato do Excel; erro de formula = o texto do erro, ex. "#DIV/0!"). Coluna estreita ("####"): o valor formatado
' com o formato da celula (PptText_FormatValue); se nada servir, o valor cru. So le a pasta (nada e alargado). Nunca propaga erro.
Public Function PptText_CellText(c As Range) As String
    Dim cl As Range, t As String, v As Variant, f As String
    On Error Resume Next
    Set cl = c.Cells(1, 1)
    t = CStr(cl.Text)
    If Len(t) > 0 And Len(Replace(t, "#", "")) = 0 Then
        v = cl.Value2
        If Not IsError(v) Then
            f = PptText_FormatValue(v, CStr(cl.NumberFormat))
            If Len(f) = 0 Then f = CStr(v)
            If Len(f) > 0 Then t = f
        End If
    End If
    Err.Clear
    PptText_CellText = t
End Function

' Valor formatado como a celula o mostraria (medido no Excel pt-BR: a funcao TEXTO do VBA devolve os separadores em INGLES, "1,234.5"). O
' formato (codigos em ingles, NumberFormat) passa pelo Format$ do VBA, que usa os separadores e os nomes de mes do Windows (os mesmos que o
' Excel mostra com "Usar separadores do sistema"); "General"/texto: CStr. Codigos que o Format$ nao entende ([cor], [$-416], _, *, ?, @): a
' funcao TEXTO com os separadores trocados para os do Excel. "" se nada servir. Nunca propaga erro.
Public Function PptText_FormatValue(v As Variant, fmt As String) As String
    Dim f As String, s As String, dec As String, thou As String
    On Error Resume Next
    f = Trim$(fmt)
    If VarType(v) = vbString Or Len(f) = 0 Or LCase$(f) = "general" Or f = "@" Then PptText_FormatValue = CStr(v): Exit Function
    If InStr(f, "[") > 0 Or InStr(f, "_") > 0 Or InStr(f, "*") > 0 Or InStr(f, "?") > 0 Or InStr(f, "@") > 0 Then
        s = CStr(Application.WorksheetFunction.Text(v, f))
        If Err.Number <> 0 Then Err.Clear: Exit Function
        dec = CStr(Application.International(xlDecimalSeparator)): thou = CStr(Application.International(xlThousandsSeparator))
        If Not CBool(Application.UseSystemSeparators) Then dec = CStr(Application.DecimalSeparator): thou = CStr(Application.ThousandsSeparator)
        If dec <> "." Or thou <> "," Then s = Replace(Replace(Replace(s, ",", Chr$(1)), ".", dec), Chr$(1), thou)
    Else
        s = Format$(v, f)
        If Err.Number <> 0 Then Err.Clear: s = ""
    End If
    PptText_FormatValue = s
End Function

' Valor seguro para a caixa de texto: caracteres de controle (quebras de linha, Tab, Chr(0)..Chr(31), Chr(127)) viram espaco e o texto e
' cortado em PPTTEXT_MAX_LEN (cut = True se cortou). Pura.
Public Function PptText_SafeValue(t As String, ByRef cut As Boolean) As String
    Dim s As String, i As Long, c As Long
    s = Replace(t, vbCrLf, " ")
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If (c >= 0 And c < 32) Or c = 127 Then Mid$(s, i, 1) = " "
    Next
    cut = (Len(s) > PPTTEXT_MAX_LEN)
    If cut Then s = Left$(s, PPTTEXT_MAX_LEN)
    PptText_SafeValue = s
End Function

' Codifica um texto para uma etiqueta do PowerPoint que sobrevive a salvar/reabrir: "\" -> "\\", CR -> "\r", LF -> "\n", Tab -> "\t", outros
' controles (Chr(0)..Chr(31), Chr(127), ex. a quebra de linha Chr(11) do PowerPoint) -> "\xHH". O resultado so tem caracteres imprimiveis. Pura.
Public Function PptText_TagEnc(s As String) As String
    Dim i As Long, c As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1): c = AscW(ch)
        Select Case c
            Case 92: out = out & "\\"
            Case 13: out = out & "\r"
            Case 10: out = out & "\n"
            Case 9: out = out & "\t"
            Case 0 To 31, 127: out = out & "\x" & Right$("0" & Hex$(c), 2)
            Case Else: out = out & ch
        End Select
    Next
    PptText_TagEnc = out
End Function

' Inverso de PptText_TagEnc. Etiqueta vem da apresentacao (de quem a fez): sequencia desconhecida ou truncada fica como esta; nunca levanta
' erro. Pura.
Public Function PptText_TagDec(s As String) As String
    Dim i As Long, n As Long, ch As String, nx As String, hx As String, out As String
    n = Len(s): i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "\" And i < n Then
            nx = Mid$(s, i + 1, 1)
            Select Case nx
                Case "\": out = out & "\": i = i + 2
                Case "r": out = out & vbCr: i = i + 2
                Case "n": out = out & vbLf: i = i + 2
                Case "t": out = out & vbTab: i = i + 2
                Case "x"
                    hx = Mid$(s, i + 2, 2)
                    If Len(hx) = 2 And hx Like "[0-9A-Fa-f][0-9A-Fa-f]" Then
                        out = out & ChrW$(CLng("&H" & hx)): i = i + 4
                    Else
                        out = out & ch: i = i + 1
                    End If
                Case Else: out = out & ch: i = i + 1
            End Select
        Else
            out = out & ch: i = i + 1
        End If
    Loop
    PptText_TagDec = out
End Function

' Endereco de reserva: planilha + "|" + endereco absoluto.
Public Function PptLink_RefOf(rng As Range) As String
    PptLink_RefOf = rng.Worksheet.Name & PPTLINK_REF_SEP & rng.Address
End Function

' Nome definido OCULTO na pasta de rng (prefixo TQQ_X_ texto, TQQ_T_ tabela). Edicao da pasta do usuario: so pela acao explicita de vincular
' (a pasta fica "nao salva"). Levanta erro se o Excel recusar (pasta com estrutura protegida etc.).
Public Function PptLink_DefineName(rng As Range, prefix As String) As String
    Dim wb As Workbook, nm As String, i As Long, n As Object
    Set wb = rng.Worksheet.Parent
    Do
        nm = prefix & Registry_NewId()
        Set n = Nothing
        On Error Resume Next
        Set n = wb.Names(nm)
        Err.Clear
        On Error GoTo 0
        i = i + 1
    Loop While Not n Is Nothing And i < 50
    wb.Names.Add Name:=nm, RefersTo:="='" & Replace(rng.Worksheet.Name, "'", "''") & "'!" & rng.Address, Visible:=False
    PptLink_DefineName = nm
End Function

' Intervalo do vinculo: pelo nome (acompanha insercoes) ou, sem ele, pelo endereco de reserva. Nothing se nenhum resolve. So le. Nunca
' propaga erro.
Public Function PptLink_ResolveRange(wb As Workbook, nm As String, ref As String) As Range
    Dim p As Long, r As Range
    On Error Resume Next
    If Len(nm) > 0 Then Set r = wb.Names(nm).RefersToRange
    Err.Clear
    If r Is Nothing And Len(ref) > 0 Then
        p = InStrRev(ref, PPTLINK_REF_SEP)
        If p > 1 And p < Len(ref) Then Set r = wb.Worksheets(Left$(ref, p - 1)).Range(Mid$(ref, p + 1))
        Err.Clear
    End If
    If Not r Is Nothing Then
        If r.Worksheet.Parent.FullName <> wb.FullName Then Set r = Nothing   ' nome apontando para outra pasta: nunca segue
    End If
    Err.Clear
    Set PptLink_ResolveRange = r
End Function

Private Sub PptText_InsertAt(tr As Object, st As Long, v As String)
    If st > Len(CStr(tr.Text)) Then tr.InsertAfter v Else tr.Characters(st, 0).InsertBefore v
End Sub

' Troca o campo [st, st + ln) do texto da forma por v (so esses caracteres: a formatacao do resto fica).
Private Sub PptText_Replace(shp As Object, st As Long, ln As Long, v As String)
    If ln > 0 Then
        shp.TextFrame.TextRange.Characters(st, ln).Text = v
    ElseIf Len(v) > 0 Then
        PptText_InsertAt shp.TextFrame.TextRange, st, v
    End If
End Sub

' Vincula a forma do slide a celula: o trecho [startPos, startPos + length) do texto vira o valor exibido da celula (length 0 = inserir no
' cursor; a forma inteira = 1 e Len). Devolve "" (ok), PPTTEXT_MSG_CUT (vinculado, valor cortado) ou a mensagem da recusa (pasta nao salva ou
' que nao e pasta do Excel, forma sem texto: nada muda). Transacional: o nome definido vem antes de tocar a forma; se o texto ou as
' etiquetas falharem, o nome e apagado, o texto volta e o erro do PowerPoint sobe.
Public Function PptText_Link(shp As Object, startPos As Long, length As Long, cell As Range) As String
    Dim wb As Workbook, tr As Object, cur As String, pre As String, suf As String, v As String, st As Long, ln As Long, hasTf As Boolean
    Dim cut As Boolean, nm As String, mid0 As String, changed As Boolean, wasLinked As Boolean, en As Long, ed As String
    Set wb = cell.Worksheet.Parent
    If Len(wb.Path) = 0 Then PptText_Link = PPTTEXT_MSG_UNSAVED: Exit Function
    If Not PptData_IsWorkbookExt(wb.FullName) Then PptText_Link = PPTTEXT_MSG_NOTBOOK: Exit Function
    If shp Is Nothing Then PptText_Link = PPTTEXT_MSG_NOTEXT: Exit Function
    On Error Resume Next
    hasTf = CBool(shp.HasTextFrame)
    Err.Clear
    On Error GoTo 0
    If Not hasTf Then PptText_Link = PPTTEXT_MSG_NOTEXT: Exit Function
    Set tr = shp.TextFrame.TextRange
    cur = CStr(tr.Text)
    st = startPos: ln = length
    If st < 1 Then st = 1
    If st > Len(cur) + 1 Then st = Len(cur) + 1
    If ln < 0 Then ln = 0
    If st + ln - 1 > Len(cur) Then ln = Len(cur) - st + 1
    pre = Left$(cur, st - 1): suf = Mid$(cur, st + ln): mid0 = Mid$(cur, st, ln)
    v = PptText_SafeValue(PptText_CellText(cell), cut)
    wasLinked = (PptSync_Tag(shp, PPTTEXT_TAG) = "1")
    nm = PptLink_DefineName(cell.Cells(1, 1), "TQQ_X_")          ' falha aqui: nada mudou
    On Error GoTo Undo
    PptText_Replace shp, st, ln, v
    changed = True
    shp.Tags.Add PPTSYNC_TAG_BOOK, wb.FullName
    shp.Tags.Add PPTLINK_TAG_NAME, nm
    shp.Tags.Add PPTLINK_TAG_REF, PptLink_RefOf(cell.Cells(1, 1))
    shp.Tags.Add PPTTEXT_TAG_PRE, PptText_TagEnc(pre)
    shp.Tags.Add PPTTEXT_TAG_SUF, PptText_TagEnc(suf)
    shp.Tags.Add PPTTEXT_TAG_LAST, PptText_TagEnc(v)
    shp.Tags.Add PPTTEXT_TAG, "1"                                ' por ultimo: so vira "texto vinculado" com as etiquetas completas
    If PptAuto_NewLinksAuto() Then PptSync_SetAuto shp, True
    If cut Then PptText_Link = PPTTEXT_MSG_CUT
    Exit Function
Undo:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    wb.Names(nm).Delete
    If changed Then
        If Len(v) > 0 Then shp.TextFrame.TextRange.Characters(st, Len(v)).Text = mid0 Else PptText_Replace shp, st, 0, mid0
    End If
    If Not wasLinked Then PptSync_TagDelete shp, PPTTEXT_TAG
    Err.Clear
    On Error GoTo 0
    Err.Raise en, "PptText_Link", ed
End Function

' Atualiza o campo da forma pela pasta wb (aberta; so LIDA: nada e gravado nela). PPTSYNC_R_SAME sem mudanca (e sem force); R_NOCHART celula
' nao encontrada; R_EDITED edicao incompativel (nada tocado); R_OK; R_FAIL (registrado). PowerPoint ocupado sobe.
Public Function PptText_Sync(shp As Object, wb As Workbook, force As Boolean) As Long
    Dim rng As Range, v As String, cur As String, pre As String, suf As String, last As String, st As Long, ln As Long, bn As Long, bd As String
    Dim cut As Boolean
    On Error GoTo Fail
    Set rng = PptLink_ResolveRange(wb, PptSync_Tag(shp, PPTLINK_TAG_NAME), PptSync_Tag(shp, PPTLINK_TAG_REF))
    If rng Is Nothing Then PptText_Sync = PPTSYNC_R_NOCHART: Exit Function
    v = PptText_SafeValue(PptText_CellText(rng), cut)
    If cut Then Errors_Log "PptText_Sync: valor da celula cortado em " & PPTTEXT_MAX_LEN & " caracteres" & vbTab & wb.Name
    pre = PptText_TagDec(PptSync_Tag(shp, PPTTEXT_TAG_PRE))
    suf = PptText_TagDec(PptSync_Tag(shp, PPTTEXT_TAG_SUF))
    last = PptText_TagDec(PptSync_Tag(shp, PPTTEXT_TAG_LAST))
    cur = CStr(shp.TextFrame.TextRange.Text)
    If Not force And v = last And cur = pre & last & suf Then PptText_Sync = PPTSYNC_R_SAME: Exit Function
    If Not PptText_Locate(cur, pre, suf, last, st, ln) Then PptText_Sync = PPTSYNC_R_EDITED: Exit Function
    If Mid$(cur, st, ln) <> v Then PptText_Replace shp, st, ln, v
    If v <> last Then shp.Tags.Add PPTTEXT_TAG_LAST, PptText_TagEnc(v)
    PptText_Sync = PPTSYNC_R_OK
    Exit Function
Fail:
    bn = Err.Number: bd = Err.Description
    On Error GoTo -1
    On Error GoTo 0                                     ' sem isto o GoTo -1 rearma o Fail e o Raise voltaria para ca (laco)
    If Ppt_IsBusyError(bn) Then Err.Raise bn, "PptText_Sync", bd
    Errors_Log "PptText_Sync" & vbTab & bn & vbTab & bd
    PptText_Sync = PPTSYNC_R_FAIL
End Function
