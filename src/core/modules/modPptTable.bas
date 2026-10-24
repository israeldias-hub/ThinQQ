Attribute VB_Name = "modPptTable"
Option Explicit
' Tabela nativa do PowerPoint vinculada a um intervalo do Excel (1.10.0; "Create tables with Excel formatting" do think-cell; spec 2026-10-15
' secao 5.4). A ultima matriz de textos escrita fica na etiqueta TQQ_TBLAST: a atualizacao escreve so as celulas cuja ORIGEM mudou (a
' formatacao feita no PowerPoint fica). Limites do PowerPoint: 75 linhas e 75 colunas (o ThinQQ limita a 25 colunas).
' ETIQUETA (adendo do controlador, medido na 1.6.0): o .pptx salvo perde Chr(30)/Chr(31) e outros controles, troca Tab por espaco e CR por LF;
' LF e "|" sobrevivem. TQQ_TBLAST = linhas separadas por LF, celulas por "|", cada celula codificada por PptText_TagEnc ("\\" "\r" "\n" "\t"
' "\xHH") e com "|" -> "\x7C": so caracteres imprimiveis e os separadores, ida e volta sem perda.
Public Const PPTTABLE_TAG As String = "TQQ_TABLE"
Public Const PPTTABLE_TAG_LAST As String = "TQQ_TBLAST"
Public Const PPTTABLE_MAX_ROWS As Long = 75
Public Const PPTTABLE_MAX_COLS As Long = 25
Public Const PPTTABLE_MAX_TAG As Long = 200000               ' TQQ_TBLAST maior que isto nao e gravada (a proxima atualizacao reescreve tudo)
Private Const PPTTABLE_RS As String = vbLf                    ' separador de linhas na etiqueta
Private Const PPTTABLE_FS As String = "|"                     ' separador de celulas na etiqueta
Private Const PPTTABLE_EMPTY1 As String = "\-"                ' etiqueta da matriz 1 x 1 vazia (I1)
Public Const PPTTABLE_MAX_CELLENC As Long = 1020             ' 4 x PPTTEXT_MAX_LEN: celula codificada maior = etiqueta recusada (I2)
Public Const PPTTABLE_NAME_PREFIX As String = "TQQ_T_"      ' nome definido oculto do intervalo (so a acao explicita de inserir o cria)
' Estilo "Sem Estilo, Grade da Tabela" do PowerPoint: sem preenchimento e com grade fina; a formatacao do Excel vai por cima (so o que difere
' do padrao e escrito: menos chamadas ao PowerPoint por celula).
Private Const PPTTABLE_STYLE_GRID As String = "{5940675A-B579-460E-94D1-54222C63F5DA}"
Private Const PPTTABLE_ERR As Long = 7120                    ' vbObjectError + 7120: recusa com mensagem para o usuario
Public Const PPTTABLE_MSG_EDITDATA As String = "Editar dados vale para gráficos do ThinQQ. Para mudar a tabela vinculada, altere as células no Excel e atualize."
Public Const PPTTABLE_MSG_NOTCHART As String = "Este comando vale só para gráficos do ThinQQ (não para tabela vinculada)."
Public Const PPTTABLE_MSG_NOSEL As String = "Selecione as células da tabela."
Public Const PPTTABLE_MSG_SAVEHINT As String = "Salve a pasta de trabalho: o vínculo fica guardado nela e acompanha linhas e colunas inseridas dentro do intervalo."

' Matriz 1..r x 1..c de textos -> etiqueta (ver cabecalho). Pura.
Public Function PptTable_Serialize(m As Variant) As String
    Dim r As Long, c As Long, rowsOut() As String, cells() As String
    ReDim rowsOut(1 To UBound(m, 1))
    ReDim cells(1 To UBound(m, 2))
    For r = 1 To UBound(m, 1)
        For c = 1 To UBound(m, 2)
            cells(c) = Replace(PptText_TagEnc(CStr(m(r, c))), PPTTABLE_FS, "\x7C")
        Next
        rowsOut(r) = Join(cells, PPTTABLE_FS)
    Next
    PptTable_Serialize = Join(rowsOut, PPTTABLE_RS)
    ' 1 x 1 vazia daria "" (o mesmo que "sem etiqueta"): marcador proprio, que PptText_TagEnc nunca produz ("\" vira "\\") (revisao final F, I1)
    If Len(PptTable_Serialize) = 0 Then PptTable_Serialize = PPTTABLE_EMPTY1
End Function

' Inverso de PptTable_Serialize: matriz 1..r x 1..c de String (colunas = as da 1a linha, no minimo 1; linha curta = celulas vazias) ou Empty
' (etiqueta vazia, maior que PPTTABLE_MAX_TAG, com mais de PPTTABLE_MAX_ROWS linhas / PPTTABLE_MAX_COLS colunas ou com uma celula codificada
' maior que PPTTABLE_MAX_CELLENC: o maximo legitimo, 255 caracteres de "\xHH"). A etiqueta vem da apresentacao (de quem a fez): nunca levanta
' erro e o trabalho e limitado (revisao final F, I2). Pura.
Public Function PptTable_Parse(s As String) As Variant
    Dim rows() As String, cells() As String, m() As String, r As Long, c As Long, nc As Long
    On Error GoTo Bad
    If Len(s) = 0 Or Len(s) > PPTTABLE_MAX_TAG Then Exit Function
    If s = PPTTABLE_EMPTY1 Then
        ReDim m(1 To 1, 1 To 1)
        PptTable_Parse = m
        Exit Function
    End If
    rows = Split(s, PPTTABLE_RS)
    If UBound(rows) + 1 > PPTTABLE_MAX_ROWS Then Exit Function
    nc = UBound(Split(rows(0), PPTTABLE_FS)) + 1
    If nc < 1 Then nc = 1                                       ' 1a celula vazia numa tabela de uma coluna: Split("") nao tem itens
    If nc > PPTTABLE_MAX_COLS Then Exit Function
    ReDim m(1 To UBound(rows) + 1, 1 To nc)
    For r = 0 To UBound(rows)
        cells = Split(rows(r), PPTTABLE_FS)
        For c = 0 To nc - 1
            If c <= UBound(cells) Then
                If Len(cells(c)) > PPTTABLE_MAX_CELLENC Then Exit Function
                m(r + 1, c + 1) = PptText_TagDec(cells(c))
            End If
        Next
    Next
    PptTable_Parse = m
    Exit Function
Bad:
    PptTable_Parse = Empty
End Function
' Celulas de newM a escrever: diferentes de oldM ou fora das dimensoes dele (oldM Empty = todas). Itens Array(linha, coluna). Pura.
Public Function PptTable_Diff(oldM As Variant, newM As Variant) As Collection
    Dim out As New Collection, r As Long, c As Long, inOld As Boolean
    For r = 1 To UBound(newM, 1)
        For c = 1 To UBound(newM, 2)
            inOld = False
            If Not IsEmpty(oldM) Then inOld = (r <= UBound(oldM, 1) And c <= UBound(oldM, 2))
            If Not inOld Then
                out.Add Array(r, c)
            ElseIf CStr(oldM(r, c)) <> CStr(newM(r, c)) Then
                out.Add Array(r, c)
            End If
        Next
    Next
    Set PptTable_Diff = out
End Function

' Texto seguro para uma celula da tabela: quebra de linha da celula (CR, LF, CRLF: Alt+Enter) vira a quebra de linha do PowerPoint (Chr(11));
' outros caracteres de controle viram espaco; cortado em PPTTEXT_MAX_LEN sem partir par substituto (emoji). Pura.
Public Function PptTable_SafeCell(t As String) As String
    Dim s As String, i As Long, c As Long
    s = Replace(Replace(Replace(t, vbCrLf, vbLf), vbCr, vbLf), vbLf, Chr$(11))
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If ((c >= 0 And c < 32) Or c = 127) And c <> 11 Then Mid$(s, i, 1) = " "
    Next
    If Len(s) > PPTTEXT_MAX_LEN Then
        c = AscW(Mid$(s, PPTTEXT_MAX_LEN, 1)) And &HFFFF&
        If c >= &HD800& And c <= &HDBFF& Then s = Left$(s, PPTTEXT_MAX_LEN - 1) Else s = Left$(s, PPTTEXT_MAX_LEN)
    End If
    PptTable_SafeCell = s
End Function

' Textos exibidos das celulas de rng (PptText_CellText: formato do Excel, coluna estreita sem "####"), saneados (PptTable_SafeCell). So le.
Public Function PptTable_TextMatrix(rng As Range) As Variant
    Dim m() As String, r As Long, c As Long
    ReDim m(1 To rng.Rows.Count, 1 To rng.Columns.Count)
    For r = 1 To rng.Rows.Count
        For c = 1 To rng.Columns.Count
            m(r, c) = PptTable_SafeCell(PptText_CellText(rng.Cells(r, c)))
        Next
    Next
    PptTable_TextMatrix = m
End Function

' "" = o intervalo serve; senao o motivo (varias areas, tamanho, mescladas).
Public Function PptTable_RefusalText(rng As Range) As String
    If rng.Areas.Count > 1 Then PptTable_RefusalText = "Selecione um único intervalo retangular.": Exit Function
    If rng.Rows.Count > PPTTABLE_MAX_ROWS Or rng.Columns.Count > PPTTABLE_MAX_COLS Then PptTable_RefusalText = "A tabela vinculada aceita até " & PPTTABLE_MAX_ROWS & " linhas e " & PPTTABLE_MAX_COLS & " colunas.": Exit Function
    If IsNull(rng.MergeCells) Then
        PptTable_RefusalText = "A tabela vinculada não aceita células mescladas no intervalo."
    ElseIf rng.MergeCells Then
        PptTable_RefusalText = "A tabela vinculada não aceita células mescladas no intervalo."
    End If
End Function

' ---- Criar e atualizar a tabela no slide (F2) ----

' Gancho de teste (Util_FailThinSyncForTest = "ppttable-<ponto>:1"): dispara UMA vez. Inerte fora do build Test.
Public Sub PptTable_TestHook(site As String)
    If BUILD_CONFIG = "Test" Then
        If Util_FailThinSyncForTest = "ppttable-" & site & ":1" Then
            Util_FailThinSyncForTest = ""
            Err.Raise vbObjectError + 9993, "ThinQQ", "Falha simulada (teste) " & site
        End If
    End If
End Sub

' Alinhamento do Excel para o do PowerPoint (ppAlignLeft 1, ppAlignCenter 2, ppAlignRight 3). Geral: numero a direita, texto a esquerda.
Private Function PptTable_Align(f As Object, src As Range) As Long
    Dim v As Variant
    Select Case f.HorizontalAlignment
        Case xlRight: PptTable_Align = 3
        Case xlCenter, xlCenterAcrossSelection: PptTable_Align = 2
        Case xlLeft: PptTable_Align = 1
        Case Else
            v = src.Value2
            If IsEmpty(v) Or IsError(v) Then
                PptTable_Align = 1
            ElseIf VarType(v) = vbString Then
                PptTable_Align = 1
            Else
                PptTable_Align = 3
            End If
    End Select
End Function

' Formatacao EXIBIDA de uma celula do Excel (DisplayFormat: inclui formatacao condicional e estilo de tabela do Excel), lida so no Excel
' (barato): Array(negrito, italico, cor automatica, cor, nome da fonte ("" = a do estilo Normal da pasta: a do tema da apresentacao fica),
' tamanho, alinhamento do PowerPoint, sem preenchimento, cor do preenchimento). Nunca propaga erro.
Private Function PptTable_Attrs(src As Range, normalFont As String) As Variant
    Dim f As Object, a(8) As Variant, nm As String
    On Error Resume Next
    Set f = src.DisplayFormat
    If f Is Nothing Then Set f = src
    Err.Clear
    a(0) = (f.Font.Bold = True): a(1) = (f.Font.Italic = True)
    a(2) = (f.Font.ColorIndex = xlColorIndexAutomatic): a(3) = CLng(f.Font.Color)
    nm = CStr(f.Font.Name)
    If StrComp(nm, normalFont, vbTextCompare) = 0 Then nm = ""
    a(4) = nm: a(5) = CSng(f.Font.Size): a(6) = PptTable_Align(f, src)
    a(7) = (f.Interior.ColorIndex = xlNone): a(8) = CLng(f.Interior.Color)
    If VarType(a(0)) <> vbBoolean Then a(0) = False                 ' misto (Null) ou ilegivel: o padrao
    If VarType(a(1)) <> vbBoolean Then a(1) = False
    If VarType(a(2)) <> vbBoolean Then a(2) = True
    If VarType(a(7)) <> vbBoolean Then a(7) = True
    If VarType(a(3)) <> vbLong Then a(3) = 0&
    If VarType(a(8)) <> vbLong Then a(8) = 0&
    If VarType(a(5)) <> vbSingle Then a(5) = 11!
    If VarType(a(6)) <> vbLong Then a(6) = 1&
    If VarType(a(4)) <> vbString Then a(4) = ""
    Err.Clear
    PptTable_Attrs = a
End Function

' Escreve na celula do slide so o que difere de prev (a formatacao que a celula ja tem: a de cima, herdada por Rows.Add, ou o padrao do estilo;
' Empty = escreve tudo). Cada leitura/escrita e uma chamada ao PowerPoint (~2,6 ms medido aqui): so as diferencas sao escritas.
' Melhor esforco: falha so registrada.
Private Sub PptTable_ApplyAttrs(cs As Object, tr As Object, a As Variant, prev As Variant)
    Dim all As Boolean, fnt As Object
    On Error Resume Next
    all = IsEmpty(prev)
    If Not all Then                                                 ' fonte igual a de cima: nem le o objeto Font
        If a(0) = prev(0) And a(1) = prev(1) And a(2) = prev(2) And a(4) = prev(4) And a(5) = prev(5) Then
            If a(2) Or a(3) = prev(3) Then GoTo NoFont
        End If
    End If
    Set fnt = tr.Font
    If all Then
        fnt.Bold = IIf(a(0), -1, 0): fnt.Italic = IIf(a(1), -1, 0)
    Else
        If a(0) <> prev(0) Then fnt.Bold = IIf(a(0), -1, 0)
        If a(1) <> prev(1) Then fnt.Italic = IIf(a(1), -1, 0)
    End If
    If a(2) Then
        If all Then
            fnt.Color.ObjectThemeColor = 13                         ' msoThemeColorText1: a cor automatica segue o tema
        ElseIf Not prev(2) Then
            fnt.Color.ObjectThemeColor = 13
        End If
    ElseIf all Then
        fnt.Color.RGB = a(3)
    ElseIf prev(2) Or a(3) <> prev(3) Then
        fnt.Color.RGB = a(3)
    End If
    If all Then
        If Len(a(4)) > 0 Then fnt.Name = a(4) Else fnt.Name = "+mn-lt"   ' "+mn-lt": a fonte do corpo do tema
    ElseIf a(4) <> prev(4) Then
        If Len(a(4)) > 0 Then fnt.Name = a(4) Else fnt.Name = "+mn-lt"
    End If
    If all Then
        fnt.Size = a(5)
    ElseIf a(5) <> prev(5) Then
        fnt.Size = a(5)
    End If
NoFont:
    If all Then
        tr.ParagraphFormat.Alignment = a(6)
    ElseIf a(6) <> prev(6) Then
        tr.ParagraphFormat.Alignment = a(6)
    End If
    If a(7) Then
        If all Then
            cs.Fill.Visible = 0
        ElseIf Not prev(7) Then
            cs.Fill.Visible = 0
        End If
    ElseIf all Then
        cs.Fill.Visible = -1: cs.Fill.Solid: cs.Fill.ForeColor.RGB = a(8)
    ElseIf prev(7) Or a(8) <> prev(8) Then
        cs.Fill.Visible = -1: cs.Fill.Solid: cs.Fill.ForeColor.RGB = a(8)
    End If
    If Err.Number <> 0 Then Errors_Log "PptTable_ApplyAttrs" & vbTab & Err.Number & vbTab & Err.Description
    Err.Clear
End Sub

' Formatacao de uma celula nova com o estilo sem preenchimento aplicado: sem negrito/italico, cor do tema, fonte do tema, sem preenchimento,
' a esquerda; tamanho desconhecido (sempre escrito).
Private Function PptTable_StyleDefaults() As Variant
    PptTable_StyleDefaults = Array(False, False, True, 0&, "", -1!, 1&, True, 0&)
End Function
' Larguras das colunas proporcionais as do Excel (total w; coluna oculta ou estreita: no minimo 2% de w). Melhor esforco.
Private Sub PptTable_ColumnWidths(t As Object, rng As Range, w As Double)
    Dim c As Long, tot As Double, cw As Double, minW As Double
    On Error Resume Next
    minW = 0.02 * w
    For c = 1 To rng.Columns.Count
        cw = rng.Columns(c).Width
        If cw < minW Then cw = minW
        tot = tot + cw
    Next
    If tot <= 0 Then Exit Sub
    For c = 1 To rng.Columns.Count
        cw = rng.Columns(c).Width
        If cw < minW Then cw = minW
        t.Columns(c).Width = w * cw / tot
    Next
    If Err.Number <> 0 Then Errors_Log "PptTable_ColumnWidths" & vbTab & Err.Number & vbTab & Err.Description
    Err.Clear
End Sub

' Fonte do estilo Normal da pasta ("" se ilegivel). Nunca propaga erro.
Private Function PptTable_NormalFont(wb As Workbook) As String
    On Error Resume Next
    PptTable_NormalFont = CStr(wb.Styles("Normal").Font.Name)
    Err.Clear
End Function

' Grava TQQ_TBLAST (ou a apaga, se passar de PPTTABLE_MAX_TAG: a proxima atualizacao reescreve os textos que diferirem). Propaga erro.
Private Sub PptTable_TagLast(shp As Object, m As Variant)
    Dim s As String
    s = PptTable_Serialize(m)
    If Len(s) > PPTTABLE_MAX_TAG Then
        Errors_Log "PptTable: matriz grande demais para a etiqueta TQQ_TBLAST (" & Len(s) & " caracteres)"
        PptSync_TagDelete shp, PPTTABLE_TAG_LAST
    ElseIf PptSync_Tag(shp, PPTTABLE_TAG_LAST) <> s Then
        shp.Tags.Add PPTTABLE_TAG_LAST, s
    End If
End Sub

' Cria no slide a tabela do intervalo rng (uma area, sem mescladas, ate 75 x 25, pasta salva como pasta do Excel; senao erro vbObjectError +
' 7120 com a mensagem). 80% da largura do slide, centralizada; textos como exibidos no Excel; formatacao inicial do Excel. Vincula pelo nome
' definido OCULTO TQQ_T_<id> (+ a marca PPTLINK_MARK) na pasta de rng (ABERTA: a acao explicita do usuario e a unica que grava nela) e pelas
' etiquetas TQQ_BOOK, TQQ_TNAME, TQQ_TREF, TQQ_TBLAST e, por ultimo, TQQ_TABLE. Transacional: falha apaga a tabela, o nome e a marca que esta
' chamada criou e sobe o erro.
Public Function PptTable_Insert(rng As Range, slide As Object) As Object
    Dim wb As Workbook, m As Variant, shp As Object, t As Object, r As Long, c As Long, sw As Double, sh As Double, w As Double, h As Double
    Dim msg As String, nm As String, hadMark As Boolean, named As Boolean, en As Long, ed As String, cs As Object, tr As Object, styled As Boolean
    Dim normalFont As String, top As Double, prevRow() As Variant, curRow() As Variant
    msg = PptTable_RefusalText(rng)
    If Len(msg) > 0 Then Err.Raise vbObjectError + PPTTABLE_ERR, "PptTable", msg
    Set wb = rng.Worksheet.Parent
    If Len(wb.Path) = 0 Then Err.Raise vbObjectError + PPTTABLE_ERR, "PptTable", PPTTEXT_MSG_UNSAVED
    If Not PptData_IsWorkbookExt(wb.FullName) Then Err.Raise vbObjectError + PPTTABLE_ERR, "PptTable", PPTTEXT_MSG_NOTBOOK
    m = PptTable_TextMatrix(rng)
    normalFont = PptTable_NormalFont(wb)
    hadMark = PptLink_HasMark(wb)
    nm = PptLink_DefineName(rng, PPTTABLE_NAME_PREFIX)          ' falha aqui: nada mudou
    named = True
    On Error GoTo Undo
    sw = slide.Parent.PageSetup.SlideWidth: sh = slide.Parent.PageSetup.SlideHeight
    If sw <= 0 Or sh <= 0 Then sw = 960: sh = 540
    w = 0.8 * sw: h = 22
    ' UMA linha primeiro; cada Rows.Add copia a formatacao da ultima linha (S6): cada celula so recebe o que difere da de cima
    Set shp = slide.Shapes.AddTable(1, rng.Columns.Count, (sw - w) / 2, (sh - h) / 2, w, 22)
    Set t = shp.Table
    On Error Resume Next
    t.ApplyStyle PPTTABLE_STYLE_GRID, False
    styled = (Err.Number = 0)
    If Not styled Then Errors_Log "PptTable: estilo sem preenchimento recusado" & vbTab & Err.Number & vbTab & Err.Description
    Err.Clear
    On Error GoTo Undo
    PptTable_ColumnWidths t, rng, w
    ReDim prevRow(1 To rng.Columns.Count)
    ReDim curRow(1 To rng.Columns.Count)
    If styled Then
        For c = 1 To rng.Columns.Count
            prevRow(c) = PptTable_StyleDefaults()
        Next
    End If
    For r = 1 To rng.Rows.Count
        If r > 1 Then t.Rows.Add
        For c = 1 To rng.Columns.Count
            curRow(c) = PptTable_Attrs(rng.Cells(r, c), normalFont)
            Set cs = t.Cell(r, c).Shape
            Set tr = cs.TextFrame.TextRange
            PptTable_ApplyAttrs cs, tr, curRow(c), prevRow(c)
            If Len(m(r, c)) > 0 Then tr.Text = m(r, c)
        Next
        For c = 1 To rng.Columns.Count
            prevRow(c) = curRow(c)
        Next
    Next
    Set cs = Nothing: Set tr = Nothing
    top = (sh - CDbl(shp.Height)) / 2                           ' as linhas crescem com o texto: centraliza de novo
    If top < 0 Then top = 0
    shp.Top = top
    PptTable_TestHook "tags"
    shp.Tags.Add PPTSYNC_TAG_BOOK, wb.FullName
    shp.Tags.Add PPTLINK_TAG_NAME, nm
    shp.Tags.Add PPTLINK_TAG_REF, PptLink_RefOf(rng)
    PptTable_TagLast shp, m
    shp.Tags.Add PPTTABLE_TAG, "1"                              ' por ultimo: so vira "tabela vinculada" com as etiquetas completas
    If PptAuto_NewLinksAuto() Then PptSync_SetAuto shp, True
    PptLinkIdx_Add shp                                          ' 1.11.0: indice de todos os vinculos (nunca propaga erro)
    Set PptTable_Insert = shp
    Exit Function
Undo:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set cs = Nothing: Set tr = Nothing: Set t = Nothing
    If Not shp Is Nothing Then shp.Delete
    If named Then wb.Names(nm).Delete
    If Not hadMark Then wb.Names(PPTLINK_MARK).Delete
    Err.Clear
    On Error GoTo 0
    Err.Raise en, "PptTable_Insert", ed
End Function

' Atualiza a tabela pela pasta wb (aberta; so LIDA: nada e gravado nela): ajusta linhas/colunas (novas herdam a formatacao da ultima; as do
' fim saem) e escreve so as celulas cuja origem mudou desde a ultima escrita (force: todas as que diferem do slide). Nunca apaga formas: so
' linhas/colunas desta tabela e textos das celulas. Resultados como PptText_Sync (R_NOCHART = intervalo nao encontrado). PowerPoint ocupado
' sobe.
' SEGURANCA (como o texto vinculado, revisao final E, I1): as etiquetas vem da apresentacao (de quem a fez). So um nome TQQ_T_ (criado pela
' acao de inserir) e seguido; o endereco de reserva so vale numa pasta com a marca PPTLINK_MARK. Senao um deck hostil leria intervalos
' arbitrarios de qualquer pasta do usuario e os gravaria no slide.
Public Function PptTable_Sync(shp As Object, wb As Workbook, force As Boolean) As Long
    Dim rng As Range, newM As Variant, oldM As Variant, t As Object, d As Collection, e As Variant, resized As Boolean, bn As Long, bd As String
    Dim nm As String, ref As String, msg As String, tr As Object, i As Long
    On Error GoTo Fail
    nm = PptSync_Tag(shp, PPTLINK_TAG_NAME)
    If Left$(nm, Len(PPTTABLE_NAME_PREFIX)) <> PPTTABLE_NAME_PREFIX Then nm = ""
    If PptLink_HasMark(wb) Then ref = PptSync_Tag(shp, PPTLINK_TAG_REF)
    If Len(nm) = 0 And Len(ref) = 0 Then PptTable_Sync = PPTSYNC_R_NOCHART: Exit Function
    Set rng = PptLink_ResolveRange(wb, nm, ref)
    If rng Is Nothing Then PptTable_Sync = PPTSYNC_R_NOCHART: Exit Function
    msg = PptTable_RefusalText(rng)
    If Len(msg) > 0 Then Errors_Log "PptTable_Sync: intervalo recusado" & vbTab & msg & vbTab & wb.Name: PptTable_Sync = PPTSYNC_R_FAIL: Exit Function
    If CLng(shp.HasTable) <> -1 Then Errors_Log "PptTable_Sync: a forma nao e uma tabela" & vbTab & CStr(shp.Name): PptTable_Sync = PPTSYNC_R_FAIL: Exit Function
    newM = PptTable_TextMatrix(rng)
    If Not force Then oldM = PptTable_Parse(PptSync_Tag(shp, PPTTABLE_TAG_LAST))
    Set t = shp.Table
    ' lacos com teto (revisao final F, M1): uma tabela que nao muda de tamanho (deck hostil, celulas mescladas no PowerPoint) nao prende o laco
    For i = 1 To PPTTABLE_MAX_ROWS + 1
        If t.Rows.Count >= rng.Rows.Count Then Exit For
        t.Rows.Add
        resized = True
    Next
    For i = 1 To PPTTABLE_MAX_ROWS + 1
        If t.Rows.Count <= rng.Rows.Count Then Exit For
        t.Rows(t.Rows.Count).Delete
        resized = True
    Next
    For i = 1 To PPTTABLE_MAX_ROWS + 1
        If t.Columns.Count >= rng.Columns.Count Then Exit For
        t.Columns.Add
        resized = True
    Next
    For i = 1 To PPTTABLE_MAX_ROWS + 1
        If t.Columns.Count <= rng.Columns.Count Then Exit For
        t.Columns(t.Columns.Count).Delete
        resized = True
    Next
    If t.Rows.Count <> rng.Rows.Count Or t.Columns.Count <> rng.Columns.Count Then
        Errors_Log "PptTable_Sync: a tabela do slide nao chegou ao tamanho do intervalo" & vbTab & t.Rows.Count & "x" & t.Columns.Count & vbTab & rng.Rows.Count & "x" & rng.Columns.Count
        PptTable_Sync = PPTSYNC_R_FAIL
        Exit Function
    End If
    If resized Then oldM = Empty                                 ' linhas no meio mudam a posicao de tudo abaixo: compara todos os textos
    Set d = PptTable_Diff(oldM, newM)
    If d.Count = 0 Then PptTable_Sync = PPTSYNC_R_SAME: Exit Function
    For Each e In d
        Set tr = t.Cell(e(0), e(1)).Shape.TextFrame.TextRange
        If CStr(tr.Text) <> CStr(newM(e(0), e(1))) Then tr.Text = newM(e(0), e(1))
    Next
    Set tr = Nothing
    PptTable_TagLast shp, newM
    PptTable_Sync = PPTSYNC_R_OK
    Exit Function
Fail:
    bn = Err.Number: bd = Err.Description
    On Error GoTo -1
    On Error GoTo 0                                     ' sem isto o GoTo -1 rearma o Fail e o Raise voltaria para ca (laco)
    If Ppt_IsBusyError(bn) Then Err.Raise bn, "PptTable_Sync", bd
    Errors_Log "PptTable_Sync" & vbTab & bn & vbTab & bd
    PptTable_Sync = PPTSYNC_R_FAIL
End Function