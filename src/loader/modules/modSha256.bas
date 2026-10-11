Attribute VB_Name = "modSha256"
Option Explicit
' SHA-256 em VBA puro: sem processo filho (nada de certutil), sem Declare, sem LongPtr (vale em Office 32 e 64 bits).
' Aritmetica de 32 bits sem sinal emulada com Long: somas feitas em Double e reduzidas modulo 2^32; as funcoes
' Sigma/sigma (rotacoes e deslocamentos, todas lineares sobre XOR) usam tabelas de 256 entradas por byte.
'
' ATENCAO: ESTE ARQUIVO EXISTE NO LOADER (src/loader/modules) E NO CORE (src/core/modules) COM O MESMO CONTEUDO
' (o core nao pode depender do loader, que e trocado em execucao). Mantenha os dois IGUAIS: tests/ps/Test-VbaLint.ps1
' falha o build se divergirem, e os dois projetos rodam os mesmos vetores de teste.
'
' API:
'   Sha256_File(path)            hex minusculo; levanta erro claro se o arquivo nao puder ser lido (nunca devolve "")
'   Sha256_Bytes(data(), length) hex minusculo dos primeiros 'length' bytes (length = 0 e valido)

Private Const CHUNK_BYTES As Long = 262144      ' multiplo de 64
Private Const TWO32 As Double = 4294967296#
Private Const INV_TWO32 As Double = 2.3283064365386963E-10

Private mReady As Boolean
Private mH(0 To 7) As Long
Private mK(0 To 63) As Long
Private mW(0 To 63) As Long
' tabelas por byte: Sa = Sigma0 (2,13,22), Sb = Sigma1 (6,11,25), Sc = sigma0 (7,18,>>3), Sd = sigma1 (17,19,>>10)
Private mSa0(0 To 255) As Long, mSa1(0 To 255) As Long, mSa2(0 To 255) As Long, mSa3(-128 To 127) As Long
Private mSb0(0 To 255) As Long, mSb1(0 To 255) As Long, mSb2(0 To 255) As Long, mSb3(-128 To 127) As Long
Private mSc0(0 To 255) As Long, mSc1(0 To 255) As Long, mSc2(0 To 255) As Long, mSc3(-128 To 127) As Long
Private mSd0(0 To 255) As Long, mSd1(0 To 255) As Long, mSd2(0 To 255) As Long, mSd3(-128 To 127) As Long

Public Function Sha256_File(path As String) As String
    Dim stm As Object, buf() As Byte, total As Long, done As Long, n As Long, nb As Long, dummy(0 To 0) As Byte
    Dim en As Long, ed As String, opened As Boolean
    On Error GoTo Fail
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open: opened = True
    stm.LoadFromFile path                       ' levanta erro se o arquivo nao existe / nao pode ser lido
    total = stm.Size
    ShaInit
    If total = 0 Then Sha256_File = ShaFinish(dummy, 0, 0, 0)
    Do While done < total
        n = total - done
        If n > CHUNK_BYTES Then n = CHUNK_BYTES
        buf = stm.Read(n)
        If UBound(buf) - LBound(buf) + 1 <> n Then Err.Raise vbObjectError + 1005, "Sha256", "Leitura incompleta do arquivo."
        nb = n \ 64
        ShaBlocks buf, 0, nb
        done = done + n
        If done >= total Then Sha256_File = ShaFinish(buf, nb * 64, n - nb * 64, total)
    Loop
    stm.Close: opened = False
    If Len(Sha256_File) <> 64 Then Err.Raise vbObjectError + 1006, "Sha256", "Resultado de SHA-256 invalido."
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    If opened Then stm.Close
    On Error GoTo 0
    Sha256_File = ""
    Err.Raise vbObjectError + 1003, "Sha256", "Falha ao calcular o SHA-256 de " & path & " (" & ed & ")"
End Function

Public Function Sha256_Bytes(data() As Byte, ByVal length As Long) As String
    Dim nb As Long
    If length < 0 Then Err.Raise 5, "Sha256", "Tamanho negativo."
    If length > UBound(data) - LBound(data) + 1 Then Err.Raise 9, "Sha256", "Tamanho maior que o vetor."
    ShaInit
    nb = length \ 64
    If nb > 0 Then ShaBlocks data, 0, nb
    Sha256_Bytes = ShaFinish(data, nb * 64, length - nb * 64, length)
End Function

' ---------------------------------------------------------------- internos

Private Function ToLong(ByVal u As Double) As Long
    If u >= 2147483648# Then ToLong = CLng(u - TWO32) Else ToLong = CLng(u)
End Function

Private Function ToUnsigned(ByVal x As Long) As Double
    If x < 0 Then ToUnsigned = CDbl(x) + TWO32 Else ToUnsigned = x
End Function

' So para montar as tabelas (lento, uma vez).
Private Function SlowRotR(ByVal x As Long, ByVal n As Long) As Long
    Dim u As Double, q As Double
    u = ToUnsigned(x)
    q = Int(u / 2 ^ n)
    SlowRotR = ToLong(q + (u - q * 2 ^ n) * 2 ^ (32 - n))
End Function

Private Function SlowShR(ByVal x As Long, ByVal n As Long) As Long
    SlowShR = ToLong(Int(ToUnsigned(x) / 2 ^ n))
End Function

Private Sub ShaInitTables()
    Dim v As Long, b As Long, x As Long, k As Long
    Dim kk As Variant
    If mReady Then Exit Sub
    For k = 0 To 3
        For v = IIf(k = 3, -128, 0) To IIf(k = 3, 127, 255)
            b = v And 255
            x = ToLong(CDbl(b) * 256# ^ k)
            Select Case k
                Case 0
                    mSa0(v) = SlowRotR(x, 2) Xor SlowRotR(x, 13) Xor SlowRotR(x, 22)
                    mSb0(v) = SlowRotR(x, 6) Xor SlowRotR(x, 11) Xor SlowRotR(x, 25)
                    mSc0(v) = SlowRotR(x, 7) Xor SlowRotR(x, 18) Xor SlowShR(x, 3)
                    mSd0(v) = SlowRotR(x, 17) Xor SlowRotR(x, 19) Xor SlowShR(x, 10)
                Case 1
                    mSa1(v) = SlowRotR(x, 2) Xor SlowRotR(x, 13) Xor SlowRotR(x, 22)
                    mSb1(v) = SlowRotR(x, 6) Xor SlowRotR(x, 11) Xor SlowRotR(x, 25)
                    mSc1(v) = SlowRotR(x, 7) Xor SlowRotR(x, 18) Xor SlowShR(x, 3)
                    mSd1(v) = SlowRotR(x, 17) Xor SlowRotR(x, 19) Xor SlowShR(x, 10)
                Case 2
                    mSa2(v) = SlowRotR(x, 2) Xor SlowRotR(x, 13) Xor SlowRotR(x, 22)
                    mSb2(v) = SlowRotR(x, 6) Xor SlowRotR(x, 11) Xor SlowRotR(x, 25)
                    mSc2(v) = SlowRotR(x, 7) Xor SlowRotR(x, 18) Xor SlowShR(x, 3)
                    mSd2(v) = SlowRotR(x, 17) Xor SlowRotR(x, 19) Xor SlowShR(x, 10)
                Case Else
                    mSa3(v) = SlowRotR(x, 2) Xor SlowRotR(x, 13) Xor SlowRotR(x, 22)
                    mSb3(v) = SlowRotR(x, 6) Xor SlowRotR(x, 11) Xor SlowRotR(x, 25)
                    mSc3(v) = SlowRotR(x, 7) Xor SlowRotR(x, 18) Xor SlowShR(x, 3)
                    mSd3(v) = SlowRotR(x, 17) Xor SlowRotR(x, 19) Xor SlowShR(x, 10)
            End Select
        Next
    Next
    kk = Array(&H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5)
    For k = 0 To 7: mK(k) = kk(k): Next
    kk = Array(&HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174)
    For k = 0 To 7: mK(8 + k) = kk(k): Next
    kk = Array(&HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA)
    For k = 0 To 7: mK(16 + k) = kk(k): Next
    kk = Array(&H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967)
    For k = 0 To 7: mK(24 + k) = kk(k): Next
    kk = Array(&H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85)
    For k = 0 To 7: mK(32 + k) = kk(k): Next
    kk = Array(&HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070)
    For k = 0 To 7: mK(40 + k) = kk(k): Next
    kk = Array(&H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3)
    For k = 0 To 7: mK(48 + k) = kk(k): Next
    kk = Array(&H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2)
    For k = 0 To 7: mK(56 + k) = kk(k): Next
    mReady = True
End Sub

Private Sub ShaInit()
    ShaInitTables
    mH(0) = &H6A09E667: mH(1) = &HBB67AE85: mH(2) = &H3C6EF372: mH(3) = &HA54FF53A
    mH(4) = &H510E527F: mH(5) = &H9B05688C: mH(6) = &H1F83D9AB: mH(7) = &H5BE0CD19
End Sub

' Processa nBlocks blocos de 64 bytes de buf a partir de off (0-based dentro do vetor).
Private Sub ShaBlocks(buf() As Byte, ByVal off As Long, ByVal nBlocks As Long)
    Dim blk As Long, t As Long, p As Long, x As Long, y As Long, s0 As Long, s1 As Long, tt As Double
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, h As Long, ch As Long, mj As Long
    Dim lb As Long, na As Long, ne As Long, te As Double
    lb = LBound(buf)
    For blk = 0 To nBlocks - 1
        p = lb + off + blk * 64
        For t = 0 To 15
            mW(t) = ((buf(p) And &H7F&) * &H1000000) Or (CLng(buf(p + 1)) * &H10000) Or (CLng(buf(p + 2)) * &H100&) Or buf(p + 3)
            If buf(p) > 127 Then mW(t) = mW(t) Or &H80000000
            p = p + 4
        Next
        For t = 16 To 63
            x = mW(t - 15): y = mW(t - 2)
            s0 = mSc0(x And &HFF&) Xor mSc1((x And &HFF00&) \ &H100&) Xor mSc2((x And &HFF0000) \ &H10000) Xor mSc3((x And &HFF000000) \ &H1000000)
            s1 = mSd0(y And &HFF&) Xor mSd1((y And &HFF00&) \ &H100&) Xor mSd2((y And &HFF0000) \ &H10000) Xor mSd3((y And &HFF000000) \ &H1000000)
            tt = CDbl(mW(t - 16)) + s0 + mW(t - 7) + s1
            tt = tt - Int(tt * INV_TWO32) * TWO32
            If tt >= 2147483648# Then mW(t) = CLng(tt - TWO32) Else mW(t) = CLng(tt)
        Next
        a = mH(0): b = mH(1): c = mH(2): d = mH(3): e = mH(4): f = mH(5): g = mH(6): h = mH(7)
        For t = 0 To 63
            s1 = mSb0(e And &HFF&) Xor mSb1((e And &HFF00&) \ &H100&) Xor mSb2((e And &HFF0000) \ &H10000) Xor mSb3((e And &HFF000000) \ &H1000000)
            ch = ((f Xor g) And e) Xor g
            tt = CDbl(h) + s1 + ch + mK(t) + mW(t)
            s0 = mSa0(a And &HFF&) Xor mSa1((a And &HFF00&) \ &H100&) Xor mSa2((a And &HFF0000) \ &H10000) Xor mSa3((a And &HFF000000) \ &H1000000)
            mj = ((a Or b) And c) Or (a And b)
            te = tt + d                         ' e novo = d + T1
            te = te - Int(te * INV_TWO32) * TWO32
            If te >= 2147483648# Then ne = CLng(te - TWO32) Else ne = CLng(te)
            tt = tt + s0 + mj
            tt = tt - Int(tt * INV_TWO32) * TWO32
            If tt >= 2147483648# Then na = CLng(tt - TWO32) Else na = CLng(tt)
            h = g: g = f: f = e: e = ne: d = c: c = b: b = a: a = na
        Next
        mH(0) = AddWrap(mH(0), a): mH(1) = AddWrap(mH(1), b): mH(2) = AddWrap(mH(2), c): mH(3) = AddWrap(mH(3), d)
        mH(4) = AddWrap(mH(4), e): mH(5) = AddWrap(mH(5), f): mH(6) = AddWrap(mH(6), g): mH(7) = AddWrap(mH(7), h)
    Next
End Sub

Private Function AddWrap(ByVal a As Long, ByVal b As Long) As Long
    Dim u As Double
    u = CDbl(a) + b
    u = u - Int(u * INV_TWO32) * TWO32
    AddWrap = ToLong(u)
End Function

' Completa o calculo: tail = bytes restantes (tailLen < 64) a partir de tailOff; totalBytes = tamanho total da mensagem.
Private Function ShaFinish(tail() As Byte, ByVal tailOff As Long, ByVal tailLen As Long, ByVal totalBytes As Long) As String
    Dim pad(0 To 127) As Byte, i As Long, nb As Long, base As Long, bits As Double, hiW As Double, loW As Double, x As Double, k As Long
    Dim out As String
    For i = 0 To tailLen - 1
        pad(i) = tail(LBound(tail) + tailOff + i)
    Next
    pad(tailLen) = &H80
    If tailLen < 56 Then nb = 1 Else nb = 2
    base = nb * 64 - 8
    bits = CDbl(totalBytes) * 8#
    hiW = Int(bits / TWO32)
    loW = bits - hiW * TWO32
    x = loW
    For k = 7 To 4 Step -1
        pad(base + k) = CByte(x - Int(x / 256#) * 256#)
        x = Int(x / 256#)
    Next
    x = hiW
    For k = 3 To 0 Step -1
        pad(base + k) = CByte(x - Int(x / 256#) * 256#)
        x = Int(x / 256#)
    Next
    ShaBlocks pad, 0, nb
    For i = 0 To 7
        out = out & Right$("0000000" & Hex$(mH(i)), 8)
    Next
    ShaFinish = LCase$(out)
End Function
