Attribute VB_Name = "TestSha256"
Option Explicit
' ATENCAO: mesmo arquivo em tests/vba/loader e tests/vba/core (mesmos vetores nos dois projetos; o modulo modSha256
' tambem e identico nos dois, verificado por Test-VbaLint.ps1). Vetores calculados com .NET SHA256 fora do VBA.
' Padrao de bytes: b(i) = (i*7 + (i\256)*13 + 13) And 255.

Private Function Pattern(n As Long) As Byte()
    Dim b() As Byte, i As Long
    ReDim b(0 To IIf(n > 0, n - 1, 0))
    For i = 0 To n - 1
        b(i) = CByte((i * 7 + (i \ 256) * 13 + 13) And 255)
    Next
    Pattern = b
End Function

Private Function Expected(n As Long) As String
    Select Case n
        Case 0: Expected = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
        Case 1: Expected = "9d1e0e2d9459d06523ad13e28a4093c2316baafe7aec5b25f30eba2e113599c4"
        Case 55: Expected = "764c574722e6e2ccaa5422f8ec731111ac72ff7039793148623e56b75a32c11f"
        Case 56: Expected = "43fbbe48a6796cb7414a92cd785d9f4a976c2f70fc59c60a309f95e3022db77a"
        Case 57: Expected = "e038a2370dbd74c3c8b89b95e7c351fec4821e3415f7aef3a0925215bc6ff953"
        Case 63: Expected = "c309180feace42e90107301813aef6f309cac604e831b3fd9692a3298aa6da54"
        Case 64: Expected = "3a38aed112131d75fc0e636437f5b675c83c01ade88d99f6b6c54b0d6129174f"
        Case 65: Expected = "2ee4bedec261c1561dafa7ba28e4e3ece281bc0f51afca40b83b3a2a7c41a050"
        Case 119: Expected = "0a70cbf85ea376617e4bfad11040a9559638f8ceb57844a901573674578af539"
        Case 120: Expected = "7d3fd765bd3d4a0587f5bab94200b1d38b23398b94544ff5257f695b2227918f"
        Case 127: Expected = "ff998a2ad3412188b7ba531324bf977b22e77aa3b1befb11c699bf2a14959ee7"
        Case 128: Expected = "8b94fd8b7db8b1ef29c089c16389697a057310b7c739c1ad844e9be970f5cfd6"
        Case 129: Expected = "22afcb610b1282b24536c87a33acc00a80c720c9d3509960ae11a9bd87501330"
        Case 1000: Expected = "00c741af63dd53fd24e5f102498055484738021c1cc05f2b3ba5c54c70184a64"
        Case 65535: Expected = "14a4e21bece95756e5fc775bbdcbd4f2ee0eea13598c8a852d5e87c4bcf6169e"
        Case 65536: Expected = "95e4af8c98bb2784f50e97b8c5582b5556c84a25fb00c40874f53e814c55d6f5"
        Case 65537: Expected = "47efe4e3402c028e6f1cf63386dbf6d6c2cbbad62071615699fb82f1c7aa6f39"
        Case 262143: Expected = "191f6006abf5729bda3f33ec02422b1375ccac23b78b878e4612809e5865facd"
        Case 262144: Expected = "88695f56b053479147a4cd90eef4e499d9083b7a24ec4bfa501d6acb92a1eaab"
        Case 262145: Expected = "8f90c78a703cc821214b66b6b46bb24c142008890c51c1467d9f0f6b09c6a905"
        Case 524289: Expected = "5aee211e2211a83d94a2b70142816b0f3848d35d1e8d77b1708fb68a7283bf72"
        Case 1048576: Expected = "68bcc32703bb70f8126a951b5789d74772cb14cf0f206180ef38635cce2cdc6e"
    End Select
End Function

Private Sub WriteBytes(path As String, b() As Byte, n As Long)
    Dim stm As Object, f As Integer
    If n = 0 Then
        f = FreeFile: Open path For Output As #f: Close #f
        Exit Sub
    End If
    ReDim Preserve b(0 To n - 1)
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open: stm.Write b: stm.SaveToFile path, 2: stm.Close
End Sub

Private Function Ascii(s As String) As Byte()
    Ascii = StrConv(s, vbFromUnicode)
End Function

Public Sub Test_Sha256_KnownVectors_Bytes()
    Dim b() As Byte
    b = Ascii("abc")
    AssertEqual "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", Sha256_Bytes(b, 3)
    b = Ascii("abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq")
    AssertEqual "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1", Sha256_Bytes(b, 56), "vetor NIST de 448 bits"
    b = Pattern(0)
    AssertEqual Expected(0), Sha256_Bytes(b, 0), "vazio"
End Sub

Public Sub Test_Sha256_BlockBoundaries_Bytes()
    Dim v As Variant, b() As Byte, n As Long
    For Each v In Array(1, 55, 56, 57, 63, 64, 65, 119, 120, 127, 128, 129, 1000, 65535, 65536, 65537)
        n = CLng(v)
        b = Pattern(n)
        AssertEqual Expected(n), Sha256_Bytes(b, n), "bytes n=" & n
    Next
End Sub

Public Sub Test_Sha256_File_BoundariesAndChunks()
    Dim v As Variant, b() As Byte, n As Long, d As String
    d = TempDir()
    For Each v In Array(0, 1, 55, 56, 63, 64, 65, 128, 129, 65536, 262143, 262144, 262145, 524289)
        n = CLng(v)
        b = Pattern(n)
        WriteBytes d & "\f.bin", b, n
        AssertEqual Expected(n), Sha256_File(d & "\f.bin"), "arquivo n=" & n
    Next
End Sub

Public Sub Test_Sha256_File_MillionA_And_Speed()
    Dim b() As Byte, i As Long, d As String, t As Double, h As String
    d = TempDir()
    ReDim b(0 To 999999)
    For i = 0 To 999999: b(i) = 97: Next
    WriteBytes d & "\a.bin", b, 1000000
    t = Timer
    h = Sha256_File(d & "\a.bin")
    t = Timer - t
    AssertEqual "cdc76e5c9914fb9281a1c7e284d73e67f1809a48a497200e046d39ccc7112cd0", h, "um milhao de 'a'"
    AssertTrue t < 5, "1 MB deve levar bem menos de 5 s (levou " & Format$(t, "0.00") & " s)"
End Sub

Public Sub Test_Sha256_File_OneMegabytePattern()
    Dim b() As Byte, d As String
    d = TempDir()
    b = Pattern(1048576)
    WriteBytes d & "\p.bin", b, 1048576
    AssertEqual Expected(1048576), Sha256_File(d & "\p.bin")
End Sub

Public Sub Test_Sha256_File_Missing_RaisesAndNeverReturnsEmpty()
    Dim h As String, en As Long
    h = "x"
    On Error Resume Next
    h = Sha256_File(TempDir() & "\nao-existe.bin")
    en = Err.Number
    On Error GoTo 0
    AssertTrue en <> 0, "arquivo ausente deve levantar erro"
    AssertEqual "x", h, "nao pode devolver hash (nem vazio) para arquivo ausente"
End Sub

Public Sub Test_Sha256_Bytes_RejectsBadLength()
    Dim b() As Byte, en As Long
    b = Pattern(4)
    On Error Resume Next
    Sha256_Bytes b, 99
    en = Err.Number
    On Error GoTo 0
    AssertTrue en <> 0
End Sub
