Attribute VB_Name = "TestJsonLite"
Option Explicit

Public Sub Test_Json_String()
    AssertEqual "1.2.3", JsonGetString("{""coreVersion"": ""1.2.3"", ""x"": 1}", "coreVersion")
End Sub

Public Sub Test_Json_Escapes()
    AssertEqual "a" & vbLf & "b ""c"" \ é /", JsonGetString("{""notes"":""a\nb \""c\"" \\ \u00e9 \/""}", "notes")
End Sub

Public Sub Test_Json_MissingKey_ReturnsEmpty()
    AssertEqual "", JsonGetString("{""a"":""b""}", "coreUrl")
End Sub

Public Sub Test_Json_Number_And_Bool()
    AssertEqual 2, JsonGetNumber("{""schemaVersion"":2,""a"":""b""}", "schemaVersion")
    AssertEqual "true", JsonGetString("{""flag"": true}", "flag")
End Sub

Public Sub Test_Json_KeyPrefixNotConfused()
    AssertEqual "B", JsonGetString("{""coreUrlX"":""A"",""coreUrl"":""B""}", "coreUrl")
End Sub

Public Sub Test_Json_MultilinePretty()
    Dim j As String
    j = "{" & vbCrLf & "  ""coreVersion"": ""1.0.0""," & vbCrLf & "  ""notes"": ""x""" & vbCrLf & "}"
    AssertEqual "1.0.0", JsonGetString(j, "coreVersion")
    AssertEqual "x", JsonGetString(j, "notes")
End Sub

Public Sub Test_Json_BadUnicodeEscape_DoesNotRaise()
    Dim bs As String, r As String
    bs = Chr$(92)
    r = JsonGetString("{""a"":""x" & bs & "uZZZZy""}", "a")
    AssertEqual "x?ZZZZy", r
End Sub

Public Sub Test_Json_TruncatedUnicodeEscape_DoesNotRaise()
    Dim bs As String
    bs = Chr$(92)
    AssertEqual "x?12", JsonGetString("{""a"":""x" & bs & "u12", "a")
    AssertEqual "x?", JsonGetString("{""a"":""x" & bs & "u", "a")
End Sub
