Attribute VB_Name = "TestVersion"
Option Explicit

Public Sub Test_Version_Valid()
    AssertTrue IsValidVersion("1.0.0")
    AssertTrue IsValidVersion("10.20.30")
End Sub

Public Sub Test_Version_Invalid()
    AssertFalse IsValidVersion("")
    AssertFalse IsValidVersion("1.0")
    AssertFalse IsValidVersion("v1.0.0")
    AssertFalse IsValidVersion("1.0.0-beta")
    AssertFalse IsValidVersion("1..0")
End Sub

Public Sub Test_Version_Compare()
    AssertEqual 1, CompareVersions("1.2.10", "1.2.9")
    AssertEqual -1, CompareVersions("1.9.0", "1.10.0")
    AssertEqual 0, CompareVersions("2.0.0", "2.0.0")
    AssertEqual 1, CompareVersions("0.0.1", "")
    AssertEqual -1, CompareVersions("abc", "0.0.1")
End Sub
