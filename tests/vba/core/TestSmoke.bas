Attribute VB_Name = "TestSmoke"
Option Explicit

Public Sub Test_Core_Ping()
    AssertEqual "pong", Core_Ping()
End Sub

Public Sub Test_BuildInfo_IsTestConfig()
    AssertEqual "Test", BUILD_CONFIG
    AssertEqual ".test", BUILD_SUFFIX
End Sub
