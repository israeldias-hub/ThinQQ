Attribute VB_Name = "modSettings"
Option Explicit
' Configurações por usuário em HKCU\Software\VB and VBA Program Settings\ThinQQ\Updater
Private Const REG_APP As String = "ThinQQ"
Private Const REG_SECTION As String = "Updater"
Public Const KEY_UPDATE_SOURCE As String = "UpdateSource"
Public Const KEY_LAST_CHECK As String = "LastCheck"
Public Const KEY_SKIPPED As String = "SkippedVersion"
Public Const KEY_PENDING_LOADER As String = "PendingLoaderPath"
Public Const KEY_INTERVAL As String = "CheckIntervalHours"
Public Const KEY_LAST_NOTIFIED_CORE As String = "LastNotifiedCore"
Public Const KEY_PENDING_LOADER_SHA As String = "PendingLoaderSha256"
Public Const KEY_LAST_FAILURE As String = "LastFailure"
Public Const KEY_ALLOW_UNVERIFIED As String = "AllowUnverified"   ' "1" = aceita manifesto sem SHA-256 (decisao do usuario)
Public Const DEFAULT_UPDATE_SOURCE As String = "https://raw.githubusercontent.com/israeldias-hub/ThinQQ/main/release/manifest.json"

Public Function Settings_Get(key As String, Optional defaultValue As String = "") As String
    Settings_Get = GetSetting(REG_APP, REG_SECTION, key, defaultValue)
End Function

Public Sub Settings_Set(key As String, value As String)
    SaveSetting REG_APP, REG_SECTION, key, value
End Sub

Public Sub Settings_Delete(key As String)
    On Error Resume Next
    DeleteSetting REG_APP, REG_SECTION, key
End Sub

Public Function Settings_UpdateSource() As String
    Settings_UpdateSource = Settings_Get(KEY_UPDATE_SOURCE, DEFAULT_UPDATE_SOURCE)
End Function
