Attribute VB_Name = "modSettings"
Option Explicit
' Configurações por usuário em HKCU\Software\VB and VBA Program Settings\ThinQQ\Updater (build Test: ...\ThinQQ.Test\Updater, ver Settings_AppName)
Private Const REG_APP As String = "ThinQQ"
Private Const REG_APP_TEST As String = "ThinQQ.Test"
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
Public Const KEY_SWAP_IN_PROGRESS As String = "SwapInProgress"       ' timestamp gravado pelo core durante a troca do loader
Public Const KEY_PENDING_FAIL_NOTIFIED As String = "PendingLoaderFailNotified"
Public Const KEY_PPAM_VERSION As String = "PpamVersion"             ' versao do ThinQQ.ppam instalado (instalador e atualizador gravam)
Public Const KEY_PPAM_AUTO As String = "PpamAutoUpdate"             ' "0" = nao atualizar o suplemento do PowerPoint (padrao "1")
Public Const KEY_PENDING_PPAM As String = "PendingPpam"             ' "1" = ThinQQ.ppam.new aguardando o PowerPoint liberar o arquivo
Public Const KEY_PENDING_PPAM_SHA As String = "PendingPpamSha256"
Public Const KEY_PENDING_PPAM_VER As String = "PendingPpamVersion"
Public Const DEFAULT_UPDATE_SOURCE As String = "https://raw.githubusercontent.com/israeldias-hub/ThinQQ/main/release/manifest.json"

' Nome do aplicativo no registro. O build Test usa um ramo proprio (ThinQQ.Test): os testes nunca leem nem gravam as configuracoes reais do
' atualizador, nem mesmo se o Excel morrer no meio de um teste. O core usa o mesmo nome (Util_UpdaterRegApp) na troca do loader.
Public Function Settings_AppName() As String
    If BUILD_CONFIG = "Test" Then Settings_AppName = REG_APP_TEST Else Settings_AppName = REG_APP
End Function

Public Function Settings_Get(key As String, Optional defaultValue As String = "") As String
    Settings_Get = GetSetting(Settings_AppName(), REG_SECTION, key, defaultValue)
End Function

Public Sub Settings_Set(key As String, value As String)
    SaveSetting Settings_AppName(), REG_SECTION, key, value
End Sub

Public Sub Settings_Delete(key As String)
    On Error Resume Next
    DeleteSetting Settings_AppName(), REG_SECTION, key
End Sub

Public Function Settings_UpdateSource() As String
    Settings_UpdateSource = Settings_Get(KEY_UPDATE_SOURCE, DEFAULT_UPDATE_SOURCE)
End Function
