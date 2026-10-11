# Smoke test em Excel REAL (instancia propria e invisivel) do caminho de producao da troca do loader:
# o core executa Core_ApplyPendingLoader, que fecha o loader Release antigo (dispara Workbook_BeforeClose ->
# Loader_Shutdown), troca o arquivo e reabre. O CORE precisa continuar aberto. Os testes VBA unitarios nao
# cobrem isso (o build de Teste sai cedo em Workbook_Open/BeforeClose).
#
# Uso (a partir da raiz do repositorio):
#   pwsh build/build.ps1 -Target all -Configuration Release
#   pwsh build/build.ps1 -Target all -Configuration Test     (ThinQQ.test.xlam = loader "novo" inofensivo)
#   pwsh tests/ps/Test-LoaderSwap.ps1
# Seguranca: nao instala nada, nao toca no registro de add-ins nem em %LOCALAPPDATA%\ThinQQ: o Excel iniciado aqui
# herda um LOCALAPPDATA redirecionado para uma pasta temporaria (Loader_SetRootOverride e inerte em Release, M1).
# Valores do registro do atualizador (HKCU\Software\VB and VBA Program Settings\ThinQQ\Updater)
# sao salvos e restaurados. Usa SO o Excel iniciado aqui (New-ExcelApp valida PID novo; watchdog por PID).
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [int]$TimeoutSec = 120
)
. (Join-Path $RepoRoot 'build\lib\Build-Common.ps1')
Assert-VbomTrusted
$script:fail = 0
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m"; $script:fail++ } }

$loaderSrc     = Join-Path $DistDir 'ThinQQ.xlam'
$coreSrc       = Join-Path $DistDir 'ThinQQ.core.xlam'
$loaderTestSrc = Join-Path $DistDir 'ThinQQ.test.xlam'
foreach ($f in $loaderSrc, $coreSrc, $loaderTestSrc) {
    if (-not (Test-Path $f)) { Write-Host "not ok - artefato ausente: $f (rode build.ps1 em Release e em Test)"; exit 1 }
}

$regPath = 'HKCU:\Software\VB and VBA Program Settings\ThinQQ\Updater'
$keys = 'PendingLoaderPath', 'PendingLoaderSha256', 'SwapInProgress', 'LastCheck', 'PendingLoaderFailNotified', 'LastFailure'
function Get-RegVal([string]$name) {
    try { return (Get-ItemProperty -Path $regPath -Name $name -ErrorAction Stop).$name } catch { return $null }
}
function Get-DocProperty($wb, [string]$Name) {
    $bf = [System.Reflection.BindingFlags]
    $props = [System.__ComObject].InvokeMember('BuiltinDocumentProperties', $bf::GetProperty, $null, $wb, $null)
    $prop = [System.__ComObject].InvokeMember('Item', $bf::GetProperty, $null, $props, @($Name))
    return [System.__ComObject].InvokeMember('Value', $bf::GetProperty, $null, $prop, $null)
}
# Get-FileHash falha em arquivo aberto pelo Excel (FileShare.Read); aqui abre com ReadWrite.
function Get-OpenFileSha([string]$path) {
    $fs = New-Object IO.FileStream($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try { $h = [Security.Cryptography.SHA256]::Create().ComputeHash($fs); return (($h | ForEach-Object { $_.ToString('x2') }) -join '') } finally { $fs.Dispose() }
}
# Pastas de trabalho de add-in (.xlam) NAO aparecem na enumeracao de Workbooks; so por nome.
function Get-Wb([string]$name) {
    try { return $script:xl.Workbooks.Item($name) } catch { return $null }
}
$backup = @{}
foreach ($k in $keys) { $backup[$k] = Get-RegVal $k }
function Restore-Reg {
    foreach ($k in $keys) {
        if ($null -ne $backup[$k]) { New-Item -Path $regPath -Force | Out-Null; Set-ItemProperty -Path $regPath -Name $k -Value $backup[$k] }
        else { Remove-ItemProperty -Path $regPath -Name $k -ErrorAction SilentlyContinue }
    }
}

$tmp = Join-Path ([IO.Path]::GetTempPath()) ("thinqq_swap_" + [guid]::NewGuid().ToString('N').Substring(0, 8))
$localApp = Join-Path $tmp 'localapp'
$rootDir  = Join-Path $localApp 'ThinQQ'
$instDir  = Join-Path $tmp 'inst'
$stageDir = Join-Path $tmp 'stage'
$prevLocalApp = $env:LOCALAPPDATA
$verDir   = Join-Path $rootDir 'versions\1.0.0'
New-Item -ItemType Directory -Force $rootDir, $instDir, $stageDir, $verDir | Out-Null
$env:LOCALAPPDATA = $localApp        # herdado pelo Excel desta execucao (Loader_Root = %LOCALAPPDATA%\ThinQQ)
$target  = Join-Path $instDir 'ThinQQ.xlam'
$staged  = Join-Path $stageDir 'ThinQQ.xlam'
$coreTmp = Join-Path $verDir 'ThinQQ.core.xlam'
Copy-Item $coreSrc $coreTmp
Copy-Item $loaderSrc $target
Copy-Item $loaderTestSrc $staged

$xl = $null; $wd = $null
try {
    $xl = New-ExcelApp
    $wd = Start-ExcelWatchdog $xl $TimeoutSec
    $xl.AutomationSecurity = 1            # msoAutomationSecurityLow: macros do Release rodam sem prompt nesta instancia
    $xl.EnableEvents = $false

    # 1) Loader "novo" = build de TESTE (Workbook_Open/BeforeClose saem cedo: reabri-lo nao inicia um loader de
    #    producao contra o %LOCALAPPDATA% real), com uma propriedade de documento distinguivel.
    $s = $xl.Workbooks.Open($staged)
    Set-DocProperty $s 'Comments' 'swap-test'
    $s.Save(); $s.Close($false)
    $stagedHash = (Get-FileHash $staged -Algorithm SHA256).Hash.ToLower()
    $oldHash = (Get-FileHash $target -Algorithm SHA256).Hash.ToLower()
    Check ($stagedHash -ne $oldHash) 'loader pendente difere do instalado'

    # 2) Loader de PRODUCAO (Release) aberto SEM eventos (evita o Loader_Startup contra a pasta real); redireciona a
    #    raiz e abre o core temporario pelo mesmo caminho de codigo do Startup (Loader_OpenCore).
    $lw = $xl.Workbooks.Open($target)
    $opened = $xl.Run("'ThinQQ.xlam'!Loader_OpenCore", $coreTmp)
    Check ($opened -eq $true) 'Loader_OpenCore abriu o core Release temporario'
    Check ($xl.Run("'ThinQQ.core.xlam'!Core_Ping") -eq 'pong') 'Core_Ping responde antes da troca'
    Check ($xl.Run("'ThinQQ.core.xlam'!Core_LoaderPath") -eq $target) 'core registrou o loader instalado'
    $xl.EnableEvents = $true              # daqui em diante Workbook_BeforeClose do loader Release dispara de verdade

    # 2b) Controle do guarda no loader Release: com a flag de troca recente, Loader_Shutdown NAO fecha o core;
    #     sem flag, a decisao e fechar. (Obs.: neste fluxo de automacao o Workbook_BeforeClose do add-in nao foi
    #     observado disparando em Workbook.Close; por isso o guarda e exercitado chamando Loader_Shutdown direto.)
    New-Item -Path $regPath -Force | Out-Null
    Set-ItemProperty -Path $regPath -Name 'SwapInProgress' -Value (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Check ($xl.Run("'ThinQQ.xlam'!Loader_ShouldCloseCoreOnShutdown") -eq $false) 'flag recente => loader nao deve fechar o core'
    $xl.Run("'ThinQQ.xlam'!Loader_Shutdown") | Out-Null
    Check ($null -ne (Get-Wb 'ThinQQ.core.xlam')) 'Loader_Shutdown com flag recente deixa o core aberto'
    Remove-ItemProperty -Path $regPath -Name 'SwapInProgress' -ErrorAction SilentlyContinue
    Check ($xl.Run("'ThinQQ.xlam'!Loader_ShouldCloseCoreOnShutdown") -eq $true) 'sem flag => loader fecha o core normalmente'

    # 3) Pendente + SHA no registro; executa a troca como o OnTime faria (loader fora da pilha).
    New-Item -Path $regPath -Force | Out-Null
    Set-ItemProperty -Path $regPath -Name 'PendingLoaderPath' -Value $staged
    Set-ItemProperty -Path $regPath -Name 'PendingLoaderSha256' -Value $stagedHash
    Remove-ItemProperty -Path $regPath -Name 'SwapInProgress' -ErrorAction SilentlyContinue
    $xl.Run("'ThinQQ.core.xlam'!Core_ApplyPendingLoader") | Out-Null

    # 4) Verificacoes
    $newHash = Get-OpenFileSha $target
    Check ($newHash -eq $stagedHash) 'arquivo do loader substituido (hash igual ao pendente)'
    Check (-not (Test-Path "$target.new") -and -not (Test-Path "$target.bak")) 'sem sobras .new/.bak'
    $core = Get-Wb 'ThinQQ.core.xlam'
    Check ($null -ne $core) 'CORE continua aberto apos a troca do loader'
    if ($core) { Check ($xl.Run("'ThinQQ.core.xlam'!Core_Ping") -eq 'pong') 'Core_Ping responde depois da troca' }
    $nl = Get-Wb 'ThinQQ.xlam'
    Check ($null -ne $nl) 'novo loader aberto'
    if ($nl) { Check ((Get-DocProperty $nl 'Comments') -eq 'swap-test') 'o loader aberto e o NOVO (propriedade distintiva)' }
    Check ([string]::IsNullOrEmpty((Get-RegVal 'SwapInProgress'))) 'SwapInProgress limpo'
    Check ([string]::IsNullOrEmpty((Get-RegVal 'PendingLoaderPath'))) 'PendingLoaderPath limpo'
    Check (-not (Test-Path $staged)) 'arquivo pendente consumido'
} catch {
    Write-Host "not ok - excecao: $($_.Exception.Message)"; $script:fail++
} finally {
    Stop-ExcelWatchdog $wd
    if ($xl) {
        try { $xl.DisplayAlerts = $false; $xl.EnableEvents = $false } catch { }
        Close-ExcelApp $xl
    }
    Restore-Reg
    $env:LOCALAPPDATA = $prevLocalApp
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
if (Test-ExcelWatchdogFired $wd) { Write-Host 'not ok - TIMEOUT (dialogo modal ou travamento)'; $script:fail++ }
if ($script:fail -gt 0) { Write-Host "FALHAS loader-swap: $script:fail"; exit 1 }
Write-Host 'loader-swap: ok'
