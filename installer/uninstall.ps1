# Desinstala o ThinQQ do usuario atual. Compativel com Windows PowerShell 5.1.
param(
    [string]$Root = (Join-Path $env:LOCALAPPDATA 'ThinQQ'),
    [string]$RegistryRoot = 'HKCU:\Software',
    [switch]$KeepFiles,
    [switch]$SkipExcelCheck
)
$ErrorActionPreference = 'Stop'

# --- Protecoes de caminho (antes de qualquer alteracao) ---
$Root = [IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Root)).TrimEnd('\')
if ([IO.Path]::GetFileName($Root) -ne 'ThinQQ' -or $Root -ieq $env:USERPROFILE.TrimEnd('\')) {
    throw "Pasta recusada: '$Root'. O nome da pasta precisa ser 'ThinQQ' (padrão: %LOCALAPPDATA%\ThinQQ)."
}
if ((Test-Path -LiteralPath $Root) -and -not ((Test-Path -LiteralPath (Join-Path $Root 'ThinQQ.xlam')) -or (Test-Path -LiteralPath (Join-Path $Root 'versions')))) {
    throw "A pasta '$Root' não parece uma instalação do ThinQQ (sem ThinQQ.xlam nem 'versions'). Nada foi removido."
}
$RegistryRoot = $RegistryRoot.TrimEnd('\')
if ($RegistryRoot -ne 'HKCU:\Software' -and $RegistryRoot -notmatch '^HKCU:\\Software\\ThinQQ_Test_[A-Za-z0-9_]+$') {
    throw "Chave de registro recusada: '$RegistryRoot'. Só HKCU:\Software é permitido."
}
if (-not $SkipExcelCheck) {
    if (Get-Process excel -ErrorAction SilentlyContinue) { throw 'Feche o Excel antes de desinstalar o ThinQQ.' }
    if (Get-Process powerpnt -ErrorAction SilentlyContinue) { throw 'Feche o PowerPoint antes de desinstalar o ThinQQ.' }
}

# --- Suplemento (OPEN/OPENn): so mexe se houver entrada do ThinQQ ---
$opt = Join-Path $RegistryRoot 'Microsoft\Office\16.0\Excel\Options'
if (Test-Path -LiteralPath $opt) {
    $key = Get-Item -LiteralPath $opt
    $entries = @($key.GetValueNames() | Where-Object { $_ -match '^OPEN\d*$' } | Sort-Object { if ($_ -eq 'OPEN') { 0 } else { [int]$_.Substring(4) } } | ForEach-Object {
        [pscustomobject]@{ Name = $_; Value = $key.GetValue($_, $null, 'DoNotExpandEnvironmentNames'); Kind = $key.GetValueKind($_).ToString() }
    })
    $isMine = { param($e) ([string]$e.Value) -match '(^|[\\/"\s])ThinQQ\.xlam"?\s*$' }
    if (@($entries | Where-Object { & $isMine $_ }).Count -gt 0) {
        $keep = @($entries | Where-Object { -not (& $isMine $_) })
        $targetNames = @()
        for ($i = 0; $i -lt $keep.Count; $i++) {
            $name = 'OPEN'
            if ($i -gt 0) { $name = "OPEN$i" }
            $targetNames += $name
            $cur = $entries | Where-Object { $_.Name -eq $name }
            if (-not $cur -or $cur.Value -ne $keep[$i].Value -or $cur.Kind -ne $keep[$i].Kind) {
                Set-ItemProperty -LiteralPath $opt -Name $name -Value $keep[$i].Value -Type $keep[$i].Kind
            }
        }
        foreach ($e in $entries) { if ($targetNames -notcontains $e.Name) { Remove-ItemProperty -LiteralPath $opt -Name $e.Name } }
    }
}
# PowerPoint (1.4.0): so as duas chaves que o instalador cria; as configuracoes do atualizador (PpamVersion) saem com 'VB and VBA Program Settings\ThinQQ'
Remove-Item -LiteralPath (Join-Path $RegistryRoot 'Microsoft\Office\16.0\PowerPoint\AddIns\ThinQQ') -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path $RegistryRoot 'Microsoft\Office\16.0\PowerPoint\Security\Trusted Locations\LocationThinQQ') -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path $RegistryRoot 'Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ') -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path $RegistryRoot 'VB and VBA Program Settings\ThinQQ') -Recurse -Force -ErrorAction SilentlyContinue
if (-not $KeepFiles -and (Test-Path -LiteralPath $Root)) { Remove-Item -LiteralPath $Root -Recurse -Force }
Write-Host 'ThinQQ desinstalado.'
