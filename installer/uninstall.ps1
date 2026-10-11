# Desinstala o ThinQQ do usuario atual. Compativel com Windows PowerShell 5.1.
param(
    [string]$Root = (Join-Path $env:LOCALAPPDATA 'ThinQQ'),
    [string]$RegistryRoot = 'HKCU:\Software',
    [switch]$KeepFiles,
    [switch]$SkipExcelCheck
)
$ErrorActionPreference = 'Stop'
if (-not $SkipExcelCheck) {
    if (Get-Process excel -ErrorAction SilentlyContinue) { throw 'Feche o Excel antes de desinstalar o ThinQQ.' }
}
$opt = Join-Path $RegistryRoot 'Microsoft\Office\16.0\Excel\Options'
if (Test-Path $opt) {
    $props = @((Get-ItemProperty $opt).PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' } | Sort-Object { if ($_.Name -eq 'OPEN') { 0 } else { [int]$_.Name.Substring(4) } })
    $values = @($props | Where-Object { [string]$_.Value -notlike '*ThinQQ.xlam*' } | ForEach-Object { $_.Value })
    foreach ($p in $props) { Remove-ItemProperty -Path $opt -Name $p.Name }
    for ($i = 0; $i -lt $values.Count; $i++) {
        $name = 'OPEN'
        if ($i -gt 0) { $name = "OPEN$i" }
        Set-ItemProperty -Path $opt -Name $name -Value $values[$i]
    }
}
Remove-Item (Join-Path $RegistryRoot 'Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ') -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $RegistryRoot 'VB and VBA Program Settings\ThinQQ') -Recurse -Force -ErrorAction SilentlyContinue
if (-not $KeepFiles -and (Test-Path $Root)) { Remove-Item $Root -Recurse -Force }
Write-Host 'ThinQQ desinstalado.'
