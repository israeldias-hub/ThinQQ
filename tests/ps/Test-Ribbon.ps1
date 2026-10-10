# Verifica que o pacote ThinQQ.test.xlam contem o customUI14.xml identico ao fonte e o relacionamento no .rels.
param([string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
Add-Type -AssemblyName System.IO.Compression.FileSystem
$pkg = Join-Path $RepoRoot 'dist\ThinQQ.test.xlam'
$src = Join-Path $RepoRoot 'src\loader\ribbon\customUI14.xml'
$script:fail = 0
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m"; $script:fail++ } }
$fs = New-Object System.IO.FileStream($pkg, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
try {
    $e = $zip.GetEntry('customUI/customUI14.xml')
    Check ($null -ne $e) 'pacote contem customUI/customUI14.xml'
    if ($e) {
        $ms = New-Object System.IO.MemoryStream; $s = $e.Open(); $s.CopyTo($ms); $s.Dispose()
        $a = [System.Convert]::ToBase64String($ms.ToArray())
        $b = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($src))
        Check ($a -eq $b) 'customUI14.xml do pacote identico ao fonte'
    }
    $r = $zip.GetEntry('_rels/.rels')
    Check ($null -ne $r) 'pacote contem _rels/.rels'
    if ($r) {
        $sr = New-Object System.IO.StreamReader($r.Open()); $txt = $sr.ReadToEnd(); $sr.Dispose()
        Check ($txt -match 'http://schemas\.microsoft\.com/office/2007/relationships/ui/extensibility') 'relationship ui/extensibility presente'
        Check ($txt -match 'customUI/customUI14\.xml') 'relationship aponta para customUI/customUI14.xml'
    }
} finally { $zip.Dispose(); $fs.Dispose() }
if ($script:fail -gt 0) { Write-Host "FALHAS ribbon-pacote: $script:fail"; exit 1 }
