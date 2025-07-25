param(
    [switch]$Sign,
    [string]$CertThumbprint,
    [switch]$RequireAdmin
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root   = Split-Path -Parent $PSCommandPath
$repo   = Split-Path -Parent $root
$dist   = Join-Path $repo 'dist'
$input  = Join-Path $repo 'MainGUI5.ps1'
$outExe = Join-Path $dist 'AddUserGUI.exe'

if (!(Test-Path $dist)) { New-Item -ItemType Directory -Path $dist | Out-Null }

$psd1 = Join-Path $repo 'AddUser.psd1'
$manifest = Import-PowerShellDataFile $psd1
$version  = $manifest.ModuleVersion

try {
    Import-Module ps2exe -ErrorAction Stop
} catch {
    Write-Host "ps2exe nicht gefunden – installiere..." -ForegroundColor Yellow
    Install-Module ps2exe -Scope CurrentUser -Force -AllowClobber
    Import-Module ps2exe -Force
}

$args = @{
    InputFile   = $input
    OutputFile  = $outExe
    Title       = 'BMU AddUser'
    Product     = 'BMU AddUser'
    Company     = 'BMU'
    Description = 'BMU AddUser GUI (portable)'
    Version     = $version
}
if ($RequireAdmin) { $args.RequireAdmin = $true }

Write-Host "==> Baue $outExe (v$version) ..." -ForegroundColor Cyan
ps2exe @args

$payload = @(
    @{ Src = Join-Path $repo 'config';  Dst = Join-Path $dist 'config'  },
    @{ Src = Join-Path $repo 'Xaml';    Dst = Join-Path $dist 'Xaml'    },
    @{ Src = Join-Path $repo 'Modules'; Dst = Join-Path $dist 'Modules' },
    @{ Src = Join-Path $repo 'AddUser.psd1'; Dst = Join-Path $dist 'AddUser.psd1' },
    @{ Src = Join-Path $repo 'AddUser.psm1'; Dst = Join-Path $dist 'AddUser.psm1' }
)

foreach ($item in $payload) {
    if (Test-Path $item.Src) {
        if (Test-Path $item.Dst) { Remove-Item $item.Dst -Recurse -Force }
        Copy-Item $item.Src $item.Dst -Recurse -Force
    }
}

if ($Sign) {
    if (-not $CertThumbprint) { throw "-Sign ohne -CertThumbprint." }
    $cert = Get-ChildItem Cert:\CurrentUser\My\$CertThumbprint, Cert:\LocalMachine\My\$CertThumbprint -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $cert) { throw "Zertifikat $CertThumbprint nicht gefunden." }
    Set-AuthenticodeSignature -FilePath $outExe -Certificate $cert -TimestampServer 'http://timestamp.sectigo.com'
}

Write-Host "Fertig. EXE: $outExe" -ForegroundColor Green
