Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name Pester)) {
    Install-Module Pester -Force -Scope CurrentUser
}
Import-Module Pester -Force

Invoke-Pester -Path (Join-Path $PSScriptRoot '..\tests') -CI
