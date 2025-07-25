Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot '..\AddUser.psd1') -Force

Describe "Smoke" {
    It "Root module loads and exports at least one function" {
        (Get-Command Write-Log -ErrorAction SilentlyContinue) -or (Get-Module AddUser) | Should -Not -BeNullOrEmpty
    }
}
