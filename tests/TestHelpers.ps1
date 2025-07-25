Set-StrictMode -Version Latest

function Import-AddUserRoot {
    param(
        [string]$Root = (Join-Path $PSScriptRoot '..')
    )
    Import-Module (Join-Path $Root 'AddUser.psd1') -Force
}

function Use-TestConfig {
    param(
        [string]$ConfigPath = (Join-Path $PSScriptRoot 'fixtures\config.min.json')
    )
    # Direkt in den ConfigModule-Cache schreiben ist unsauber – deshalb:
    # Wir setzen Env-Var, so dass Get-AppConfig unseren Testpfad nimmt.
    $env:ADDUSER_CONFIG_PATH = $ConfigPath
}

