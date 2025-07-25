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
    $env:ADDUSER_LOGPATH    = Join-Path $env:TEMP 'AddUserTest.log'
    Import-Module (Join-Path $PSScriptRoot '..\ConfigModule.psm1') -Force
    $global:AppConfig = Get-AppConfig
}

