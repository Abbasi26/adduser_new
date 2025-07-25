<#  ConfigModule.psm1  –  zentraler, gecachter Zugriff auf AppConfig (mit flexiblen Pfaden) #>
$script:AppConfig = $null

function Resolve-ConfigPath {
    [CmdletBinding()]
    param(
        [string]$Path
    )
    if ($Path) { return $Path }

    # 1) Env Overrides
    if ($env:ADDUSER_CONFIG_PATH) { return $env:ADDUSER_CONFIG_PATH }
    if ($env:ADDUSER_CONFIG_DIR)  { return (Join-Path $env:ADDUSER_CONFIG_DIR 'config.json') }

    # 2) Default relativ zur kompilierten/verteilten Struktur (dist\...)
    #    $PSScriptRoot = ...\Modules\Public
    $modulesRoot = Split-Path -Parent $PSScriptRoot
    $distRoot    = Split-Path -Parent $modulesRoot
    $candidate   = Join-Path $distRoot 'config\config.json'
    if (Test-Path -LiteralPath $candidate) { return $candidate }

    # 3) Fallback: Repo-Struktur (…\config\config.json)
    $repoRoot  = Split-Path -Parent $distRoot
    $candidate2 = Join-Path $repoRoot 'config\config.json'
    if (Test-Path -LiteralPath $candidate2) { return $candidate2 }

    throw "Keinen gültigen Config-Pfad gefunden. Nutze -Path, ADDUSER_CONFIG_PATH oder ADDUSER_CONFIG_DIR."
}

function Get-AppConfig {
    [CmdletBinding()]
    param(
        [string]$Path
    )
    if ($script:AppConfig) { return $script:AppConfig }

    $resolved = Resolve-ConfigPath -Path $Path
    if (-not (Test-Path -LiteralPath $resolved)) {
        Write-Log -Message "Config not found at '$resolved'." -Category ERROR
        throw "Config not found: $resolved"
    }
    try {
        $script:AppConfig = Get-Content -LiteralPath $resolved -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($env:ADDUSER_LOGPATH) {
            $script:AppConfig.Paths.LogPath = $env:ADDUSER_LOGPATH
        }
        Write-Log -Message "Config loaded & cached from '$resolved'." -Category DEBUG
        return $script:AppConfig
    } catch {
        Write-Log -Message "Config load error: $($_.Exception.Message)" -Category ERROR
        throw
    }
}
Export-ModuleMember -Function Get-AppConfig,Resolve-ConfigPath
