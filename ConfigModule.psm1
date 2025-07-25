<#  ConfigModule.psm1  –  zentraler, gecachter Zugriff auf AppConfig  #>
$script:AppConfig = $null
function Get-AppConfig {
    [CmdletBinding()]
    param(
        [string]$Path = (Join-Path $PSScriptRoot 'config.json')
    )
    if ($script:AppConfig) { return $script:AppConfig }
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log -Message "Config not found at '$Path'." -Category ERROR
        throw "Config not found: $Path"
    }
    try {
        $script:AppConfig = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        Write-Log -Message "Config loaded & cached from '$Path'." -Category DEBUG
        return $script:AppConfig
    } catch {
        Write-Log -Message "Config load error: $($_.Exception.Message)" -Category ERROR
        throw
    }
}
Export-ModuleMember -Function Get-AppConfig
