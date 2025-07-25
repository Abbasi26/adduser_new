<#  -----------------------------------------------------------------------
    LogModule.psm1  –  BMU AddUser  (Stand 25 Jul 2025)
    -----------------------------------------------------------------------
    - Schreibt jede Logzeile gleichzeitig
        1. in die zentrale Datei   $global:AppConfig.Paths.LogPath
        2. in die lokale Datei     "$env:TEMP\AddUser_{yyyyMMdd}.log"
    - Funktioniert in GUI-, Konsolen- und Job-Kontexten.
    - Erstellt Zielordner automatisch; fällt niemals mit "-Path = $null".
    - Öffentliche API:
        Initialize-Logger   Stop-Logger
        Write-Log (Alias: MyWrite-Log)
        WriteJobLog
        Get-LogPath    Get-TempLogPath
        Get-FullLog    Clear-FullLog
   -----------------------------------------------------------------------#>

#region -- private state
$script:CentralLogFile = $null      # UNC / Fileshare
$script:TempLogFile    = $null      # immer vorhanden
$script:WpfLogControl  = $null
$script:InJobContext   = $false
#endregion

function Initialize-Logger {
    [CmdletBinding()] param(
        [object] $WpfControl,
        [switch] $InJob,
        [switch] $Append   = $true
    )

    # 1) Pfade ermitteln
    $script:CentralLogFile = $null
    if ($global:AppConfig -and $global:AppConfig.Paths.LogPath) {
        $script:CentralLogFile = $global:AppConfig.Paths.LogPath.Trim()
    }
    $script:TempLogFile = Join-Path $env:TEMP ("AddUser_{0:yyyyMMdd}.log" -f (Get-Date))

    # 2) Ordner anlegen / Testschreiben
    foreach ($file in @($script:CentralLogFile,$script:TempLogFile) | Where-Object { $_ }) {
        try {
            $dir = Split-Path $file -Parent
            if (-not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
            if (-not $Append) { "" | Out-File $file -Encoding UTF8 -Force }
            elseif (-not (Test-Path $file)) { "" | Out-File $file -Encoding UTF8 }
        } catch {
            # Wenn die ZENTRALE Datei nicht erreichbar ist -> nur Temp.
            if ($file -eq $script:CentralLogFile) { $script:CentralLogFile = $null }
        }
    }

    $script:WpfLogControl = $WpfControl
    $script:InJobContext  = $InJob.IsPresent

    Write-Log "Logger initialisiert. Central='$script:CentralLogFile'; Temp='$script:TempLogFile'" INFO
}

function Stop-Logger {
    Write-Log "Logger gestoppt." INFO
    $script:CentralLogFile = $null
    $script:TempLogFile    = $null
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')]
        [string]$Category = 'INFO'
    )

    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line  = "[$stamp] [$Category] $Message"

    # 1/2  Datei(en)
    foreach ($f in @($script:CentralLogFile,$script:TempLogFile) | Where-Object { $_ }) {
        try { $line | Out-File $f -Encoding UTF8 -Append } catch {}
    }

    # 2/2  GUI oder Host
    if ($script:WpfLogControl -and $script:WpfLogControl.Dispatcher) {
        try {
            $script:WpfLogControl.Dispatcher.Invoke([Action]{
                $run        = New-Object Windows.Documents.Run($line + "`r`n")
                $run.Foreground = switch ($Category) {
                    'ERROR'   { [System.Windows.Media.Brushes]::Red }
                    'SUCCESS' { [System.Windows.Media.Brushes]::Green }
                    'WARN'    { [System.Windows.Media.Brushes]::Orange }
                    'DEBUG'   { [System.Windows.Media.Brushes]::Gray }
                    default   { [System.Windows.Media.Brushes]::Black }
                }
                $para = New-Object Windows.Documents.Paragraph($run)
                $script:WpfLogControl.Document.Blocks.Add($para); $script:WpfLogControl.ScrollToEnd()
            })
        } catch {}
    } elseif (-not $script:InJobContext) {
        Write-Host $line
    }

    # Hintergrund-Job -> Objekt zurückgeben
    if ($script:InJobContext) {
        [pscustomobject]@{ Timestamp=$stamp; Category=$Category; Message=$Message; LogLine=$line }
    }
}
Set-Alias MyWrite-Log Write-Log -Scope Global

function WriteJobLog { param([string]$msg,[string]$Category='INFO') Write-Log $msg $Category }

function Get-LogPath     { $script:CentralLogFile }
function Get-TempLogPath { $script:TempLogFile   }

function Get-FullLog {
    foreach ($f in @($script:CentralLogFile,$script:TempLogFile) | Where-Object { $_ -and (Test-Path $_) }) {
        "=== $f ==="; Get-Content $f -Encoding UTF8
    }
}

function Clear-FullLog {
    foreach ($f in @($script:CentralLogFile,$script:TempLogFile) | Where-Object { $_ -and (Test-Path $_) }) {
        Clear-Content $f -ErrorAction SilentlyContinue
    }
}

Export-ModuleMember -Function Initialize-Logger,Stop-Logger,Write-Log,MyWrite-Log,WriteJobLog,Get-LogPath,Get-TempLogPath,Get-FullLog,Clear-FullLog
