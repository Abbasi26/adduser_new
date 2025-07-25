# -----------------------------------------------------------------------
# LogModule.psm1              Stand: 25 Jul 2025
# – schreibt jede Zeile gleichzeitig
#   • in den Primär‑Pfad  $global:AppConfig.Paths.LogPath     (UNC)
#   • in den Fallback‑Pfad $env:TEMP\AddUser\AddUser_yyyyMMdd.log
# – bricht niemals mit „‑Path $null“ ab
# – GUI‑ und Job‑fähig
# -----------------------------------------------------------------------

## interner Zustand
$script:PrimaryLogFile   = $null
$script:TempLogFile      = Join-Path $env:TEMP ("AddUser\AddUser_{0:yyyyMMdd}.log" -f (Get-Date))
$script:WpfLogControl    = $null
$script:LoggerReady      = $false

## Hilfsfunktionen
function Resolve-PrimaryLogPath {
    if ($global:AppConfig -and $global:AppConfig.Paths.LogPath -is [string] -and $global:AppConfig.Paths.LogPath.Trim()) {
        return $global:AppConfig.Paths.LogPath.Trim()
    }
    return $null
}

function Ensure-Directory {
    param([string]$Path)
    try {
        $dir = Split-Path $Path -Parent
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
    } catch { }
}

function Write-ToLogFiles {
    param([string]$Line)
    # Primär
    if ($script:PrimaryLogFile) {
        try { $Line | Out-File -LiteralPath $script:PrimaryLogFile -Append -Encoding UTF8 } catch { }
    }
    # Temp (immer angelegt)
    $Line | Out-File -LiteralPath $script:TempLogFile -Append -Encoding UTF8
}

## Öffentliche API
function Initialize-Logger {
    [CmdletBinding()]
    param(
        [System.Windows.Controls.RichTextBox]$RichTextBox,
        [switch]$Append = $true
    )
    # 1 Primär
    $script:PrimaryLogFile = Resolve-PrimaryLogPath
    if ($script:PrimaryLogFile) {
        Ensure-Directory $script:PrimaryLogFile
        if (-not $Append) { "" | Out-File $script:PrimaryLogFile -Encoding UTF8 }
    }
    # 2 Temp
    Ensure-Directory $script:TempLogFile
    if (-not $Append) { "" | Out-File $script:TempLogFile -Encoding UTF8 }
    # 3 GUI
    $script:WpfLogControl = $RichTextBox
    $script:LoggerReady   = $true
    Write-Log "Logger initialisiert  → Primär='$script:PrimaryLogFile'  Temp='$script:TempLogFile'"
}

function Write-Log {
    [CmdletBinding()] param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')] [string]$Level = 'INFO'
    )
    if (-not $script:LoggerReady) { return }
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    Write-ToLogFiles $line
    # GUI optional
    if ($script:WpfLogControl -and $script:WpfLogControl.Dispatcher) {
        try {
            $script:WpfLogControl.Dispatcher.Invoke([Action]{
                $run = New-Object Windows.Documents.Run ($line + "`r`n")
                switch ($Level) {
                    'ERROR'   { $run.Foreground = [System.Windows.Media.Brushes]::Red }
                    'SUCCESS' { $run.Foreground = [System.Windows.Media.Brushes]::Green }
                    'WARN'    { $run.Foreground = [System.Windows.Media.Brushes]::Orange }
                    'DEBUG'   { $run.Foreground = [System.Windows.Media.Brushes]::Gray }
                    default   { $run.Foreground = [System.Windows.Media.Brushes]::Black }
                }
                $para = [Windows.Documents.Paragraph]::new($run)
                $script:WpfLogControl.Document.Blocks.Add($para)
                $script:WpfLogControl.ScrollToEnd()
            })
        } catch { }
    }
}

function Get-LogPath     { $script:PrimaryLogFile }
function Get-TempLogPath { $script:TempLogFile }

Export-ModuleMember -Function Initialize-Logger, Write-Log, Get-LogPath, Get-TempLogPath