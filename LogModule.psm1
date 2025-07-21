# LogModule.psm1
# -----------------------------------------
# Korrigierte Version für RichTextBox-Logging
# über $global:WpfLogControl.Dispatcher
# -----------------------------------------

function MyWrite-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [string]$Color = "Black"
    )

    # Zeitstempel
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formattedMessage = "[$timestamp] $Message"

    # 1) Optional: In Datei schreiben (falls konfiguriert)
    if ($global:AppConfig -and $global:AppConfig.LogPath) {
        $formattedMessage | Out-File -FilePath $global:AppConfig.LogPath -Append -Encoding UTF8
    }

    # 2) In die RichTextBox-Logausgabe
    if ($global:WpfLogControl -and $global:WpfLogControl.Dispatcher) {
        $global:WpfLogControl.Dispatcher.Invoke([Action]{
            $run = New-Object Windows.Documents.Run
            $run.Text = "$formattedMessage`r`n"
            switch ($Color.ToLower()) {
                "red"   { $run.Foreground = [System.Windows.Media.Brushes]::Red }
                "green" { $run.Foreground = [System.Windows.Media.Brushes]::Green }
                "blue"  { $run.Foreground = [System.Windows.Media.Brushes]::Blue }
                default { $run.Foreground = [System.Windows.Media.Brushes]::Black }
            }
            $paragraph = New-Object Windows.Documents.Paragraph($run)
            $global:WpfLogControl.Document.Blocks.Add($paragraph)
            $global:WpfLogControl.ScrollToEnd()
        })
    }
}

function WriteJobLog {
    param(
        [string]$msg,
        [string]$category = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    if ($InJob) {
        [pscustomobject]@{
            Log = "[$timestamp] [$category] $msg"
            Category = $category
        } | Write-Output
        Start-Sleep -Milliseconds 10
    }
    else {
        $run = New-Object Windows.Documents.Run
        $run.Text = "[$timestamp] [$category] $msg`r`n"
        switch ($category) {
            "ERROR"   { $run.Foreground = [System.Windows.Media.Brushes]::Red }
            "SUCCESS" { $run.Foreground = [System.Windows.Media.Brushes]::Green }
            "WARN"    { $run.Foreground = [System.Windows.Media.Brushes]::Orange }
            "INFO"    { $run.Foreground = [System.Windows.Media.Brushes]::Black }
            default   { $run.Foreground = [System.Windows.Media.Brushes]::Black }
        }
        $paragraph = New-Object Windows.Documents.Paragraph($run)
        $LogTextbox.Document.Blocks.Add($paragraph)
        $LogTextbox.ScrollToEnd()
    }
}




Export-ModuleMember -Function MyWrite-Log, WriteJobLog