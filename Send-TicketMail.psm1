# Send-TicketMail.psm1

function Send-TicketMail {
    param(
        [Parameter(Mandatory=$true)][string]$TicketNr,
        [Parameter(Mandatory=$true)][string]$logText
    )

    # Filtern
    $filteredLines = $logText -split "`r?`n" | Where-Object { $_ -match "(Gruppe|Login-Daten|Extension)" } | ForEach-Object {
        $_ -replace "\[[^\]]*\]", ""
    }
    if ($filteredLines.Count -eq 0) {
        $filteredText = "Keine Gruppenaktionen oder Login-Daten protokolliert."
    }
    else {
        $filteredText = $filteredLines -join "<br>"
    }

    $ticketBody = @"
<html>
  <body>
    <h3>Gruppen &amp; Login-Daten &amp; Attribute</h3>
    <p>$filteredText</p>
  </body>
</html>
"@
    $recipients = "it-benutzerservice@bmuv.bund.de"
    $smtpServer = "strmail.office.dir"
    $subject    = "$TicketNr"

    try {
        Send-MailMessage -From "it-benutzerservice@bmuv.bund.de" `
                         -To $recipients `
                         -Subject $subject `
                         -Body $ticketBody `
                         -SmtpServer $smtpServer `
                         -BodyAsHtml `
                         -Encoding UTF8 -ErrorAction Stop

        Write-Log -Message "Ticket-E-Mail mit Betreff '$subject' erfolgreich gesendet." -Color "green"
    }
    catch {
        Write-Log -Message "Fehler beim Senden der Ticket-E-Mail: $($_.Exception.Message)" -Color "red"
    }
}

Export-ModuleMember -Function Send-TicketMail
