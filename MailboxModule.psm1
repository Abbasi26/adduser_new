# MailboxModule.psm1
function Create-Mailbox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$site,
        
        [Parameter(Mandatory=$true)]
        [string[]]$accounts,

        [switch]$extern,
        [switch]$hiddenFromAddressLists,
        [switch]$externAccount,
        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.RichTextBox]$LogTextbox
    )

    if ($LogTextbox) {
        MyWrite-Log "Starte Create-Mailbox: Site=$site, extern=$extern, hidden=$hiddenFromAddressLists, externAccount=$externAccount" -Color "blue" -LogTextbox $LogTextbox
    }

    try {
        $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange `
                                         -ConnectionUri "http://rspsvexch12.office.dir/PowerShell/" `
                                         -Authentication Kerberos
        Import-PSSession $exchangeSession -ErrorAction Stop
    }
    catch {
        if ($LogTextbox) {
            MyWrite-Log "FEHLER: Fehler beim Herstellen der Exchange-Session: $($_.Exception.Message)" -Color "red" -LogTextbox $LogTextbox
        }
        return
    }

    foreach ($samAccountName in $accounts) {
        try {
            $adUser = Get-ADUser -Identity $samAccountName -Properties Surname, GivenName -ErrorAction Stop
            $dn = $adUser.DistinguishedName

            if ($externAccount) {
                Set-ADUser -Identity $samAccountName -Add @{ extensionAttribute11="x" } -ErrorAction Stop
                if ($LogTextbox) {
                    MyWrite-Log "extensionAttribute11 wurde für $samAccountName auf x gesetzt (Extern Account)." -Color "green" -LogTextbox $LogTextbox
                }
            }

            if ($extern) {
                $newDisplayName = "$($adUser.Surname), $($adUser.GivenName) (EXTERN)"
                Set-ADUser -Identity $samAccountName -DisplayName $newDisplayName -ErrorAction Stop
                Rename-ADObject -Identity $adUser.DistinguishedName -NewName $newDisplayName -ErrorAction Stop
            }

            if ($LogTextbox) {
                MyWrite-Log "Suche Mailbox-Datenbank für $($samAccountName) (Site=$site)" -Color "blue" -LogTextbox $LogTextbox
            }

            $mailboxSite = "$site" + "DB*"
            $mailboxDatabases = Get-MailboxDatabase -Status | Where-Object { $_.Name -like $mailboxSite }

            if (-not $mailboxDatabases) {
                if ($LogTextbox) {
                    MyWrite-Log "KEINE passende Mailbox-Datenbank gefunden für $($samAccountName). Muster=$mailboxSite" -Color "red" -LogTextbox $LogTextbox
                }
                continue
            }

            $dbFreeSpaces = @()
            foreach ($db in $mailboxDatabases) {
                $spaceString = $db.AvailableNewMailboxSpace -replace '.*\(' -replace '\).*' -replace ',' -replace 'MB.*',''
                [double]$val = 0
                [void][double]::TryParse($spaceString, [ref]$val)
                $db.AvailableNewMailboxSpace = $val
                $dbFreeSpaces += $db
            }

            $userMailBoxDB = $dbFreeSpaces | Sort-Object AvailableNewMailboxSpace -Descending | Select-Object -First 1
            if (-not $userMailBoxDB) {
                if ($LogTextbox) {
                    MyWrite-Log "Keine geeignete Mailbox-Datenbank gefunden für $($samAccountName)." -Color "red" -LogTextbox $LogTextbox
                }
                continue
            }

            $archiveDbName = ($userMailBoxDB.Name -replace '^([A-Za-z]+)DB','${1}ADB')

            if ($LogTextbox) {
                MyWrite-Log "Primäre DB: $($userMailBoxDB.Name), Archiv-DB: $archiveDbName" -Color "green" -LogTextbox $LogTextbox
            }

            Enable-Mailbox -Identity $dn `
                           -Alias $samAccountName `
                           -Database $userMailBoxDB.Name `
                           -RetentionPolicy 'BMU Benutzerpostfach Archivierungsrichtlinie' `
                           -ErrorAction Stop

            Enable-Mailbox -Identity $dn `
                           -ArchiveDatabase $archiveDbName `
                           -Archive `
                           -ErrorAction Stop

            if ($extern) {
                $smtpMail = ($adUser.GivenName.Split(" ")[0] + "." + $adUser.Surname + ".extern@bmuv.bund.de").ToLower()
                Set-Mailbox -Identity $dn -EmailAddressPolicyEnabled $false -PrimarySmtpAddress $smtpMail -ErrorAction Stop
            }

            if ($hiddenFromAddressLists) {
                Set-Mailbox -Identity $dn -HiddenFromAddressListsEnabled $true -ErrorAction Stop
            }

            if ($LogTextbox) {
                MyWrite-Log "Mailbox erstellt: $($samAccountName), DB=$($userMailBoxDB.Name), Archiv=$archiveDbName" -Color "green" -LogTextbox $LogTextbox
            }
        }
        catch {
            if ($LogTextbox) {
                MyWrite-Log "FEHLER bei $($samAccountName): $($_.Exception.Message)" -Color "red" -LogTextbox $LogTextbox
            }
        }
    }
    Remove-PSSession $exchangeSession -ErrorAction SilentlyContinue
}

Export-ModuleMember -Function Create-Mailbox
