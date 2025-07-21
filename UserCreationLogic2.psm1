

# ------------------------------------------------
# Funktion: Convert-DepartmentShort
# ------------------------------------------------
function Convert-DepartmentShort {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Department
    )
    # Entfernt zunächst alle Leerzeichen (z.B. "Z II 5" -> "ZII5")
    $converted = $Department -replace '\s+', ''

    # Ersetze römische Zahlen durch arabische.
    $converted = $converted -replace 'VI','6'
    $converted = $converted -replace 'IV','4'
    $converted = $converted -replace 'III','3'
    $converted = $converted -replace 'II','2'
    $converted = $converted -replace 'I','1'
    $converted = $converted -replace 'V','5'

return $converted
}

# ------------------------------------------------
# Funktion: ProcessUserCreation
# ------------------------------------------------
function ProcessUserCreation {
    param (
        [string]$UserID,
        [string]$givenName,
        [string]$lastName,
        [string]$gender,
        [string]$Buro,
        [string]$Rufnummer,
        [string]$Handynummer,
        [string]$titleValue,
        [string]$amtsbez,        # Amtsbezeichnung
        [string]$laufgruppe,     # Laufbahngruppe
        [string]$roleSelection,  # Praktikant/Hospitant/Azubi
        [string]$Site,
        [string]$ExpDate,
        [string]$desc,
        [string]$Department,
        [string]$TicketNr,
        [string]$EntryDate,
        [string]$sonderkenn,
        [string]$funktion,
        [string]$refUser,
        [string]$isIVBB,
        [string]$isGVPL,
        [string]$isVIP,
        [string]$isFemale,
        [bool]$isExtern,
        [bool]$isVerstecken,
        [bool]$isPhonebook,
        [string]$isNatPerson,
        [string]$isResMailbox,
        [string]$isAbgeordnet,
        [string]$isConet,
        [string]$isExternAccount,
        [string]$makeMailbox,
        [hashtable]$departmentOGMapping,
        [hashtable]$departmentMapping,
        [String[]]$AdditionalGroups,
        [System.Windows.Controls.RichTextBox]$LogTextbox,  # Angepasst auf WPF RichTextBox
        [switch]$InJob,                                     # Hintergundmodus (z. B. per Scheduled Job)
        [scriptblock]$ProgressCallback                      # Fortschrittsrückruf-Funktion
    )
    # Prüfe, ob ein Aktivierungsdatum gesetzt wurde (und nicht "S" für sofort)
    #$useDelayedActivation = $false
    #if ($EntryDate -and $EntryDate -ne "S") {
    #    $useDelayedActivation = $true
    #    WriteJobLog "Aktivierungsdatum gesetzt ($EntryDate). Accountoptionen werden vorerst zurückgesetzt (nicht im Telefonbuch, nicht im GVPL, Mailbox versteckt)." "INFO"
    #}
    
    # ------------------------------------------------
    # Hilfsfunktion: WriteJobLog
    # - Schreibt Log-Einträge entweder in das RichTextBox-Element (GUI)
    #   oder gibt sie als PSCustomObject aus (für den Job-Mode).
    # ------------------------------------------------
    function WriteJobLog {
        param(
            [string]$msg,
            [string]$category = "INFO"
        )

        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        if ($InJob) {
            # Bei Job-Mode: Ausgabe in Pipeline, damit Logs gesammelt werden können.
            [pscustomobject]@{ 
                Log      = "[$timestamp] [$category] $msg"
                Category = $category 
            } | Write-Output
            
            # Kurz sleepen, um sicherzustellen, dass Output nicht gepuffert bleibt
            Start-Sleep -Milliseconds 10
        }
        else {
            # Bei GUI-Modus: Ausgabe in RichTextBox (Farbe je nach Kategorie)
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

    # ------------------------------------------------
    # Start: Benutzererstellung
    # ------------------------------------------------
    WriteJobLog "Starte ProcessUserCreation für $UserID"
    
    # Schritt 0: Fortschritt initialisieren
    if ($ProgressCallback) { & $ProgressCallback 0 }

    # ------------------------------------------------
    # Schritt 1: AD-User prüfen/erstellen
    # ------------------------------------------------
    WriteJobLog "Prüfe Existenz von $UserID in AD..."
    if ($ProgressCallback) { & $ProgressCallback 5 }

    # Suche den Benutzer in AD
    $adUser = Get-ADUser -Filter { SamAccountName -eq $UserID } -ErrorAction SilentlyContinue 
    $created = $false
    $response = $null

    if (-not $adUser) {
        # Benutzer existiert nicht -> neu anlegen
        WriteJobLog "Benutzer $UserID existiert nicht, lege an."
        
        # new-ADAccountSettings erstellt ein Objekt mit allen notwendigen Daten
        $newData = new-ADAccountSettings -newUserID $UserID -givenName $givenName -lastName $lastName -expirationDate $ExpDate -userDescription $desc
        if ($newData) {
            # Set-NewADAccount legt den AD-Benutzer an
            $setData = Set-NewADAccount -newUserData $newData 
            if ($setData) { 
                $created = $true
                WriteJobLog "Benutzer $UserID erfolgreich erstellt." "SUCCESS"
            }
        }
        else {
            WriteJobLog "Fehler: new-ADAccountSettings hat keine Daten zurückgegeben." "ERROR"
            throw "Benutzererstellung fehlgeschlagen: Keine Daten von new-ADAccountSettings"
        }
    }
    else {
        # Benutzer existiert bereits
        WriteJobLog "Benutzer $UserID existiert bereits." "WARN"
        
        # Wenn wir im Job-Mode laufen, abbrechen (je nach Logik)
        if ($InJob) {
            WriteJobLog "Abbruch, weil $UserID schon existiert (Job-Mode)." "WARN"
            return
        }
        else {
            # Im GUI-Kontext: Nachfragen, ob wir fortfahren oder abbrechen wollen
            $response = [System.Windows.MessageBox]::Show("User existiert bereits. Fortfahren?", "Benutzer existiert bereits", [System.Windows.MessageBoxButton]::YesNo)
            if ($response -eq [System.Windows.MessageBoxResult]::Yes) {
                WriteJobLog "Setze DFS- und Fileserver-Berechtigungen zurück für $UserID."
                set-FilePermissions -UserID $UserID -Site $Site 
                set-FolderOwnership -UserID $UserID -Site $Site 
                WriteJobLog "Fertig. Benutzer $UserID aktualisiert." "SUCCESS"
                return
            }
            elseif ($response -eq [System.Windows.MessageBoxResult]::No) {
                WriteJobLog "Abbruch durch User. $UserID wird nicht geändert." "WARN"
                return
            }
        }
    }

    # Schritt 1: abgeschlossen
    if ($ProgressCallback) { & $ProgressCallback 10 }

    # ------------------------------------------------
    # Schritt 2: Standardattribute setzen
    # ------------------------------------------------
    WriteJobLog "Setze Büro, Rufnummer usw. für $UserID."
    if ($ProgressCallback) { & $ProgressCallback 15 }

    # Wenn Büro-Eingabe leer, setze Standard
    if ($Buro -eq "") {
        $officeValue = "$Site/-"
    }
    else {
        $officeValue = "$Site/$Buro"
    }

    # Hash für zu ändernde AD-Attribute vorbereiten
    $replaceHash = @{
        physicalDeliveryOfficeName = $officeValue
    }
    if ($Rufnummer   -ne "") { $replaceHash["telephoneNumber"]      = $Rufnummer }
    if ($Handynummer -ne "") { $replaceHash["mobile"]               = $Handynummer }
    if ($Department  -ne "") { $replaceHash["department"]           = $Department }
    if ($titleValue  -ne "") { $replaceHash["title"]                = $titleValue }
    if ($amtsbez     -ne "") { $replaceHash["extensionAttribute2"]  = $amtsbez }
    if ($laufgruppe  -ne "") { $replaceHash["extensionAttribute8"]  = $laufgruppe }
    if ($sonderkenn  -ne "") { $replaceHash["extensionAttribute10"] = $sonderkenn }
    if ($funktion    -ne "") { $replaceHash["extensionAttribute9"]  = $funktion }

    # Versuche die AD-Attribute zu setzen
    try {
        Set-ADUser -Identity $UserID -Replace $replaceHash -ErrorAction Stop
        WriteJobLog "Standardattribute für $UserID gesetzt."
    }
    catch {
        WriteJobLog "Fehler beim Setzen der Standardattribute: $($_.Exception.Message)" "ERROR"
        throw
    }

    # Standardgruppen hinzufügen
    WriteJobLog "Füge Standardgruppen hinzu."
    $standardGroups = @(
        "ProxyUser",
        "Users-Win7",
        "VBMUBAlleMitarbeiter",
        "Verteiler − BMU"
    )

    # Nur intern hinzufügen, wenn User kein Extern-Account ist
    if (-not $isExternAccount -and -not $isExtern) {
        $standardGroups += "VBMUIntern"
    }

    foreach ($group in $standardGroups) {
        try {
            Add-ADGroupMember -Identity $group -Members $UserID -ErrorAction Stop
            WriteJobLog "Gruppe $group zugewiesen an $UserID."
        }
        catch {
            WriteJobLog "Fehler beim Hinzufügen zu ${group}: $($_.Exception.Message)" "WARN"
        }
    }

    # Schritt 2: abgeschlossen
    if ($ProgressCallback) { & $ProgressCallback 20 }

    # ------------------------------------------------
    # Schritt 3: Standortabhängige Gruppen
    # ------------------------------------------------
    WriteJobLog "Füge standortabhängige Gruppen zu ($Site)."
    if ($ProgressCallback) { & $ProgressCallback 25 }

    # VMWareView-Pool je nach Standort
    if ($Site -eq "RSP") {
        Add-ADGroupMember -Identity "RG-VMWareView-Pool-RSP-STD" -Members $UserID -ErrorAction SilentlyContinue
    }
    elseif ($Site -in @("STR", "KTR", "KRA", "ZIM")) {
        Add-ADGroupMember -Identity "RG-VMWareView-Pool-STR-STD" -Members $UserID -ErrorAction SilentlyContinue
    }

    # Verteiler pro Liegenschaft
    if ($Site -eq "RSP") {
        Add-ADGroupMember -Identity "Verteiler - Liegenschaft RSP" -Members $UserID -ErrorAction SilentlyContinue
    }
    elseif ($Site -eq "STR") {
        Add-ADGroupMember -Identity "Verteiler - Liegenschaft STR" -Members $UserID -ErrorAction SilentlyContinue
        Add-ADGroupMember -Identity "Verteiler - Liegenschaft Berlin" -Members $UserID -ErrorAction SilentlyContinue
    }
    elseif ($Site -eq "KTR") {
        Add-ADGroupMember -Identity "Verteiler - Liegenschaft STR" -Members $UserID -ErrorAction SilentlyContinue
    }

    # Spezielle KTR-Gruppen (Gebäude)
    if ($officeValue -like "*KTR/02*" -or $officeValue -like "*KTR/03*") {
        Add-ADGroupMember -Identity "Verteiler - Liegenschaft KTR 2-3" -Members $UserID -ErrorAction SilentlyContinue
    }
    elseif ($officeValue -like "*KTR/04*") {
        Add-ADGroupMember -Identity "Verteiler - Liegenschaft KTR 4-1-819963447" -Members $UserID -ErrorAction SilentlyContinue
    }

    WriteJobLog "Standardgruppen zugewiesen an $UserID."

    if ($ProgressCallback) { & $ProgressCallback 30 }

    # ------------------------------------------------
    # Schritt 4: Gruppenmitgliedschaften via Referenz-User kopieren
    # ------------------------------------------------
    if ($refUser -and $refUser.Trim() -ne "") {
        WriteJobLog "Kopiere Gruppenmitgliedschaften von Referenz-Benutzer $refUser."
        if ($ProgressCallback) { & $ProgressCallback 35 }

        try {
            $refADUser = Get-ADUser -Identity $refUser -Properties MemberOf -ErrorAction Stop
            $refGroups = $refADUser.MemberOf
            if ($refGroups) {
                foreach ($groupDN in $refGroups) {
                    try {
                        Add-ADGroupMember -Identity $groupDN -Members $UserID -ErrorAction Stop
                        WriteJobLog "Referenz $refUser -> Gruppe $groupDN hinzugefügt für $UserID."
                    }
                    catch {
                        WriteJobLog "Fehler bei Referenz $refUser -> ${groupDN}: $($_.Exception.Message)" "WARN"
                    }
                }
            }
            else {
                WriteJobLog "Keine Gruppenmitgliedschaften bei Referenz-User $refUser gefunden." "WARN"
            }
        }
        catch {
            WriteJobLog "Fehler beim Abrufen Referenz-User ${refUser}: $($_.Exception.Message)" "ERROR"
        }
    }

    # ------------------------------------------------
    # Schritt 5: Frauen-Verteiler
    # ------------------------------------------------
    if ($isFemale -eq "j") {
        if ($Site -eq "RSP") {
            Add-ADGroupMember -Identity "Verteiler - Frauen Bonn" -Members $UserID -ErrorAction SilentlyContinue
            WriteJobLog "Benutzer $UserID -> Verteiler - Frauen Bonn"
        }
        elseif ($Site -in @("STR", "KTR")) {
            Add-ADGroupMember -Identity "Verteiler - Frauen Berlin" -Members $UserID -ErrorAction SilentlyContinue
            WriteJobLog "Benutzer $UserID -> Verteiler - Frauen Berlin"
        }
    }
    if ($ProgressCallback) { & $ProgressCallback 40 }

    # ------------------------------------------------
    # Schritt 6: Abteilungs-Verteiler hinzufügen (falls Department-Pattern passt)
    # ------------------------------------------------
    if ($Department -and $Department.Trim() -match '^(ZG|HL|[A-Z])') {
        $abteilungLetter = $Matches[1]
        $abteilungGroup  = "Verteiler - Abteilung $abteilungLetter"
        
        try {
            Add-ADGroupMember -Identity $abteilungGroup -Members $UserID -ErrorAction Stop
            WriteJobLog "Benutzer $UserID wurde in '$abteilungGroup' aufgenommen."
        }
        catch {
            WriteJobLog "Fehler beim Hinzufügen zu Abteilungsgruppe '$abteilungGroup': $($_.Exception.Message)" "WARN"
        }
    }
    if ($ProgressCallback) { & $ProgressCallback 45 }

    # ------------------------------------------------
    # Schritt 7: RG-Gruppen für Praktikant/Hospitant/Azubi
    # ------------------------------------------------
    WriteJobLog "Rolle=$roleSelection, Abgeordnet=$isAbgeordnet"
    if ($ProgressCallback) { & $ProgressCallback 50 }

    $convertedDept = Convert-DepartmentShort $Department

    # Praktikant/Hospitant/Azubi haben in der Regel ein eigenes Freigabe-Verzeichnis (RG-FIL...)
    if ($roleSelection -in @("Praktikant", "Hospitant", "Azubi")) {

        # Wähle je nach Standort den Dateiserver und den Pfad
        if ($Site -eq "RSP") {
            $RGPrefix = "RG-FIL-RSPSVFIL02_"
            $resPath  = "\\RSPSVFIL02\E$\Freigaben\OrgDataRSP\$convertedDept\$UserID"
        }
        else {
            $RGPrefix = "RG-FIL-STRSVFIL02_"
            $resPath  = "\\STRSVFIL02\E$\Freigaben\OrgDataSTR\$convertedDept\$UserID"
        }
        $gruppeName = $RGPrefix + $convertedDept + "_" + $UserID + "_RWXD"

        # Versuche, RG-Gruppe anzulegen und zuweisen
        try {
            New-ADGroup -Name $gruppeName -GroupScope DomainLocal -GroupCategory Security -Path "OU=ResGroups,OU=File-SVC,OU=Service,DC=office,DC=dir" -ErrorAction Stop
            Set-ADGroup -Identity $gruppeName -Description "Ressource: $resPath" -ErrorAction Stop
            Add-ADGroupMember -Identity $gruppeName -Members $UserID -ErrorAction Stop
            WriteJobLog "Erstellt RG-Gruppe '$gruppeName' -> $resPath" "SUCCESS"
        }
        catch {
            WriteJobLog "Fehler beim Erstellen RG-Gruppe '$gruppeName': $($_.Exception.Message)" "ERROR"
        }

        # Je nach Rolle in Praktikanten- oder Azubi-Gruppe
        if ($roleSelection -in @("Praktikant", "Hospitant")) {
            try {
                Add-ADGroupMember -Identity "Praktikanten" -Members $UserID -ErrorAction Stop
                WriteJobLog "$UserID -> Gruppe Praktikanten"
            }
            catch {
                WriteJobLog "Fehler beim Hinzufügen in 'Praktikanten': $($_.Exception.Message)" "WARN"
            }
        }
        elseif ($roleSelection -eq "Azubi") {
            $azubiGroup = if ($Site -eq "RSP") { "Azubis-Bonn" } else { "Azubis-Berlin" }
            try {
                Add-ADGroupMember -Identity $azubiGroup -Members $UserID -ErrorAction Stop
                WriteJobLog "$UserID -> Gruppe $azubiGroup"
            }
            catch {
                WriteJobLog "Fehler bei Azubis-Gruppe ${azubiGroup}: $($_.Exception.Message)" "WARN"
            }
        }
    }
    if ($ProgressCallback) { & $ProgressCallback 55 }

    # ------------------------------------------------
    # Schritt 8: Abgeordneten-Status
    # ------------------------------------------------
    if ($isAbgeordnet -eq "j" -and $Department) {
        WriteJobLog "Abgeordneten-Status aktiv für $UserID -> Department $Department"
        if ($ProgressCallback) { & $ProgressCallback 60 }

        $convertedDept = Convert-DepartmentShort $Department
        if ($Site -eq "RSP") {
            $RGPrefix = "RG-FIL-RSPSVFIL02_"
            $resPath  = "\\RSPSVFIL02\E$\Freigaben\OrgDataRSP\$convertedDept\$UserID"
        }
        else {
            $RGPrefix = "RG-FIL-STRSVFIL02_"
            $resPath  = "\\STRSVFIL02\E$\Freigaben\OrgDataSTR\$convertedDept\$UserID"
        }
        $abgGruppe = $RGPrefix + $convertedDept + "_" + $UserID + "_RWXD"

        try {
            Add-ADGroupMember -Identity $abgGruppe -Members $UserID -ErrorAction Stop
            WriteJobLog "$UserID -> Abgeordnet Gruppe '$abgGruppe'"
        }
        catch {
            WriteJobLog "Fehler bei Abgeordnet-Gruppe ${abgGruppe}: $($_.Exception.Message)" "WARN"
        }
    }
    else {
        WriteJobLog "Prüfe OG-Gruppenzuweisung, da kein Abgeordneten-Status."
    }
    if ($ProgressCallback) { & $ProgressCallback 65 }

    # ------------------------------------------------
    # Schritt 9: OG-Gruppenzuweisung (statisches Mapping)
    # ------------------------------------------------
    if ($Department -and $isAbgeordnet -ne 'j' -and ($roleSelection -eq '')) {

        $ogName = $null
        if ($departmentOGMapping.ContainsKey($Department)) {
            $ogName = $departmentOGMapping[$Department]
            WriteJobLog "Mapping-Treffer: '$Department' → OG-Gruppe '$ogName'"
        }
        else {
            WriteJobLog "Kein Mapping für Department '$Department' gefunden – überspringe OG-Gruppe" "WARN"
        }

        if ($ogName) {
            try {
                Add-ADGroupMember -Identity $ogName -Members $UserID -ErrorAction Stop
                WriteJobLog "$UserID -> OG-Gruppe '$ogName'" "SUCCESS"
            }
            catch {
                WriteJobLog "Fehler beim Hinzufügen zu '$ogName': $($_.Exception.Message)" "WARN"
            }
        }
    }
    if ($ProgressCallback) { & $ProgressCallback 75 }
    # ------------------------------------------------------------------------

    # ------------------------------------------------
    # Schritt 10: Referats-Verteiler
    # ------------------------------------------------
    WriteJobLog "Suche Referat-Verteiler für Department '$Department'."
    if ($ProgressCallback) { & $ProgressCallback 75 }

    if ($Department -and $Department.Trim() -ne "") {
        # Beispiel: ReferatVerteiler.json auslesen (Pfad anpassen!)
        $referatVerteilerPath = "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v5\ReferatVerteiler.json"
        if (Test-Path $referatVerteilerPath) {
            try {
                $referatVerteiler = Get-Content -Path $referatVerteilerPath -Raw | ConvertFrom-Json
                $referatGroup = "Verteiler - Referat " + $Department.Trim()
                $foundGroup = $referatVerteiler | Where-Object { $_.Name -ieq $referatGroup }
                if ($foundGroup) {
                    try {
                        Add-ADGroupMember -Identity $foundGroup.Name -Members $UserID -ErrorAction Stop
                        WriteJobLog "$UserID -> Referats-Verteiler '$referatGroup'"
                    }
                    catch {
                        WriteJobLog "Fehler beim Hinzufügen zu Referats-Gruppe '$referatGroup': $($_.Exception.Message)" "WARN"
                    }
                }
            }
            catch {
                WriteJobLog "Fehler beim Laden ReferatVerteiler.json: $($_.Exception.Message)" "ERROR"
            }
        }
    }

    # ------------------------------------------------
    # Schritt 11: CONET-Prefix für DisplayName, falls gewünscht
    # ------------------------------------------------
    if ($isConet -eq "j") {
        # Beispiel-Funktion: Set-ConetDisplayName -UserID ...
        WriteJobLog "Setze CONET-DisplayName für $UserID"
        Set-ConetDisplayName -UserID $UserID
    }
    if ($ProgressCallback) { & $ProgressCallback 80 }

    # ------------------------------------------------
    # Schritt 12: Eventuelles Entfernen von Verteiler-/OG-Gruppen bei Azubi/Praktikant/Hospitant
    # ------------------------------------------------
    if ($roleSelection -in @("Azubi", "Praktikant", "Hospitant")) {
        WriteJobLog "Entferne Verteiler-Abteilung / Verteiler-Referat / OG-Gruppen für $UserID (Rolle=$roleSelection)"
        if ($ProgressCallback) { & $ProgressCallback 82 }

        try {
            $currentGroups = (Get-ADUser -Identity $UserID -Properties MemberOf).MemberOf
            if ($currentGroups) {
                foreach ($grpDN in $currentGroups) {
                    $grpObj  = Get-ADGroup $grpDN -Properties SamAccountName
                    $grpName = $grpObj.SamAccountName
                    
                    # Prüfe, ob es sich um Abteilung/Referat/OG-Gruppen handelt
                    if ($grpName -like "Verteiler - Abteilung*" -or
                        $grpName -like "Verteiler - Referat*"   -or
                        $grpName -like "OG*") {
                        
                        WriteJobLog "Entferne $UserID aus Gruppe '$grpName' (Rolle=$roleSelection)"
                        Remove-ADGroupMember -Identity $grpName -Members $UserID -Confirm:$false -ErrorAction SilentlyContinue
                    }
                }
            }
        }
        catch {
            WriteJobLog "Fehler beim Entfernen aus Abteilung/Referat/OG-Gruppen: $($_.Exception.Message)" "ERROR"
        }
    }

    # ------------------------------------------------
    # Schritt 13: Zusätzliche Gruppen
    # ------------------------------------------------
    if ($AdditionalGroups -and $AdditionalGroups.Count -gt 0) {
        foreach ($grp in $AdditionalGroups) {
            try {
                Add-ADGroupMember -Identity $grp -Members $UserID -ErrorAction Stop
                WriteJobLog "$UserID -> Zusatz-Gruppe '$grp'"
            }
            catch {
                WriteJobLog "Fehler Zusatz-Gruppe '$grp': $($_.Exception.Message)" "WARN"
            }
        }
    }
    if ($ProgressCallback) { & $ProgressCallback 85 }

    # ------------------------------------------------
    # Schritt 14: ExtensionAttributes
    WriteJobLog "Setze ExtensionAttributes …" "INFO"
    if ($ProgressCallback) { & $ProgressCallback 87 }

    Set-ExtensionAttributes `
        -UserID          $UserID `
        -gender          $gender `
        -isIVBB          $isIVBB `
        -isGVPL          $isGVPL `
        -isPhonebook     $isPhonebook `
        -isResMailbox    $isResMailbox `
        -isExternAccount $isExternAccount `
        -isVIP           $isVIP `
        -isFutureUser    $false

    # Schritt 14: abgeschlossen
    if ($ProgressCallback) { & $ProgressCallback 90 }

    # ------------------------------------------------
    # Schritt 15: Mailbox (optional erstellen)
    # ------------------------------------------------
    if ($makeMailbox -eq "j") {
        WriteJobLog "Erstelle Mailbox: $UserID @ $Site"
        if ($ProgressCallback) { & $ProgressCallback 92 }

        # Wenn Standort KTR, dann Mailbox in STR anlegen (so die Vorgabe)
        $mailboxSite = if ($Site -eq "KTR") { "STR" } else { $Site }
        $mailParams = @{
            site                   = $mailboxSite
            accounts               = @($UserID)
            hiddenFromAddressLists = $false
        }
        if ($isExtern)              { $mailParams["extern"] = $true }
        if ($isVerstecken)          { $mailParams["hiddenFromAddressLists"] = $true }
        if ($isExternAccount -eq "j") { $mailParams["externAccount"] = $true }

        Create-Mailbox @mailParams
    }

    # ------------------------------------------------
    # Schritt 16: HomeDir/Profile und Ordnerstruktur
    # ------------------------------------------------
    # Dieser Teil ruft (u. a.) ggf. dfsutil auf, wo der Fehler auftreten könnte.
    WriteJobLog "Setze HomeDir und Profile für $UserID..."
    if ($ProgressCallback) { & $ProgressCallback 95 }

    # Erstelle Ordnerstruktur (lokale Funktion oder externes Skript)
    WriteJobLog "Erstelle Ordnerstruktur für $UserID @ $Site"
    create-FolderStructure -UserID $UserID -Site $Site 

    # Setzt FilePermissions asynchron
    WriteJobLog "Set-FilePermissions asynchron für $UserID @ $Site"
    $jobPerm = set-FilePermissions -UserID $UserID -Site $Site 

    # Ownership & DFS-Berechtigung
    WriteJobLog "Ownership & DFS-Berechtigung für $UserID @ $Site"
    set-FolderOwnership -UserID $UserID -Site $Site



    # Schritt 17: AD-Objekt in die richtige OU verschieben
    WriteJobLog "Prüfe OU-Verschiebung für $UserID"
    try {
        Move-UserToTargetOU `
            -UserID   $UserID `
            -Gender   $gender `
            -Role     $roleSelection `
            -IsConet  $isConet `
            -IsExtern $isExtern
        WriteJobLog "OU-Verschiebung geprüft/ausgeführt." "INFO"
    }
    catch {
        WriteJobLog $_ "ERROR"
        throw
    }


    # ------------------------------------------------
    # Schritt 18: Abschluss / Login-Daten
    # ------------------------------------------------
    if ($created -and $newData) {
        # Wenn der User neu erstellt wurde, haben wir evtl. das Klartext-Passwort
        $loginInfo = "Benutzer: $UserID – Passwort: $($newData.passwordClear)"
        WriteJobLog "Login-Daten: $loginInfo" "SUCCESS"

        if (-not $InJob) {
            [System.Windows.MessageBox]::Show(
                "Benutzer $UserID wurde angelegt.`nLogin-Daten:`n$($newData.passwordClear)",
                "Fertig", 
                [System.Windows.MessageBoxButton]::OK
            ) | Out-Null
        }
    }
    else {
        # User war bereits vorhanden / Update
        if (-not $InJob) {
            [System.Windows.MessageBox]::Show(
                "Benutzer $UserID wurde erfolgreich erstellt/aktualisiert.",
                "Fertig", 
                [System.Windows.MessageBoxButton]::OK
            ) | Out-Null
        }
    }

    WriteJobLog "Fertig. Benutzer $UserID." "SUCCESS"

    # Schritt 17 (Ende): Fortschritt abschließen
    if ($ProgressCallback) { & $ProgressCallback 100 }
}

# ------------------------------------------------
# Modul-Export
# ------------------------------------------------
Export-ModuleMember -Function ProcessUserCreation, Convert-DepartmentShort
