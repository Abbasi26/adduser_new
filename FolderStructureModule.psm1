# ---------------------------------------------
# FolderStructureModule.psm1
# ---------------------------------------------

# Hilfsfunktion, um den Pfad zu dfsutil.exe
# je nach 32-/64-Bit-Situation sicher zu ermitteln.
function Get-DfsUtilPath {
    # Prüft, ob der aktuelle Prozess 64-Bit ist
    if ([Environment]::Is64BitProcess) {
        # Bei 64-Bit-Prozess liegt dfsutil regulär in System32
        $path = Join-Path $env:windir "System32\dfsutil.exe"
    }
    else {
        # Bei 32-Bit-Prozess auf 64-Bit-OS muss man "Sysnative" verwenden,
        # damit nicht auf SysWOW64 umgeleitet wird.
        $sysnative = Join-Path $env:windir "Sysnative\dfsutil.exe"
        if (Test-Path $sysnative) {
            $path = $sysnative
        }
        else {
            # Falls wir ein echtes 32-Bit-Windows haben (oder Sysnative nicht existiert),
            # dann bleibt System32 der korrekte Ort:
            $path = Join-Path $env:windir "System32\dfsutil.exe"
        }
    }
    if (!(Test-Path $path)) {
        throw "dfsutil.exe wurde nicht gefunden: $path"
    }
    return $path
}

function create-FolderStructure {
    param(
        [Parameter(Mandatory=$true)][string]$UserID,
        [Parameter(Mandatory=$true)][string]$Site
    )

    Write-Log -Message "Erstelle Ordnerstruktur für Benutzer $UserID (@$Site)" -Color "blue"

    if ($Site -eq "RSP") {
        $fileServerShare = "RSPSVFIL01.office.dir"
        $ProfilesFolder  = "\\$fileServerShare\ProfilesRSP$\$UserID"
        $AppDataFolder   = "\\$fileServerShare\AppDataRSP$\$UserID"
        $UserDataFolder  = "\\$fileServerShare\UserDataRSP$\$UserID"
    }
    else {
        $fileServerShare = "STRSVFIL01.office.dir"
        $ProfilesFolder  = "\\$fileServerShare\ProfilesSTR$\$UserID"
        $AppDataFolder   = "\\$fileServerShare\AppDataSTR$\$UserID"
        $UserDataFolder  = "\\$fileServerShare\UserDataSTR$\$UserID"
    }

    # Lege die Verzeichnisstruktur an (wenn nicht vorhanden)
    if (!(Test-Path "$ProfilesFolder\Profile"))      { mkdir "$ProfilesFolder\Profile"      | Out-Null }
    if (!(Test-Path "$ProfilesFolder\Profile.V2"))   { mkdir "$ProfilesFolder\Profile.V2"   | Out-Null }
    if (!(Test-Path "$ProfilesFolder\Profile.V6"))   { mkdir "$ProfilesFolder\Profile.V6"   | Out-Null }
    if (!(Test-Path "$ProfilesFolder\TSProfile"))    { mkdir "$ProfilesFolder\TSProfile"    | Out-Null }
    if (!(Test-Path "$ProfilesFolder\TSProfile.V2")) { mkdir "$ProfilesFolder\TSProfile.V2" | Out-Null }

    if (!(Test-Path "$AppDataFolder\Local"))   { mkdir "$AppDataFolder\Local"   | Out-Null }
    if (!(Test-Path "$AppDataFolder\Roaming")) { mkdir "$AppDataFolder\Roaming" | Out-Null }

    if (!(Test-Path "$UserDataFolder\Contacts"))   { mkdir "$UserDataFolder\Contacts"   | Out-Null }
    if (!(Test-Path "$UserDataFolder\Desktop"))    { mkdir "$UserDataFolder\Desktop"    | Out-Null }
    if (!(Test-Path "$UserDataFolder\Documents"))  { mkdir "$UserDataFolder\Documents"  | Out-Null }
    if (!(Test-Path "$UserDataFolder\Downloads"))  { mkdir "$UserDataFolder\Downloads"  | Out-Null }
    if (!(Test-Path "$UserDataFolder\Favorites"))  { mkdir "$UserDataFolder\Favorites"  | Out-Null }
    if (!(Test-Path "$UserDataFolder\Links"))      { mkdir "$UserDataFolder\Links"      | Out-Null }
    if (!(Test-Path "$UserDataFolder\Music"))      { mkdir "$UserDataFolder\Music"      | Out-Null }
    if (!(Test-Path "$UserDataFolder\Pictures"))   { mkdir "$UserDataFolder\Pictures"   | Out-Null }
    if (!(Test-Path "$UserDataFolder\Saved Games")){ mkdir "$UserDataFolder\Saved Games"| Out-Null }
    if (!(Test-Path "$UserDataFolder\Searches"))   { mkdir "$UserDataFolder\Searches"   | Out-Null }
    if (!(Test-Path "$UserDataFolder\Videos"))     { mkdir "$UserDataFolder\Videos"     | Out-Null }
    if (!(Test-Path "$UserDataFolder\Application Data")) { mkdir "$UserDataFolder\Application Data" | Out-Null }

    Write-Log -Message "Ordner erstellt für $UserID." -Color "green"

    # DFS-Verlinkung herstellen
    $DfsNameSpaceBenutzer = "\\office.dir\Benutzer$"
    $DfsNameSpaceDFS      = "\\office.dir\Files\Benutzer"

    # Hole korrekten Pfad zu dfsutil (32/64 Bit)
    $dfsUtil = Get-DfsUtilPath

    if (!(Test-Path "$DfsNameSpaceBenutzer\$UserID\Profiles")) {
        & $dfsUtil link add "$DfsNameSpaceBenutzer\$UserID\Profiles" "$ProfilesFolder" | Out-Null
    }
    if (!(Test-Path "$DfsNameSpaceBenutzer\$UserID\AppData")) {
        & $dfsUtil link add "$DfsNameSpaceBenutzer\$UserID\AppData" "$AppDataFolder" | Out-Null
    }
    if (!(Test-Path "$DfsNameSpaceBenutzer\$UserID\UserData")) {
        & $dfsUtil link add "$DfsNameSpaceBenutzer\$UserID\UserData" "$UserDataFolder" | Out-Null
    }
    & $dfsUtil link add "$DfsNameSpaceDFS\$UserID" "$DfsNameSpaceBenutzer\$UserID" | Out-Null

    Write-Log -Message "DFS-Links erstellt für $UserID." -Color "green"

    $site1 = $site+"SVFIL01"
    # (Achtung: Pfad hier ggf. anpassen, falls es den wirklich so gibt)
    $source = "\\$fileServerShare\$site1\e$\Freigaben\AppDataRSP\_DefaultChromeProfile\"
    robocopy $source "$AppDataFolder\Roaming\google\Chrome\User Data\Default" /E | Out-Null
}

function set-FilePermissions {
    param(
        [Parameter(Mandatory=$true)][string]$UserID, 
        [Parameter(Mandatory=$true)][string]$Site
    )

    Write-Log -Message "Starte set-FilePermissions (asynchron) für $UserID @ $Site" -Color "blue"
    
    # Wir starten einen Job, damit das Setzen der Rechte im Hintergrund läuft:
    $job = Start-Job -ScriptBlock {
        param($UserID, $Site)

        # Hier ggf. MyWrite-Log durch Write-Host ersetzen oder aus dem Modul mitgeben
        Write-Host "Beginne ACL-Setzung für $UserID @ $Site ..."

        if ($Site -eq "RSP") {
            $ProfilesFolder = "\\RSPSVFIL01.office.dir\ProfilesRSP$\$UserID"
            $AppDataFolder  = "\\RSPSVFIL01.office.dir\AppDataRSP$\$UserID"
            $UserDataFolder = "\\RSPSVFIL01.office.dir\UserDataRSP$\$UserID"
        }
        elseif ($Site -in @("KRA","STR","KTR")) {
            $ProfilesFolder = "\\STRSVFIL01.office.dir\ProfilesSTR$\$UserID"
            $AppDataFolder  = "\\STRSVFIL01.office.dir\AppDataSTR$\$UserID"
            $UserDataFolder = "\\STRSVFIL01.office.dir\UserDataSTR$\$UserID"
        }
        else {
            Write-Host "Ungültige Site: $Site"
            return
        }

        $DomainUserX = "BMUDOM\$UserID"

        # ACLs zurücksetzen und neu setzen
        icacls "$ProfilesFolder" /reset /T /Q
        icacls "$ProfilesFolder" /grant "$($DomainUserX):(OI)(CI)(IO)F" /T /Q
        icacls "$ProfilesFolder" /grant "$($DomainUserX):(OI)(CI)(RX)" /T /Q

        icacls "$AppDataFolder" /reset /T /Q
        icacls "$AppDataFolder" /grant "$($DomainUserX):(OI)(CI)(IO)F" /T /Q
        icacls "$AppDataFolder" /grant "$($DomainUserX):(OI)(CI)(RX)" /T /Q

        icacls "$UserDataFolder" /reset /T /Q
        icacls "$UserDataFolder" /grant "$($DomainUserX):(OI)(CI)(IO)F" /T /Q
        icacls "$UserDataFolder" /grant "$($DomainUserX):(OI)(CI)(RX)" /T /Q

        Start-Sleep -Seconds 2
        # Ownership (Besitzrechte) setzen
        icacls "$UserDataFolder" /setowner "$DomainUserX" /T /Q
        icacls "$AppDataFolder"  /setowner "$DomainUserX" /T /Q
        icacls "$ProfilesFolder" /setowner "$DomainUserX" /T /Q

        Write-Host "Done setting ACLs for $UserID @ $Site asynchronously."
    } -ArgumentList $UserID,$Site

    return $job
}

function set-FolderOwnership {
    param(
        [Parameter(Mandatory=$true)][string]$UserID, 
        [Parameter(Mandatory=$true)][string]$Site
    )

    Write-Log -Message "Setze Ownership & DFS-Berechtigung (synchron) für $UserID @ $Site" -Color "blue"

    $DfsNameSpaceDFS       = "\\office.dir\Files\Benutzer"
    $FileServiceAdminGroup = "SG-Fileservices-Admins"
    $FileServiceBmuvMa     = "SG-Fileservices-BMUV-MA"

    # Hier ebenfalls dfsutil über den richtigen Pfad aufrufen
    $dfsUtil = Get-DfsUtilPath

    & $dfsUtil Property SD Reset   "$DfsNameSpaceDFS\$UserID"
    Start-Sleep -Seconds 1
    & $dfsUtil Property SD Control "$DfsNameSpaceDFS\$UserID" protect
    & $dfsUtil Property SD Grant   "$DfsNameSpaceDFS\$UserID" "BMUDOM\$UserID:M"
    & $dfsUtil Property SD Grant   "$DfsNameSpaceDFS\$UserID" "$FileServiceAdminGroup:M"
    & $dfsUtil Property SD Grant   "$DfsNameSpaceDFS\$UserID" "$FileServiceBmuvMa:M"

    Write-Log -Message "Ownership + DFS-Berechtigungen gesetzt (synchron) für $UserID." -Color "green"
}

Export-ModuleMember -Function create-FolderStructure, set-FilePermissions, set-FolderOwnership
