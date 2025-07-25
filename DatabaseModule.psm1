# DatabaseModule.psm1

function Ensure-Directory {
    param([string]$Path)
    if (-not $Path) {
        throw "Pfad darf nicht NULL sein (aufrufende Funktion: $MyInvocation.MyCommand)"
    }
    $dir = [IO.Path]::GetDirectoryName($Path)
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}
function set-database {
    param([Parameter(Mandatory=$true)][string]$pathToFile)
    Write-Log -Message "set-database: Prüfe $pathToFile" -Color "blue"
    Ensure-Directory $pathToFile
    if (!(Test-Path $pathToFile)) {
        try {
            $XmlWriter = New-Object System.Xml.XmlTextWriter($pathToFile, $null)
            $XmlWriter.Formatting = "Indented"
            $XmlWriter.Indentation = 4
            $XmlWriter.WriteStartDocument()
            $XmlWriter.WriteStartElement("Users")
            $XmlWriter.WriteEndElement()
            $XmlWriter.WriteEndDocument()
            $XmlWriter.Flush()
            $XmlWriter.Close()
            Write-Log -Message "Datenbank angelegt: $pathToFile" -Color "green"
            return $true
        }
        catch {
            Write-Log -Message "Fehler set-database: $($_.Exception.Message)" -Color "red"
            return $false
        }
    }
    return $true
}

function check-Database {
    param([Parameter(Mandatory=$true)][string]$UserID)
    Write-Log -Message "check-Database: Prüfe DB-Eintrag für $UserID" -Color "blue"
    try {
        $dbPath = $global:AppConfig.Paths.FilePath
        if ($dbPath -and (Test-Path -LiteralPath $dbPath)) {
            [xml]$doc = Get-Content $dbPath
            $users = $doc.SelectSingleNode("//User[translate(@ID, 'A-Z', 'a-z') = translate('$UserID', 'A-Z', 'a-z')]")
            if ($users) { return $users }
            else        { return $false }
        }
    }
    catch {
        Write-Log -Message "Fehler check-Database: $($_.Exception.Message)" -Color "red"
        return $false
    }
    return $false
}

function delete-record {
    param([Parameter(Mandatory=$true)][string]$UserID)
    Write-Log -Message "delete-record: Lösche $UserID in DB" -Color "blue"
    try {
        $dbPath = $global:AppConfig.Paths.FilePath
        Ensure-Directory $dbPath
        [xml]$doc = Get-Content $dbPath
        $users = $doc.SelectSingleNode("//User[translate(@ID, 'A-Z', 'a-z') = translate('$UserID', 'A-Z', 'a-z')]")
        if ($users) {
            [Void]$users.ParentNode.RemoveChild($users)
            $doc.Save($dbPath)
            Write-Log -Message "Tupel $UserID gelöscht." -Color "green"
            return $true
        }
        else {
            Write-Log -Message "Tupel $UserID nicht gefunden." -Color "red"
            return $false
        }
    }
    catch {
        Write-Log -Message "Fehler delete-record: $($_.Exception.Message)" -Color "red"
        return $false
    }
}

function append-database {
    param(
        [Parameter(Mandatory=$true)]$userAttributes,
        [Parameter(Mandatory=$true)]$ticketNr
    )
    Write-Log -Message "append-database: $($userAttributes.id)" -Color "blue"
    $dbPath = $global:AppConfig.Paths.FilePath
    if (set-database -pathToFile $dbPath) {
        try {
            [xml]$doc = Get-Content $dbPath
            $user = $doc.CreateElement("User")
            $user.SetAttribute("ID", $userAttributes.id)
            $user.SetAttribute("Date", $userAttributes.datum)
            $user.SetAttribute("TicketNr", $ticketNr)

            $node = $user.AppendChild($doc.CreateElement("extensionAttribute13"))
            $node.AppendChild($doc.CreateTextNode($userAttributes.extensionAttribute13))
            $node = $user.AppendChild($doc.CreateElement("extensionAttribute3"))
            $node.AppendChild($doc.CreateTextNode($userAttributes.extensionAttribute3))
            $node = $user.AppendChild($doc.CreateElement("extensionAttribute14"))
            $node.AppendChild($doc.CreateTextNode($userAttributes.extensionAttribute14))

            $doc.DocumentElement.AppendChild($user)
            $doc.Save($dbPath)
            Write-Log -Message "Datensatz für $($userAttributes.id) angehängt." -Color "green"
            return $true
        }
        catch {
            Write-Log -Message "Fehler append-database: $($_.Exception.Message)" -Color "red"
            return $false
        }
    }
    return $false
}

function write-XMLLog {
    param(
        [Parameter(Mandatory=$true)]$logUserAttributes,
        [Parameter(Mandatory=$true)]$ticketNr
    )
    Write-Log -Message "write-XMLLog: $($logUserAttributes.id) - $ticketNr" -Color "blue"
    $xmlPath = $global:AppConfig.Paths.XMLLogPath
    if (set-database -pathToFile $xmlPath) {
        try {
            [xml]$logDoc = Get-Content $xmlPath
            $editor   = whoami
            $EditDate = Get-Date -Format "dd.MM.yyyy HH.mm"
            $logUser = $logDoc.CreateElement("Ticket")
            $logUser.SetAttribute("TicketNr", $ticketNr)

            $logNode = $logUser.AppendChild($logDoc.CreateElement("Editor"))
            $logNode.AppendChild($logDoc.CreateTextNode($editor))
            $logNode = $logUser.AppendChild($logDoc.CreateElement("EditDate"))
            $logNode.AppendChild($logDoc.CreateTextNode($EditDate))
            $logNode = $logUser.AppendChild($logDoc.CreateElement("UserName"))
            $logNode.AppendChild($logDoc.CreateTextNode($logUserAttributes.id))
            $logNode = $logUser.AppendChild($logDoc.CreateElement("StartDate"))
            $logNode.AppendChild($logDoc.CreateTextNode($logUserAttributes.datum))
            $logNode = $logUser.AppendChild($logDoc.CreateElement("extensionAttribute3"))
            $logNode.AppendChild($logDoc.CreateTextNode($logUserAttributes.extensionAttribute3))
            $logNode = $logUser.AppendChild($logDoc.CreateElement("extensionAttribute13"))
            $logNode.AppendChild($logDoc.CreateTextNode($logUserAttributes.extensionAttribute13))
            $logNode = $logUser.AppendChild($logDoc.CreateElement("extensionAttribute14"))
            $logNode.AppendChild($logDoc.CreateTextNode($logUserAttributes.extensionAttribute14))

            $logDoc.DocumentElement.AppendChild($logUser)
            $logDoc.Save($xmlPath)
            Write-Log -Message "write-XMLLog: Ticket $ticketNr eingetragen." -Color "green"
            return $true
        }
        catch {
            Write-Log -Message "Fehler write-XMLLog: $($_.Exception.Message)" -Color "red"
            return $false
        }
    }
    return $false
}

Export-ModuleMember -Function set-database, check-Database, delete-record, append-database, write-XMLLog
