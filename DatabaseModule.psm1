# DatabaseModule.psm1
function set-database {
    param([Parameter(Mandatory=$true)][string]$pathToFile)
    MyWrite-Log "set-database: Prüfe $pathToFile" -Color "blue"
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
            MyWrite-Log "Datenbank angelegt: $pathToFile" -Color "green"
            return $true
        }
        catch {
            MyWrite-Log "Fehler set-database: $($_.Exception.Message)" -Color "red"
            return $false
        }
    }
    return $true
}

function check-Database {
    param([Parameter(Mandatory=$true)][string]$UserID)
    MyWrite-Log "check-Database: Prüfe DB-Eintrag für $UserID" -Color "blue"
    try {
        if (Test-Path $global:AppConfig.FilePath) {
            [xml]$doc = Get-Content $global:AppConfig.FilePath
            $users = $doc.SelectSingleNode("//User[translate(@ID, 'A-Z', 'a-z') = translate('$UserID', 'A-Z', 'a-z')]")
            if ($users) { return $users }
            else        { return $false }
        }
    }
    catch {
        MyWrite-Log "Fehler check-Database: $($_.Exception.Message)" -Color "red"
        return $false
    }
    return $false
}

function delete-record {
    param([Parameter(Mandatory=$true)][string]$UserID)
    MyWrite-Log "delete-record: Lösche $UserID in DB" -Color "blue"
    try {
        [xml]$doc = Get-Content $global:AppConfig.FilePath
        $users = $doc.SelectSingleNode("//User[translate(@ID, 'A-Z', 'a-z') = translate('$UserID', 'A-Z', 'a-z')]")
        if ($users) {
            [Void]$users.ParentNode.RemoveChild($users)
            $doc.Save($global:AppConfig.FilePath)
            MyWrite-Log "Tupel $UserID gelöscht." -Color "green"
            return $true
        }
        else {
            MyWrite-Log "Tupel $UserID nicht gefunden." -Color "red"
            return $false
        }
    }
    catch {
        MyWrite-Log "Fehler delete-record: $($_.Exception.Message)" -Color "red"
        return $false
    }
}

function append-database {
    param(
        [Parameter(Mandatory=$true)]$userAttributes,
        [Parameter(Mandatory=$true)]$ticketNr
    )
    MyWrite-Log "append-database: $($userAttributes.id)" -Color "blue"
    if (set-database -pathToFile $global:AppConfig.FilePath) {
        try {
            [xml]$doc = Get-Content $global:AppConfig.FilePath
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
            $doc.Save($global:AppConfig.FilePath)
            MyWrite-Log "Datensatz für $($userAttributes.id) angehängt." -Color "green"
            return $true
        }
        catch {
            MyWrite-Log "Fehler append-database: $($_.Exception.Message)" -Color "red"
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
    MyWrite-Log "write-XMLLog: $($logUserAttributes.id) - $ticketNr" -Color "blue"
    if (set-database -pathToFile $global:AppConfig.XMLLogPath) {
        try {
            [xml]$logDoc = Get-Content $global:AppConfig.XMLLogPath
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
            $logDoc.Save($global:AppConfig.XMLLogPath)
            MyWrite-Log "write-XMLLog: Ticket $ticketNr eingetragen." -Color "green"
            return $true
        }
        catch {
            MyWrite-Log "Fehler write-XMLLog: $($_.Exception.Message)" -Color "red"
            return $false
        }
    }
    return $false
}

Export-ModuleMember -Function set-database, check-Database, delete-record, append-database, write-XMLLog
