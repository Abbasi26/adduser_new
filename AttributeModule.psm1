# AttributeModule.psm1
function new-ADAccountSettings {
    param(
        [Parameter(Mandatory=$true)][string]$newUserID,
        [Parameter(Mandatory=$false)][string]$givenName,
        [Parameter(Mandatory=$false)][string]$lastName,
        [Parameter(Mandatory=$false)][string]$expirationDate = "U",
        [Parameter(Mandatory=$false)][string]$userDescription = ""
    )

    MyWrite-Log "new-ADAccountSettings für $newUserID" -Color "blue"

    $UserDisplayName   = if ($lastName -and $givenName) { "$lastName, $givenName" } else { $newUserID }
    $userLoginPassword = "BmuB_$(Get-Date -Format ddHHmm)"

    $newUserData = @{
        displayName       = $UserDisplayName
        givenName         = $givenName
        lastName          = $lastName
        UserloginName     = $newUserID
        userLoginPassword = (ConvertTo-SecureString -AsPlainText $userLoginPassword -Force)
        reqPasswordChange = $true
        OUPath            = "OU=Benutzer,OU=AnwenderRes,DC=office,DC=dir"
        expirationDate    = $null
        description       = $userDescription
        UserPrincipalName = "$($newUserID)@office.dir"
        passwordClear     = $userLoginPassword
    }

    <#if ($expirationDate -ne "U") {
        $newUserData.description     = "befr. bis $expirationDate; " + $userDescription
        $newUserData.expirationDate = ([datetime](Get-Date ($expirationDate + " 00:00:00")).AddDays(1))
    }#>

    if ($global:AppConfig.LogInLog) {
        "$($newUserData.UserloginName) : $userLoginPassword" | Out-File -FilePath $global:AppConfig.LogInLog -Append
    }

    MyWrite-Log "new-ADAccountSettings - Datenobjekt erstellt für $newUserID" -Color "green"
    return $newUserData
}

function Set-NewADAccount { 
    param([Parameter(Mandatory=$true)][hashtable]$newUserData)
    try {
        New-ADUser `
            -Name                  $newUserData.displayName `
            -UserPrincipalName     $newUserData.UserPrincipalName `
            -Description           $newUserData.description `
            -GivenName             $newUserData.givenName `
            -Surname               $newUserData.lastName `
            -Enabled               $true `
            -DisplayName           $newUserData.displayName `
            -SamAccountName        $newUserData.UserloginName `
            -AccountPassword       $newUserData.userLoginPassword `
            -ChangePasswordAtLogon $newUserData.reqPasswordChange `
            -AccountExpirationDate $newUserData.expirationDate `
            -Path                  $newUserData.OUPath `
            -ProfilePath           "\\office.dir\Files\Benutzer\$($newUserData.UserloginName)\Profiles\Profile"

        MyWrite-Log "Benutzer $($newUserData.UserloginName) erfolgreich erstellt." -Color "green"
        return [pscustomobject]@{
            Success = $true
            Message = "Benutzer erstellt"
            PasswordClear = $newUserData.passwordClear
        }
    }
    catch {
        MyWrite-Log "Fehler Set-NewADAccount: $($_.Exception.Message)" -Color "red"
        return [pscustomobject]@{
            Success = $false
            Message = $_.Exception.Message
        }
    }
}

function Set-ConetDisplayName {
    param([Parameter(Mandatory = $true)][string]$UserID)

    $adUser = Get-ADUser -Identity $UserID -Properties DisplayName, DistinguishedName, CN -ErrorAction SilentlyContinue
    if ($adUser) {
        if (-not $adUser.DisplayName.StartsWith("CONET ")) {
            $newDisp = "CONET " + $adUser.DisplayName
            Set-ADUser -Identity $UserID -DisplayName $newDisp -ErrorAction SilentlyContinue
            MyWrite-Log "Anzeigename mit 'CONET ' erweitert: $newDisp" -Color "green"
            try {
                Rename-ADObject -Identity $adUser.DistinguishedName -NewName $newDisp -ErrorAction Stop
                MyWrite-Log "CN (Common Name) mit 'CONET ' erweitert: $newDisp" -Color "green"
            }
            catch {
                MyWrite-Log "Fehler beim Ändern des CN: $_" -Color "red"
            }
        }
    }
    else {
        MyWrite-Log "Benutzer $UserID wurde nicht gefunden." -Color "red"
    }
}

Export-ModuleMember -Function new-ADAccountSettings, Set-NewADAccount, Set-ConetDisplayName

function Std-Attributes {
    param([Parameter(Mandatory=$true)]$userAttributes)
    MyWrite-Log "Std-Attributes für $($userAttributes.id)" -Color "blue"
    $stdparams = @{
        "ID"                   = $userAttributes.id
        "extensionAttribute13" = "x"
        "extensionAttribute14" = "x"
        "extensionAttribute3"  = ""
    }
    return $stdparams
}

function Set-ADExAttributes {
    param([Parameter(Mandatory=$true)]$userAttributes)
    try {
        MyWrite-Log "Set-ADExAttributes: $($userAttributes.id)" -Color "blue"
        Set-ADUser -Identity $userAttributes.id -Clear "extensionAttribute3","extensionAttribute13","extensionAttribute14"
        $params = @{}
        if ($userAttributes.extensionAttribute3)  { $params["extensionAttribute3"]  = "IVBB" }
        if ($userAttributes.extensionAttribute13) { $params["extensionAttribute13"] = $userAttributes.extensionAttribute13 }
        if ($userAttributes.extensionAttribute14) { $params["extensionAttribute14"] = $userAttributes.extensionAttribute14 }
        if ($params.Count -gt 0) {
            Set-ADUser -Identity $userAttributes.id -Add $params
            MyWrite-Log "Attribute gesetzt: $($params.Keys)" -Color "green"
        }
        return $true
    }
    catch {
        MyWrite-Log "Fehler Set-ADExAttributes: $($_.Exception.Message)" -Color "red"
        return $false
    }
}

function check-date {
    param(
        [Parameter(Mandatory=$true)][string]$exception,
        [Parameter(Mandatory=$true)][string]$date
    )
    if ($date -ne $exception) {
        try {
            [datetime](Get-Date $date -Format "dd.MM.yyyy") | Out-Null
            return $true
        }
        catch {
# Set-ExtensionAttributes
function Set-ExtensionAttributes {
    param(
        [string]$UserID,
        [string]$gender,
        [string]$isIVBB,
        [string]$isGVPL,
        [bool]  $isPhonebook,
        [string]$isResMailbox,
        [string]$isExternAccount,
        [string]$isVIP,
        [bool]  $isFutureUser = $false
    )

    $ext = @{}

    if ($isIVBB  -eq 'j') { $ext.extensionAttribute3  = 'IVBB' }
    if ($isGVPL  -eq 'j' -and -not $isFutureUser) { $ext.extensionAttribute13 = 'x' }
    if ($isPhonebook      -and -not $isFutureUser) { $ext.extensionAttribute14 = 'x' }
    if ($isResMailbox -eq 'j') { $ext.extensionAttribute7  = 'ResourceMB' }
    if ($isExternAccount -eq 'j') { $ext.extensionAttribute11 = 'x' }

    if ($isFutureUser) {
        $ext.extensionAttribute13 = 'x'
        $ext.extensionAttribute14 = 'x'
    }

    switch ($gender) {
        'Mann'  { $ext.extensionAttribute4 = 'Herr' }
        'Frau'  { $ext.extensionAttribute4 = 'Frau' }
        'Divers'{ $ext.extensionAttribute4 = ''     }
        'Nicht nat\xFCrliche Person (NNP)' {
            $ext.extensionAttribute4 = ''
            $ext.extensionAttribute5 = '1'
        }
    }

    if ($ext.Count) {
        Set-ADUser -Identity $UserID -Replace $ext -ErrorAction Stop
        WriteJobLog "ExtensionAttributes gesetzt: $($ext.Keys -join ', ')" "SUCCESS"
    }

    if ($isVIP -eq 'j') {
        try {
            Set-ADUser -Identity $UserID -Add @{ pager = 'VIP' } -ErrorAction Stop
            WriteJobLog "VIP (Quota) gesetzt." "INFO"
        } catch {
            WriteJobLog "Fehler VIP: $($_.Exception.Message)" "WARN"
        }
    }
}

Export-ModuleMember -Function Std-Attributes, Set-ADExAttributes, check-date, set-attributes, Set-ExtensionAttributes
            return $false
        }
    }
    return $true
}

function set-attributes {
    param([Parameter(Mandatory=$true)][string]$attribute)
    MyWrite-Log "set-attributes $attribute" -Color "blue"
    if ($attribute -eq "j") {
        return @{ value="x"; display="Ja" }
    }
    elseif ($attribute -eq "n") {
        return @{ value=$null; display="Nein" }
    }
    return $null
}

Export-ModuleMember -Function Std-Attributes, Set-ADExAttributes, check-date, set-attributes
