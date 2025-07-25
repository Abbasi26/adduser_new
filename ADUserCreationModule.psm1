# ADUserCreationModule.psm1

function new-ADAccountSettings {
    param(
        [Parameter(Mandatory=$true)][string]$newUserID,
        [Parameter(Mandatory=$false)][string]$givenName,
        [Parameter(Mandatory=$false)][string]$lastName,
        [Parameter(Mandatory=$false)][string]$expirationDate = "U",
        [Parameter(Mandatory=$false)][string]$userDescription = ""
    )

    Write-Log -Message "new-ADAccountSettings für $newUserID" -Color "blue"

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
    } #>

    if ($global:AppConfig.LogInLog) {
        "$($newUserData.UserloginName) : $userLoginPassword" | Out-File -FilePath $global:AppConfig.LogInLog -Append
    }

    Write-Log -Message "new-ADAccountSettings - Datenobjekt erstellt für $newUserID" -Color "green"
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

        Write-Log -Message "Benutzer $($newUserData.UserloginName) erfolgreich erstellt." -Color "green"
        # OPTIONAL: returning a small customobject
        return [pscustomobject]@{
            Success = $true
            Message = "Benutzer erstellt"
            PasswordClear = $newUserData.passwordClear
        }
    }
    catch {
        Write-Log -Message "Fehler Set-NewADAccount: $($_.Exception.Message)" -Color "red"
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
            Write-Log -Message "Anzeigename mit 'CONET ' erweitert: $newDisp" -Color "green"
            try {
                Rename-ADObject -Identity $adUser.DistinguishedName -NewName $newDisp -ErrorAction Stop
                Write-Log -Message "CN (Common Name) mit 'CONET ' erweitert: $newDisp" -Color "green"
            }
            catch {
                Write-Log -Message "Fehler beim Ändern des CN: $_" -Color "red"
            }
        }
    }
    else {
        Write-Log -Message "Benutzer $UserID wurde nicht gefunden." -Color "red"
    }
}



Export-ModuleMember -Function new-ADAccountSettings, Set-NewADAccount, Set-ConetDisplayName