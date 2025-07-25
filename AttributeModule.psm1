# AttributeModule.psm1

function Std-Attributes {
    param([Parameter(Mandatory=$true)]$userAttributes)
    Write-Log -Message "Std-Attributes für $($userAttributes.id)" -Color "blue"
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
        Write-Log -Message "Set-ADExAttributes: $($userAttributes.id)" -Color "blue"
        Set-ADUser -Identity $userAttributes.id -Clear "extensionAttribute3","extensionAttribute13","extensionAttribute14"
        $params = @{}
        if ($userAttributes.extensionAttribute3)  { $params["extensionAttribute3"]  = "IVBB" }
        if ($userAttributes.extensionAttribute13) { $params["extensionAttribute13"] = $userAttributes.extensionAttribute13 }
        if ($userAttributes.extensionAttribute14) { $params["extensionAttribute14"] = $userAttributes.extensionAttribute14 }
        if ($params.Count -gt 0) {
            Set-ADUser -Identity $userAttributes.id -Add $params
            Write-Log -Message "Attribute gesetzt: $($params.Keys)" -Color "green"
        }
        return $true
    }
    catch {
        Write-Log -Message "Fehler Set-ADExAttributes: $($_.Exception.Message)" -Color "red"
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
        } catch {
            Write-Log -Message "Ungültiges Datum: $date" -Color Red
            return $false
        }
    }
    return $true
}

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
        'Nicht natürliche Person (NNP)' {
            $ext.extensionAttribute4 = ''
            $ext.extensionAttribute5 = '1'
        }
    }

    if ($ext.Count) {
        Set-ADUser -Identity $UserID -Replace $ext -ErrorAction Stop
        # Erfolgsmeldung konsistent über -msg und -Color ausgeben
        WriteJobLog -msg "ExtensionAttributes gesetzt: $($ext.Keys -join ', ')" -Color Green
    }

    if ($isVIP -eq 'j') {
        try {
            Set-ADUser -Identity $UserID -Add @{ pager = 'VIP' } -ErrorAction Stop
            WriteJobLog -msg "VIP (Quota) gesetzt." -Color Blue
        } catch {
            WriteJobLog -msg "Fehler VIP: $($_.Exception.Message)" -Color Orange
        }
    }
}


function set-attributes {
    param([Parameter(Mandatory=$true)][string]$attribute)
    Write-Log -Message "set-attributes $attribute" -Color "blue"
    if ($attribute -eq "j") {
        return @{ value="x"; display="Ja" }
    }
    elseif ($attribute -eq "n") {
        return @{ value=$null; display="Nein" }
    }
    return $null
}

Export-ModuleMember -Function Std-Attributes, Set-ADExAttributes, check-date, set-attributes
