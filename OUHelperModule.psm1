# OUHelperModule.psm1
# --------------------
# Contains helper functions for OU operations

function Move-UserToTargetOU {
    <#
        .SYNOPSIS
            Verschiebt einen Benutzer in die korrekte Ziel-OU
            anhand Conet-Flag, Gender, Extern-Status und Rolle.

        .PARAMETER UserID
            SamAccountName des Benutzers.

        .PARAMETER Gender
            'Mann' | 'Frau' | 'Divers' | 'Nicht natürliche Person (NNP)'

        .PARAMETER Role
            '', 'Azubi', 'Praktikant', 'Hospitant', 'Referendar'

        .PARAMETER IsConet
            'j' / 'n'

        .PARAMETER IsExtern
            [bool]  Extern-Checkbox aus GUI

        .PARAMETER DryRun
            Nur logging, kein echtes Move-ADObject.
    #>
    param(
        [string]$UserID,
        [string]$Gender,
        [string]$Role,
        [string]$IsConet,
        [bool]  $IsExtern,
        [switch]$DryRun
    )

    $rootOU = "OU=Benutzer,OU=AnwenderRes,DC=office,DC=dir"
    switch ($true) {
        { $IsConet -eq 'j' }                                                   { $child = 'GU-IT' }
        { $Gender  -eq 'Nicht natürliche Person (NNP)' }                       { $child = 'Funktionsaccounts' }
        { $IsExtern }                                                          { $child = 'Extern' }
        { $Role -in @('Azubi','Praktikant','Hospitant','Referendar') }         { $child = 'Referendare-Praktikanten-Hospitanten' }
        default                                                                { $child = '' }  # Standard
    }

    $targetOU = ($child ? "OU=$child,$rootOU" : $rootOU)

    try {
        $u = Get-ADUser -Identity $UserID -ErrorAction Stop
        $currentOU = ($u.DistinguishedName -split ',OU=',2)[1]

        if ($currentOU -ieq $targetOU) {
            Write-Verbose "[$UserID] bleibt in OU '$currentOU'"
            return
        }

        Write-Verbose "[$UserID] Move:  $currentOU  →  $targetOU"
        if (-not $DryRun) {
            Move-ADObject -Identity $u.DistinguishedName -TargetPath $targetOU -ErrorAction Stop
        }
    }
    catch {
        throw "Move-UserToTargetOU  –  $($_.Exception.Message)"
    }
}

Export-ModuleMember -Function Move-UserToTargetOU
