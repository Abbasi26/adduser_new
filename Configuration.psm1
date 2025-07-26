# Configuration.psm1  – _single source of truth for all "environment" constants_
# IMPORTANT:  _never_ add business logic here – only data and tiny helpers.

$Script:Config = @{
    # --------------------------------------------------------------------
    # === UNC / local folders ==================================================
    ToolRoot               = '\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v22'
    OrgDataRSP             = '\\RSPSVFIL02\E$\Freigaben\OrgDataRSP\'
    OrgDataSTR             = '\\STRSVFIL02\E$\Freigaben\OrgDataSTR\'
    ReferatVerteilerJson   = '\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v5\ReferatVerteiler.json'
    LogoPath               = '\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\ToolBox\Res\BMUV_logo.jpg'

    # log / audit
    ActivateAccountLog     = '\\rspsvtask01.office.dir\c$\Scripte\Accounting-Logs\Log\activateAccountLog.txt'
    ActivateAccountData    = '\\rspsvtask01.office.dir\c$\Scripte\Accounting-Logs\Data\activateAccount.dat'
    ActivateAccountXmlLog  = '\\rspsvtask01.office.dir\c$\Scripte\Accounting-Logs\Log\activateAccountLog.log'
    LoginLog               = '\\rspsvtask01.office.dir\c$\Scripte\Accounting-Logs\Log\LoginLog.txt'

    # exchange / e-mail
    ExchangePSSessionUri   = 'http://rspsvexch12.office.dir/PowerShell/'
    RetentionPolicy        = 'BMU Benutzerpostfach Archivierungsrichtlinie'

    # JSON meta files used by the GUI
    AmtsJson               = 'amtsbezeichnungen.json'
    LaufJson               = 'laufbahnen.json'
    FunktionenJson         = 'Funktionen.json'
    SonderJson             = 'sonder.json'
    DeptsLongOGJson        = 'departments_long_oggroups.json'
    Depts2Json             = 'departments2.json'
    MainGuiXaml            = 'MainGUI.xaml'
    PreviewWindowXaml      = 'PreviewWindow.xaml'
    MassPreviewWindowXaml  = 'MassPreviewWindow.xaml'
}

function Get-Path {
    <#
        .SYNOPSIS
            Returns a resolved path/URI from the central configuration.

        .EXAMPLE
            Get-Path ToolRoot
    #>
    param (
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Name
    )
    if (-not $Script:Config.ContainsKey($Name)) {
        throw "Configuration key '$Name' not found."
    }
    return $Script:Config[$Name]
}

Export-ModuleMember -Variable Config -Function Get-Path

# If you need to override anything for dev → prod, dot-source a file named Configuration.Local.ps1 after importing this module and patch $Config.
