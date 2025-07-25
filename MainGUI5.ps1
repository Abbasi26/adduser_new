#region Initialisierung und Imports
# ------------------------------------------------#
# 1) Initialisierung
# ------------------------------------------------#

# Funktion zum Verstecken der PowerShell-Konsole
function Hide-Console {
    $signature = @"
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
"@
    Add-Type -MemberDefinition $signature -Name "Win32Functions" -Namespace "Win32" -PassThru
    $consoleHandle = [Win32.Win32Functions]::GetConsoleWindow()
    [Win32.Win32Functions]::ShowWindow($consoleHandle, 0)  # 0 = Verstecken
}
Hide-Console

# Lade notwendige Assemblies für WPF und Windows Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Erstelle eine eigene WPF-Application-Instanz, falls nicht vorhanden
if (-not [System.Windows.Application]::Current) {
Import-Module "$Script:ToolRoot\LogModule.psm1" -Force

    $app = New-Object System.Windows.Application
    $global:CustomApplication = $app
}

# Lade Konfigurationsdateien
$global:AppConfig = Get-Content "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v22\config.json" -Raw -Encoding UTF8 | ConvertFrom-Json
$global:StdProfiles = Get-Content "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v22\stdprofiles.json" -Raw -Encoding UTF8 | ConvertFrom-Json
Import-Module "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v22\MailboxModule.psm1"
Import-Module "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v22\ProfileModule.psm1"

# Standard-Suchattribute für Gruppensuche
$global:SearchAttributes = @("cn")

# Importiere erforderliche Module
foreach ($module in $global:AppConfig.Modules) {
    $modulePath = $module
    if (Test-Path $modulePath) {
        Import-Module $modulePath -Force
    } else {
        Write-Host "FEHLER: Modul $modulePath konnte nicht gefunden werden."
    }
}

# Lade die Haupt-GUI aus XAML
[xml]$rawXaml = Get-Content $global:AppConfig.Paths.MainGUIXaml -Raw
$reader = New-Object System.Xml.XmlNodeReader($rawXaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
if (-not $window) {
    Write-Host "FEHLER: Konnte $($global:AppConfig.Paths.MainGUIXaml) nicht laden."
    exit
}



# Definiere C#-Klasse für Gruppen-Items
Add-Type -TypeDefinition @"
public class GroupItem {
    public string Name { get; set; }
    public bool IsChecked { get; set; }
}
"@ -Language CSharp

#endregion

#region UI-Hilfsfunktionen
# ------------------------------------------------#
# 2) UI-Hilfsfunktionen
# ------------------------------------------------#

# Funktion zum Ausführen von UI-Änderungen im Dispatcher-Thread
function Update-UI {
    param (
        [ScriptBlock]$Action
    )
    $window.Dispatcher.Invoke($Action, [System.Windows.Threading.DispatcherPriority]::Render)
}

# Funktion zum Aktualisieren von Fortschrittsbalken und Fenstertitel
function Update-ProgressControls {
    param (
        [int]$percent
    )
    Update-UI {
        $window.Title = "Add-User - GUI – $percent %"
        $progressBar.Value = $percent
    }
}

# Funktion zum Hinzufügen von Log-Einträgen ins RichTextBox
function AddToLog {
    param (
        [string]$message,
        [string]$category = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Update-UI {
        $run = New-Object Windows.Documents.Run
        $run.Text = "[$timestamp] [$category] $message`r`n"
        switch ($category) {
            "ERROR"   { $run.Foreground = [System.Windows.Media.Brushes]::Red }
            "SUCCESS" { $run.Foreground = [System.Windows.Media.Brushes]::Green }
            "WARN"    { $run.Foreground = [System.Windows.Media.Brushes]::Orange }
            "INFO"    { $run.Foreground = [System.Windows.Media.Brushes]::Black }
            default   { $run.Foreground = [System.Windows.Media.Brushes]::Black }
        }
        $paragraph = New-Object Windows.Documents.Paragraph($run)
        $txtLog.Document.Blocks.Add($paragraph)
        $txtLog.ScrollToEnd()
    }
}

# Funktion zum Anzeigen von Fehler-Meldungen
function Show-Error {
    param (
        [string]$message,
        [Exception]$exception
    )
    $details = if ($exception) { $exception.Message } else { "Keine weiteren Details" }
    Update-UI {
        [System.Windows.MessageBox]::Show("$message`n`nDetails: $details", "Fehler",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error)
    }
}

# Hilfsfunktion für farbige Log-Ausgabe
function Write-ColoredLog {
    param (
        [string]$Message,
        [string]$Color = "Black"
    )
    $txtLog.Dispatcher.Invoke([Action]{
        $run = New-Object Windows.Documents.Run
        $run.Text = "[$((Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))] $Message`r`n"
        $run.Foreground = [System.Windows.Media.Brushes]::$Color
        $paragraph = New-Object Windows.Documents.Paragraph($run)
        $txtLog.Document.Blocks.Add($paragraph)
        $txtLog.ScrollToEnd()
    }, [System.Windows.Threading.DispatcherPriority]::Render)
}

# Funktion zum Aktualisieren des Status-Textes
function Update-Status {
    param (
        [string]$text
    )
    Update-UI {
        $statusText.Text = $text
    }
}

# Funktion zum Filtern des Logs für relevante Einträge
function Get-FilteredLog {
    $textRange = New-Object System.Windows.Documents.TextRange($txtLog.Document.ContentStart, $txtLog.Document.ContentEnd)
    $lines = $textRange.Text -split "`r?`n" | Where-Object { $_ }
    $filtered = $lines | Where-Object {
        $_ -match "Standardgruppen" -or
        $_ -match "ExtensionAttributes" -or
        $_ -match "Gruppe" -or
        $_ -match "Login-Daten"
    } | ForEach-Object {
        if ($_ -match "^\[.*?\]\s*(.*)$") { $matches[1] } else { $_ }
    }
    return $filtered -join "`r`n"
}

#endregion

#region UI-Bindings und Initialisierung
# ------------------------------------------------#
# 3) UI-Bindings und Initialisierung
# ------------------------------------------------#

# UI-Elemente binden
$txtUser         = $window.FindName("txtUser")
$txtGivenName    = $window.FindName("txtGivenName")
$txtLastName     = $window.FindName("txtLastName")
$txtBuro         = $window.FindName("txtBuro")
$comboSite       = $window.FindName("comboSite")
$txtRufnummer    = $window.FindName("txtRufnummer")
$txtHandynummer  = $window.FindName("txtHandynummer")
$txtTitle        = $window.FindName("txtTitle")
$comboAmts       = $window.FindName("comboAmts")
$comboLauf       = $window.FindName("comboLauf")
$comboDept       = $window.FindName("comboDept")
$txtExp          = $window.FindName("txtExp")
$txtAktiv        = $window.FindName("txtAktiv")
$txtTicket       = $window.FindName("txtTicket")
$comboRolle      = $window.FindName("comboRolle")
$comboFunktion   = $window.FindName("comboFunktion")
$comboSonder     = $window.FindName("comboSonder")
$txtDesc         = $window.FindName("txtDesc")
$txtRefUser      = $window.FindName("txtRefUser")
$lstGroups       = $window.FindName("lstGroups")
$btnSearchGroups = $window.FindName("btnSearchGroups")

$chkIVBB          = $window.FindName("chkIVBB")
$chkGVPL          = $window.FindName("chkGVPL")
$chkPhonebook     = $window.FindName("chkPhonebook")
$chkVIP           = $window.FindName("chkVIP")
$chkIsFemale      = $window.FindName("chkIsFemale")
$chkAbgeordnet    = $window.FindName("chkAbgeordnet")
$chkNatPerson     = $window.FindName("chkNatPerson")
$chkConet         = $window.FindName("chkConet")
$chkExternAccount = $window.FindName("chkExternAccount")
$chkMailbox       = $window.FindName("chkMailbox")
$chkExtern        = $window.FindName("chkExtern")
$chkVerstecken    = $window.FindName("chkVerstecken")
$chkResMailbox    = $window.FindName("chkResMailbox")

$txtLog           = $window.FindName("txtLog")
$statusText       = $window.FindName("statusText")
$progressBar      = $window.FindName("progressBar")

$btnStart         = $window.FindName("btnStart")
$btnCancel        = $window.FindName("btnCancel")
$btnExit          = $window.FindName("btnExit")
$btnCopyLog       = $window.FindName("btnCopyLog")
$btnNewUser       = $window.FindName("btnNewUser")
$btnSaveProfile   = $window.FindName("btnSaveProfile")
$btnLoadProfile   = $window.FindName("btnLoadProfile")
$btnMassCreation  = $window.FindName("btnMassCreation")

$comboStdProfile  = $window.FindName("comboStdProfile")
$imgLogo          = $window.FindName("imgLogo")

$btnSearchGroups = $window.FindName("btnSearchGroups")
$btnSearchGroups_DropDown = $window.FindName("btnSearchGroups_DropDown")
$lstGroups = $window.FindName("lstGroups")
$comboDept = $window.FindName("comboDept")
$txtLog = $window.FindName("txtLog")
$comboGender = $window.FindName("comboGender")

# Steuerelement-Map fuer Profilfunktionen
$global:UI = @{
    txtUser        = $txtUser
    txtGivenName   = $txtGivenName
    txtLastName    = $txtLastName
    txtBuro        = $txtBuro
    comboSite      = $comboSite
    comboDept      = $comboDept
    txtRufnummer   = $txtRufnummer
    txtHandynummer = $txtHandynummer
    txtTitle       = $txtTitle
    comboAmts      = $comboAmts
    comboLauf      = $comboLauf
    txtExp         = $txtExp
    txtAktiv       = $txtAktiv
    txtTicket      = $txtTicket
    comboRolle     = $comboRolle
    comboFunktion  = $comboFunktion
    comboSonder    = $comboSonder
    txtDesc        = $txtDesc
    txtRefUser     = $txtRefUser
    lstGroups      = $lstGroups
    comboGender    = $comboGender
    chkNatPerson   = $chkNatPerson
    chkIVBB        = $chkIVBB
    chkGVPL        = $chkGVPL
    chkVIP         = $chkVIP
    chkExtern      = $chkExtern
    chkVerstecken  = $chkVerstecken
    chkPhonebook   = $chkPhonebook
    chkAbgeordnet  = $chkAbgeordnet
    chkConet       = $chkConet
    chkExternAccount = $chkExternAccount
    chkResMailbox  = $chkResMailbox
    chkMailbox     = $chkMailbox
}

# ToolTips setzen
$txtUser.ToolTip        = "Geben Sie die UserID ein (z. B. AbbasiW)"
$txtGivenName.ToolTip   = "Vorname des Benutzers"
$txtLastName.ToolTip    = "Nachname des Benutzers"
$txtBuro.ToolTip        = "Büronummer (RSP/STR/KTR/ wird automatisch gesetzt)"
$comboSite.ToolTip      = "Wählen Sie den Standort (RSP = Bonn, STR = Berlin, KTR = Berlin)"
$txtRufnummer.ToolTip   = "Telefonnummer im Format +49 228 99305 3992 oder +49 30 18305 3992"
$txtHandynummer.ToolTip = "Mobilnummer im Format +49 17632697932 (optional)"
$txtTitle.ToolTip       = "Titel des Benutzers (z. B. Dr., Prof.)"
$comboAmts.ToolTip      = "Amtsbezeichnung auswählen oder eingeben"
$comboLauf.ToolTip      = "Laufbahngruppe auswählen"
$comboDept.ToolTip      = "Abteilung/Department auswählen"
$txtExp.ToolTip         = "Befristung (U = unbefristet, oder Datum im Format dd.MM.yyyy)"
$txtAktiv.ToolTip       = "Aktivierungsdatum (S = sofort, oder dd.MM.yyyy)"
$txtTicket.ToolTip      = "Ticketnummer aus dem System (Pflichtfeld)"
$comboRolle.ToolTip     = "Rolle des Benutzers (optional)"
$comboFunktion.ToolTip  = "Funktion des Benutzers (optional)"
$comboSonder.ToolTip    = "Sonderkennzeichnung (optional)"
$txtDesc.ToolTip        = "Beschreibung/Kommentar zum Benutzer"
$txtRefUser.ToolTip     = "Referenz-Benutzer (z. B. für Gruppenkopie)"
$lstGroups.ToolTip      = "Zusätzliche AD-Gruppen (per Suche hinzufügen)"
$btnSearchGroups.ToolTip= "Gruppen basierend auf Department suchen"
Initialize-Logger -WpfControl $txtLog

# Globale Variable für Log-Control setzen
#$global:WpfLogControl = $txtLog
#I#nitialize-Logger -RichTextBox $txtLog


# Lade das Logo
$imgPath = "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\ToolBox\Res\BMUV_logo.jpg"
if (Test-Path $imgPath) {
    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
    $bitmap.BeginInit()
    $bitmap.UriSource = New-Object Uri($imgPath)
    $bitmap.EndInit()
    $imgLogo.Source = $bitmap
}

# Lade JSON-Daten für Dropdowns
if (Test-Path $global:AppConfig.Paths.AmtsPath) {
    $amtsData = Get-Content -Path $global:AppConfig.Paths.AmtsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $comboAmts.ItemsSource = $amtsData
}
if (Test-Path $global:AppConfig.Paths.LaufPath) {
    $laufData = Get-Content -Path $global:AppConfig.Paths.LaufPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $comboLauf.ItemsSource = $laufData
}
if (Test-Path $global:AppConfig.Paths.DeptsLongOG) {
    $departmentsLongOG = Get-Content -Path $global:AppConfig.Paths.DeptsLongOG -Raw -Encoding UTF8 | ConvertFrom-Json
    $departmentsLong = $departmentsLongOG | ForEach-Object { $_.Department }
    $global:departmentOGMapping = @{}
    foreach ($item in $departmentsLongOG) {
        $global:departmentOGMapping[$item.Department] = $item.Name
    }
    $comboDept.ItemsSource = $departmentsLong
}
if (Test-Path $global:AppConfig.Paths.Depts2Json) {
    $global:departmentMapping = Get-Content -Path $global:AppConfig.Paths.Depts2Json -Raw -Encoding UTF8 | ConvertFrom-Json
}
if (Test-Path $global:AppConfig.Paths.FktPath) {
    $fktData = Get-Content -Path $global:AppConfig.Paths.FktPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $comboFunktion.ItemsSource = $fktData
}
if (Test-Path $global:AppConfig.Paths.SonderPath) {
    $sonderData = Get-Content -Path $global:AppConfig.Paths.SonderPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $comboSonder.ItemsSource = $sonderData
}

# Erstelle Department-Mapping
$departmentMapping = @{}
if ($departmentsLong -and $departmentsShort) {
    for ($i = 0; $i -lt $departmentsLong.Count; $i++) {
        $departmentMapping[$departmentsLong[$i]] = $departmentsShort[$i]
    }
}

#endregion

#region UI-Validierung und Feedback
# ------------------------------------------------#
# 4) UI-Validierung und Feedback
# ------------------------------------------------#

# UserID-Validierung mit farblicher Hervorhebung
$txtUser.Add_TextChanged({
    $text = $txtUser.Text.Trim()
    Update-UI {
        if ($text.Length -gt 20) {
            $txtUser.Background = [System.Windows.Media.Brushes]::LightPink
            $txtUser.ToolTip = "Fehler: UserID darf maximal 20 Zeichen lang sein"
        } elseif ($text.Length -lt 2 -or $text -notmatch "^[A-Za-z0-9]+$") {
            $txtUser.Background = "#FFFFE0E0"  # Leichtes Rot für ungültige Eingabe
        } else {
            $txtUser.Background = [System.Windows.Media.Brushes]::White
            $txtUser.ToolTip = "Eindeutige ID für den Benutzer (z. B. mmuster), maximal 20 Zeichen"
        }
    }
})

# Ticketnummer-Validierung
$txtTicket.Add_TextChanged({
    Update-UI {
        if ($txtTicket.Text -notmatch "^[\d-]+$") {
        $txtTicket.Background = "#FFFFE0E0"
    } else {
        $txtTicket.Background = [System.Windows.Media.Brushes]::White
    }
    }
})

# Telefonnummer-Validierung
$txtRufnummer.Add_TextChanged({
    Update-UI {
        if ($txtRufnummer.Text -and $txtRufnummer.Text -notmatch "^\+49\s(228|30)\s\d{5}\s\d{4}$") {
            $txtRufnummer.Background = "#FFFFE0E0"  # Leichtes Rot für ungültiges Format
            $txtRufnummer.ToolTip = "Format: +49 228 99305 3992 oder +49 30 18305 3992"
        } else {
            $txtRufnummer.Background = [System.Windows.Media.Brushes]::White
            $txtRufnummer.ToolTip = "Telefonnummer im Format +49 228 99305 3992 oder +49 30 18305 3992"
        }
    }
})

# Mobilnummer-Validierung
$txtHandynummer.Add_TextChanged({
    Update-UI {
        if ($txtHandynummer.Text -and $txtHandynummer.Text -notmatch "^\+49\s\d{9,11}$") {
            $txtHandynummer.Background = "#FFFFE0E0"  # Leichtes Rot für ungültiges Format
            $txtHandynummer.ToolTip = "Format: +49 17632697932"
        } else {
            $txtHandynummer.Background = [System.Windows.Media.Brushes]::White
            $txtHandynummer.ToolTip = "Mobilnummer im Format +49 17632697932 (optional)"
        }
    }
})

# Ablaufdatum-Validierung
$txtExp.Add_TextChanged({
    Update-UI {
        if ($txtExp.Text -and $txtExp.Text -ne "U" -and $txtExp.Text -notmatch "^\d{2}\.\d{2}\.\d{4}$") {
            $txtExp.Background = "#FFFFE0E0"  # Leichtes Rot für ungültiges Datum
        } else {
            $txtExp.Background = [System.Windows.Media.Brushes]::White
        }
    }
})

# Aktivierungsdatum-Validierung
$txtAktiv.Add_TextChanged({
    Update-UI {
        if ($txtAktiv.Text -and $txtAktiv.Text -ne "S" -and $txtAktiv.Text -notmatch "^\d{2}\.\d{2}\.\d{4}$") {
            $txtAktiv.Background = "#FFFFE0E0"  # Leichtes Rot für ungültiges Datum
        } else {
            $txtAktiv.Background = [System.Windows.Media.Brushes]::White
        }
    }
})

# Keine Standardwerte setzen (auf Wunsch leer lassen)
Update-UI {
    $txtAktiv.Text = ""
    $txtExp.Text = ""
    $txtTicket.Text = ""
    $comboSite.SelectedItem = $null  # Kein Default-Standort
}

#endregion

#region Standardprofile
# ------------------------------------------------#
# 5) Standardprofile
# ------------------------------------------------#

# Event-Handler für Standardprofil-Auswahl
$comboStdProfile.Add_SelectionChanged({
    if ($comboStdProfile.SelectedItem -and $comboStdProfile.SelectedItem.Content) {
        $pKey = $comboStdProfile.SelectedItem.Content
        if ($global:StdProfiles.ContainsKey($pKey)) {
            $defs = $global:StdProfiles[$pKey]
            Update-UI {
                $comboSite.SelectedItem = $comboSite.Items | Where-Object { $_.Content -eq $defs.Site }
                $comboAmts.Text       = $defs.amtsbez
                $comboLauf.Text       = $defs.laufgruppe
                $comboRolle.SelectedItem = $comboRolle.Items | Where-Object { $_.Content -eq $defs.roleSelection }
                $txtExp.Text          = $defs.ExpDate
                $txtDesc.Text         = $defs.desc
                $comboDept.Text       = $defs.Department
                $txtTicket.Text       = $defs.TicketNr
                $txtAktiv.Text        = $defs.EntryDate
                $comboSonder.Text     = $defs.sonderkenn
                $comboFunktion.Text   = $defs.funktion
                $txtRefUser.Text      = $defs.refUser
                $chkIVBB.IsChecked    = ($defs.isIVBB -eq "j")
                $chkGVPL.IsChecked    = ($defs.isGVPL -eq "j")
                $chkVIP.IsChecked     = ($defs.isVIP -eq "j")
                $gender.IsChecked = $defs.gender
                $chkExtern.IsChecked  = $defs.isExtern
                $chkVerstecken.IsChecked = $defs.isVerstecken
                $chkPhonebook.IsChecked  = $defs.isPhonebook
                $chkNatPerson.IsChecked  = ($defs.isNatPerson -eq "j")
                $chkResMailbox.IsChecked = ($defs.isResMailbox -eq "j")
                $chkAbgeordnet.IsChecked = ($defs.isAbgeordnet -eq "j")
                $chkExternAccount.IsChecked = ($defs.isExternAccount -eq "j")
                $chkMailbox.IsChecked  = ($defs.makeMailbox -eq "j")
                $txtTitle.Text        = $defs.titleValue
                $txtBuro.Text         = $defs.Buro
                $txtRufnummer.Text    = $defs.Rufnummer
                $txtHandynummer.Text  = $defs.Handynummer
            }
            AddToLog "Standardprofil '$pKey' geladen" "INFO"
        }
    }
})

#endregion

#region Benutzererstellung
# ------------------------------------------------#
# 6) Benutzererstellung
# ------------------------------------------------#

# Funktion zur Validierung der Benutzereingaben
function Validate-UserInput {
    param (
        [string]$UserID,
        [string]$Site
    )
    $isValid = $true

    if ([string]::IsNullOrWhiteSpace($UserID)) {
        AddToLog "UserID darf nicht leer sein" "ERROR"
        $txtUser.Background = [System.Windows.Media.Brushes]::LightPink
        $isValid = $false
    } elseif ($UserID.Length -gt 20) {
        AddToLog "UserID darf maximal 20 Zeichen lang sein" "ERROR"
        $txtUser.Background = [System.Windows.Media.Brushes]::LightPink
        $isValid = $false
    }

    if ([string]::IsNullOrWhiteSpace($Site)) {
        AddToLog "Site muss ausgewählt werden" "ERROR"
        $comboSite.Background = [System.Windows.Media.Brushes]::LightPink
        $isValid = $false
    }

    if ([string]::IsNullOrWhiteSpace($TicketNr)) {
        AddToLog "Ticketnummer darf nicht leer sein" "ERROR"
        $txtTicket.Background = [System.Windows.Media.Brushes]::LightPink
        $isValid = $false
    }

    if ($isValid) {
        Update-Status "Eingaben validiert"
    } else {
        Update-Status "Validierungsfehler - überprüfe die Eingaben"
    }
    return $isValid
}

# Funktion für asynchrone Benutzererstellung (für Massenerstellung beibehalten)
function Start-UserCreationRunspace {
    param (
        [hashtable]$jobParams
    )
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("global:WpfLogControl", $txtLog)

    $psInstance = [PowerShell]::Create()
    $psInstance.Runspace = $runspace

    $scriptBlock = {
        param($params)
        $modules = @(
            "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v22\ADUserCreationModule.psm1",
            "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v22\UserCreationLogic2.psm1",
            "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v22\AttributeModule.psm1",
            "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v22\LogModule.psm1",
            "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v22\DatabaseModule.psm1",
            "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v22\FolderStructureModule.psm1",
            "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v22\MailboxModule.psm1"
        )
                    # Initialize logger in runspace


Import-Module "$params.ToolRoot\LogModule.psm1" -Force
Initialize-Logger -InJob
        foreach ($module in $modules) {
            Import-Module $module -Force -ErrorAction SilentlyContinue
        }

        $progressCallback = $params.ProgressCallback
        $params.ProgressCallback = $progressCallback

        try {
            ProcessUserCreation @params
        } catch {
            $errorMsg = "Fehler in ProcessUserCreation: $($_.Exception.Message)"
            Update-UI { AddToLog $errorMsg "ERROR" }
            throw $errorMsg
        }
    }

    $psInstance.AddScript($scriptBlock).AddArgument($jobParams) | Out-Null
    $asyncResult = $psInstance.BeginInvoke()

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(500)
    $timer.Add_Tick({
        if ($asyncResult.IsCompleted) {
            try {
                $psInstance.EndInvoke($asyncResult)
                Update-UI {
                    AddToLog "Benutzererstellung abgeschlossen" "SUCCESS"
                    Update-Status "Fertig"
                    $btnStart.IsEnabled = $true
                    Update-ProgressControls -percent 100
                }
            } catch {
                Update-UI {
                    AddToLog "Fehler bei der Ausführung: $($_.Exception.Message)" "ERROR"
                    Update-Status "Fehler"
                    $btnStart.IsEnabled = $true
                    Update-ProgressControls -percent 0
                }
            } finally {
                $timer.Stop()
                $psInstance.Dispose()
                $runspace.Close()
                $runspace.Dispose()
            }
        }
    })
    $timer.Start()
}

# Event-Handler für Einzelbenutzer-Erstellung (synchron mit live Updates)
$btnStart.Add_Click({

    # 1. Daten sammeln
    $userID      = $txtUser.Text.Trim()
    $givenName   = $txtGivenName.Text.Trim()
    $lastName    = $txtLastName.Text.Trim()
    $buro        = $txtBuro.Text.Trim()
    $site        = if ($comboSite.SelectedItem) { $comboSite.SelectedItem.Content } else { "" }
    $rufnummer   = $txtRufnummer.Text.Trim()
    $handynummer = $txtHandynummer.Text.Trim()
    $titleValue  = $txtTitle.Text.Trim()
    $amtsbez     = $comboAmts.Text.Trim()
    $laufgruppe  = $comboLauf.Text.Trim()
    $department  = $comboDept.Text.Trim()
    $expDate     = $txtExp.Text.Trim()
    $entryDate   = $txtAktiv.Text.Trim()
    $ticketNr    = $txtTicket.Text.Trim()
    $roleSelection = if ($comboRolle.SelectedItem) { $comboRolle.SelectedItem.Content } else { "" }
    $sonderkenn  = $comboSonder.Text.Trim()
    $funktion    = $comboFunktion.Text.Trim()
    $desc        = $txtDesc.Text.Trim()
    $refUser     = $txtRefUser.Text.Trim()
    # Annahme: $lstGroups.Items enthält Objekte mit einer Property "Name"
    $groups      = $lstGroups.Items | ForEach-Object { $_.Name }
    $gender      = if ($comboGender.SelectedItem) { $comboGender.SelectedItem.Content } else { "Mann" }


    # 2. Preview aufrufen
    $previewOK = Show-Preview -UserID $userID `
                              -Gender $gender`
                              -givenName $givenName `
                              -lastName $lastName `
                              -Buro $buro `
                              -Site $site `
                              -Rufnummer $rufnummer `
                              -Handynummer $handynummer `
                              -titleValue $titleValue `
                              -amtsbez $amtsbez `
                              -laufgruppe $laufgruppe `
                              -Department $department `
                              -ExpDate $expDate `
                              -EntryDate $entryDate `
                              -TicketNr $ticketNr `
                              -roleSelection $roleSelection `
                              -sonderkenn $sonderkenn `
                              -funktion $funktion `
                              -desc $desc `
                              -refUser $refUser `
                              -Groups $groups
    if (-not $previewOK) {
        # Benutzer hat die Vorschau abgebrochen, Prozess beenden
        Write-ColoredLog "Benutzererstellung abgebrochen (Preview nicht bestätigt)." "WARN"
        return
    }

    # 3. Falls Preview bestätigt, starten Sie den Erstellungsprozess:
    $jobParams = @{
        UserID = $txtUser.Text.Trim()
        givenName = $txtGivenName.Text.Trim()
        lastName = $txtLastName.Text.Trim()
        Buro = $txtBuro.Text.Trim()
        Site = if ($comboSite.SelectedItem) { $comboSite.SelectedItem.Content } else { "" }
        Rufnummer = $txtRufnummer.Text.Trim()
        Handynummer = $txtHandynummer.Text.Trim()
        titleValue = $txtTitle.Text.Trim()
        amtsbez = $comboAmts.Text.Trim()
        laufgruppe = $comboLauf.Text.Trim()
        Department = $comboDept.Text.Trim()
        ExpDate = $txtExp.Text.Trim()
        EntryDate = $txtAktiv.Text.Trim()
        TicketNr = $txtTicket.Text.Trim()
        roleSelection = if ($comboRolle.SelectedItem) { $comboRolle.SelectedItem.Content } else { "" }
        funktion = $comboFunktion.Text.Trim()
        sonderkenn = $comboSonder.Text.Trim()
        desc = $txtDesc.Text.Trim()
        refUser = $txtRefUser.Text.Trim()
        isIVBB = if ($chkIVBB.IsChecked) { "j" } else { "n" }
        isGVPL = if ($chkGVPL.IsChecked) { "j" } else { "n" }
        isPhonebook = $chkPhonebook.IsChecked
        isVIP = if ($chkVIP.IsChecked) { "j" } else { "n" }
        gender = if ($comboGender.SelectedItem) { $comboGender.SelectedItem.Content } else { "Mann" }
        #isFemale = if ($gender -eq "Frau") { "j" } else { "n" }
        isAbgeordnet = if ($chkAbgeordnet.IsChecked) { "j" } else { "n" }
        isNatPerson = if ($chkNatPerson.IsChecked) { "j" } else { "n" }
        isConet = if ($chkConet.IsChecked) { "j" } else { "n" }
        isExternAccount = if ($chkExternAccount.IsChecked) { "j" } else { "n" }
        makeMailbox = if ($chkMailbox.IsChecked) { "j" } else { "n" }
        isExtern = $chkExtern.IsChecked
        isVerstecken = $chkVerstecken.IsChecked
        isResMailbox = if ($chkResMailbox.IsChecked) { "j" } else { "n" }
        departmentOGMapping = $global:departmentOGMapping
        AdditionalGroups = ($lstGroups.Items | Where-Object { $_.IsChecked } | ForEach-Object { $_.Name })
        LogTextbox = $txtLog
        ProgressCallback = { param($p) Update-UI { Update-ProgressControls -percent $p } }
        InJob = $false
    }

    Update-UI {
        $txtLog.Document.Blocks.Clear()
        Update-ProgressControls -percent 0
        $btnStart.IsEnabled = $false
        $statusText.Text = "Benutzer wird erstellt..."
    }
    Write-ColoredLog "Starte Benutzererstellung für $($jobParams.UserID)" "Black"

    try {
        ProcessUserCreation @jobParams
        Update-UI {
            Write-ColoredLog "Benutzererstellung abgeschlossen" "Green"
            $statusText.Text = "Fertig"
            $btnStart.IsEnabled = $true
            $btnCopyLog.IsEnabled = $true
            Update-ProgressControls -percent 100
        }
    } catch {
        Update-UI {
            Write-ColoredLog "Fehler bei der Benutzererstellung: $($_.Exception.Message)" "Red"
            $statusText.Text = "Fehler"
            $btnStart.IsEnabled = $true
            Update-ProgressControls -percent 0
        }
    }
})

# Funktion zur Anzeige der Vorschau für Einzelbenutzer
function Show-Preview {
    param (
        [string]$UserID, [string]$givenName, [string]$gender,  [string]$lastName, [string]$Buro, [string]$Site,
        [string]$Rufnummer, [string]$Handynummer, [string]$titleValue, [string]$amtsbez,
        [string]$laufgruppe, [string]$Department, [string]$ExpDate, [string]$EntryDate,
        [string]$TicketNr, [string]$roleSelection, [string]$sonderkenn, [string]$funktion,
        [string]$desc, [string]$refUser, [array]$Groups
    )
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Benutzer-Vorschau"
        Height="550" Width="600"
        Background="#E0ECF8"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto">
            <StackPanel x:Name="previewPanel">
                <TextBlock Text="Vorschau des neuen Benutzers" 
                           FontSize="18" FontWeight="Bold" 
                           Foreground="#002B5E" 
                           Margin="0,0,0,10"/>
            </StackPanel>
        </ScrollViewer>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnOK" Content="Übernehmen" Width="100" Margin="0,0,10,0"/>
            <Button x:Name="btnCancel" Content="Abbrechen" Width="100"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader($xaml)
    $win = [Windows.Markup.XamlReader]::Load($reader)
    if (-not $win) {
        AddToLog "FEHLER: Konnte PreviewWindow.xaml nicht laden" "ERROR"
        return $false
    }

    $previewPanel = $win.FindName("previewPanel")
    $btnOK = $win.FindName("btnOK")
    $btnCancel = $win.FindName("btnCancel")

    $fields = [PSCustomObject]@{
        "UserID" = $UserID
        "Geschlecht" = "$gender"
        "Vorname" = $givenName
        "Nachname" = $lastName
        "Büro" = $Buro
        "Standort" = $Site
        "Telefonnummer" = $Rufnummer
        "Mobilnummer" = $Handynummer
        "Titel" = $titleValue
        "Amtsbezeichnung" = $amtsbez
        "Laufbahngruppe" = $laufgruppe
        "Abteilung" = $Department
        "Ablaufdatum" = $ExpDate
        "Aktivierungsdatum" = $EntryDate
        "Ticketnummer" = $TicketNr
        "Rolle" = $roleSelection
        "Sonderkennzeichen" = $sonderkenn
        "Funktion" = $funktion
        "Beschreibung" = $desc
        "Referenzbenutzer" = $refUser
        "Gruppen" = ($Groups -join ", ")
    }

    foreach ($prop in $fields.PSObject.Properties) {
        $stack = New-Object Windows.Controls.StackPanel
        $stack.Orientation = "Horizontal"
        $label = New-Object Windows.Controls.TextBlock
        $label.Text = "$($prop.Name): "
        $label.FontWeight = "Bold"
        $value = New-Object Windows.Controls.TextBlock
        $value.Text = $prop.Value
        $stack.Children.Add($label)
        $stack.Children.Add($value)
        $previewPanel.Children.Add($stack)
    }

    # Setze den DialogResult anhand der Schaltflächen
    $btnOK.Add_Click({
        $win.DialogResult = $true
        $win.Close()
    })
    $btnCancel.Add_Click({
        $win.DialogResult = $false
        $win.Close()
    })

    $result = $win.ShowDialog()
    return $result
}

#endregion

#region Massenerstellung
# ------------------------------------------------#
# 7) Massenerstellung
# ------------------------------------------------#

# Event-Handler für Massenerstellung
$btnMassCreation.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "JSON files (*.json)|*.json"
    $dlg.Title  = "Massen-Erstellung: JSON auswählen"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $json = Get-Content $dlg.FileName -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not $json) {
            [System.Windows.MessageBox]::Show("Leere oder ungültige JSON.")
            return
        }
        Show-MassCreationPreview -UserList $json
    }
})

# Funktion zur Anzeige der Vorschau für Massenerstellung
function Show-MassCreationPreview {
    param (
        [System.Collections.IEnumerable]$UserList
    )
    [xml]$xamlMass = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Massen-Erstellung Vorschau"
        Height="500" Width="800"
        Background="#E0ECF8"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <DataGrid x:Name="dgUsers" Grid.Row="0" Margin="10"
                  AutoGenerateColumns="False" IsReadOnly="True">
            <DataGrid.Columns>
                <DataGridTextColumn Header="UserID" Binding="{Binding UserID}" Width="*"/>
                <DataGridTextColumn Header="GivenName" Binding="{Binding GivenName}" Width="*"/>
                <DataGridTextColumn Header="LastName" Binding="{Binding LastName}" Width="*"/>
                <DataGridTextColumn Header="Site" Binding="{Binding Site}" Width="*"/>
                <DataGridTextColumn Header="Department" Binding="{Binding Department}" Width="*"/>
                <DataGridTextColumn Header="TicketNr" Binding="{Binding TicketNr}" Width="*"/>
            </DataGrid.Columns>
        </DataGrid>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center" Margin="10">
            <Button x:Name="btnOk" Content="OK" Width="100" Margin="10,0"/>
            <Button x:Name="btnCancel" Content="Abbrechen" Width="100"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader($xamlMass)
    $winMass = [Windows.Markup.XamlReader]::Load($reader)

    $dg        = $winMass.FindName("dgUsers")
    $btnOk     = $winMass.FindName("btnOk")
    $btnCancel = $winMass.FindName("btnCancel")

    $dg.ItemsSource = $UserList

    $btnOk.Add_Click({
        $winMass.Close()
        Start-MassCreation -UserList $UserList
    })
    $btnCancel.Add_Click({ $winMass.Close() })

    $winMass.Owner = $window
    $winMass.ShowDialog() | Out-Null
}

# Funktion zur Durchführung der Massenerstellung
function Start-MassCreation {
    param (
        [System.Collections.IEnumerable]$UserList
    )
    $jobs = @()
    Update-Status "Massen-Erstellung wird vorbereitet..."
    foreach ($u in $UserList) {
        $ok = Show-ModernUserPreview -UserObj $u
        if (-not $ok) {
            AddToLog "Massen-Erstellung für $($u.UserID) abgebrochen" "WARN"
            continue
        }

        $jobParams = @{
            UserID           = $u.UserID
            gender           = $u.gender
            givenName        = $u.GivenName
            lastName         = $u.LastName
            Buro             = $u.Buro
            Rufnummer        = $u.Rufnummer
            Handynummer      = $u.Handynummer
            titleValue       = $u.titleValue
            amtsbez          = $u.amtsbez
            laufgruppe       = $u.laufgruppe
            roleSelection    = $u.roleSelection
            Site             = $u.Site
            ExpDate          = $u.ExpDate
            desc             = $u.desc
            Department       = $u.Department
            TicketNr         = $u.TicketNr
            EntryDate        = $u.EntryDate
            sonderkenn       = $u.sonderkenn
            funktion         = $u.funktion
            refUser          = $u.refUser
            isIVBB           = $u.isIVBB
            isGVPL           = $u.isGVPL
            isVIP            = $u.isVIP
            #isFemale         = $u.isFemale
            isExtern         = $u.isExtern
            isVerstecken     = $u.isVerstecken
            isPhonebook      = $u.isPhonebook
            #isNatPerson      = $u.isNatPerson
            isResMailbox     = $u.isResMailbox
            isAbgeordnet     = $u.isAbgeordnet
            isExternAccount  = $u.isExternAccount
            makeMailbox      = $u.makeMailbox
            departmentOGMapping = $global:departmentOGMapping
            departmentMapping   = $global:departmentMapping
            AdditionalGroups = $u.AdditionalGroups
            InJob            = $true
            LogTextbox       = $txtLog
            ProgressCallback = { param($p) Update-UI { Update-ProgressControls -percent $p } }
        }

        Start-UserCreationRunspace -jobParams $jobParams
        $jobs += $jobParams.UserID
    }

    Update-ProgressControls -percent 0
    Update-Status "Massen-Erstellung läuft..."
    $completed = 0
    $total = $jobs.Count
    while ($completed -lt $total) {
        $completed = ($jobs | ForEach-Object { 
            $user = $_
            $logText = (New-Object System.Windows.Documents.TextRange($txtLog.Document.ContentStart, $txtLog.Document.ContentEnd)).Text
            if ($logText -match "Benutzererstellung abgeschlossen.*$user" -or $logText -match "Fehler bei der Ausführung.*$user") { 1 } else { 0 }
        } | Measure-Object -Sum).Sum
        $percent = if ($total -gt 0) { [math]::Round(($completed / $total) * 100) } else { 0 }
        Update-ProgressControls -percent $percent
        Start-Sleep -Milliseconds 500
    }
    Update-ProgressControls -percent 100

    AddToLog "Massen-Erstellung abgeschlossen" "SUCCESS"
    Update-Status "Massen-Erstellung abgeschlossen"
    Show-MassCreationDone
}

# Funktion zur Anzeige der Vorschau eines einzelnen Benutzers bei Massenerstellung
function Show-ModernUserPreview {
    param (
        [PSObject]$UserObj
    )
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Benutzer-Vorschau (Massenerstellung)"
        Height="550" Width="600"
        Background="#E0ECF8"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto">
            <StackPanel x:Name="massPreviewPanel">
                <TextBlock Text="Vorschau des neuen Benutzers" 
                           FontSize="18" FontWeight="Bold" 
                           Foreground="#002B5E" 
                           Margin="0,0,0,10"/>
            </StackPanel>
        </ScrollViewer>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnMassOK" Content="Übernehmen" Width="100" Margin="0,0,10,0"/>
            <Button x:Name="btnMassCancel" Content="Abbrechen" Width="100"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader($xaml)
    $win = [Windows.Markup.XamlReader]::Load($reader)
    if (-not $win) {
        AddToLog "FEHLER: Konnte MassPreviewWindow.xaml nicht laden" "ERROR"
        return $false
    }

    $massPreviewPanel = $win.FindName("massPreviewPanel")
    $btnMassOK = $win.FindName("btnMassOK")
    $btnMassCancel = $win.FindName("btnMassCancel")

    foreach ($prop in $UserObj.PSObject.Properties) {
        $stack = New-Object Windows.Controls.StackPanel
        $stack.Orientation = "Horizontal"
        $label = New-Object Windows.Controls.TextBlock
        $label.Text = "$($prop.Name): "
        $label.FontWeight = "Bold"
        $value = New-Object Windows.Controls.TextBlock
        $value.Text = $prop.Value
        $stack.Children.Add($label)
        $stack.Children.Add($value)
        $massPreviewPanel.Children.Add($stack)
    }

    $result = $false
    $btnMassOK.Add_Click({ $result = $true; $win.Close() })
    $btnMassCancel.Add_Click({ $result = $false; $win.Close() })

    $win.ShowDialog() | Out-Null
    return $result
}

# Funktion zur Anzeige des Abschlussfensters nach Massenerstellung
function Show-MassCreationDone {
    [xml]$xamlDone = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Fertig"
        Height="200" Width="400"
        Background="#E0ECF8"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Massen-Erstellung abgeschlossen."
                   VerticalAlignment="Center"
                   HorizontalAlignment="Center"
                   FontWeight="Bold"
                   FontSize="16"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,10,0,10">
            <Button x:Name="btnCopyLogFinal" Content="Log kopieren" Width="100" Margin="10,0"/>
            <Button x:Name="btnCloseFinal"   Content="Schließen" Width="100"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader($xamlDone)
    $winDone = [Windows.Markup.XamlReader]::Load($reader)
    $btnCopyLogFinal = $winDone.FindName("btnCopyLogFinal")
    $btnCloseFinal   = $winDone.FindName("btnCloseFinal")

    $btnCopyLogFinal.Add_Click({
        $filteredLog = Get-FilteredLog
        [System.Windows.Clipboard]::SetText($filteredLog)
        [System.Windows.MessageBox]::Show("Gefiltertes Log in Zwischenablage kopiert.")
    })
    $btnCloseFinal.Add_Click({ $winDone.Close() })

    $winDone.Owner = $window
    $winDone.ShowDialog() | Out-Null
}

#endregion

#region Profil-Verwaltung
# ------------------------------------------------#
# 8) Profil-Verwaltung
# ------------------------------------------------#

# Funktion zum Speichern eines Profils
$btnSaveProfile.Add_Click({
    $base = "\\office.dir\files\Benutzer\$($env:USERNAME.TrimStart('0','1'))\UserData"
    Save-UserProfile -UI $global:UI -BasePath $base
})

# Funktion zum Laden eines Profils
$btnLoadProfile.Add_Click({
    $base = "\\office.dir\files\Benutzer\$($env:USERNAME.TrimStart('0','1'))\UserData"
    Load-UserProfile -UI $global:UI -BasePath $base
})

# Funktion zur Anzeige der Profil-Vorschau
function Show-ProfilePreview {
    param (
        [PSObject]$ProfileData
    )
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Profilvorschau"
        Height="550" Width="600"
        Background="#E0ECF8"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto">
            <StackPanel x:Name="MainStackPanel">
                <TextBlock Text="Vorschau des geladenen Profils" 
                           FontSize="18" FontWeight="Bold" 
                           Foreground="#002B5E" 
                           Margin="0,0,0,10"/>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="UserID: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtUserID"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                      <TextBlock Text="Geschlecht: " FontWeight="Bold"/>
                      <TextBlock x:Name="txtGender"/>
                    </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Vorname: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtGivenName"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Nachname: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtLastName"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Büro: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtBuro"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Site: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtSite"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Rufnummer: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtRufnummer"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Handynummer: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtHandynummer"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Titel: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtTitle"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Amtsbezeichnung: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtAmtsbez"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Laufbahngruppe: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtLaufgruppe"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Department: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtDepartment"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Befristet bis: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtExpDate"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Aktivierungsdatum: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtEntryDate"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Ticketnummer: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtTicketNr"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Rolle: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtRolle"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Sonderkennzeichnung: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtSonderkenn"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Funktion: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtFunktion"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Beschreibung: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtDescription"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Referenz-Benutzer: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtRefUser"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Zusätzliche Gruppen: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtAdditionalGroups"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Feste Gruppen: " FontWeight="Bold"/>
                    <TextBlock x:Name="txtFixedGroups"/>
                </StackPanel>
            </StackPanel>
        </ScrollViewer>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,10,0,0">
            <Button x:Name="btnOK" Content="Profil übernehmen" Width="130" Margin="10,0"/>
            <Button x:Name="btnCancel" Content="Abbrechen" Width="100"/>
        </StackPanel>
    </Grid>
</Window>
"@
    try {
        $reader = New-Object System.Xml.XmlNodeReader($xaml)
        $win = [Windows.Markup.XamlReader]::Load($reader)

        $btnOK = $win.FindName("btnOK")
        $btnCancel = $win.FindName("btnCancel")

        $win.FindName("txtUserID").Text = $ProfileData.UserID
        $win.FindName("txtGivenName").Text = $ProfileData.GivenName
        $win.FindName("txtLastName").Text = $ProfileData.LastName
        $win.FindName("txtBuro").Text = $ProfileData.Buro
        $win.FindName("txtSite").Text = $ProfileData.Site
        $win.FindName("txtRufnummer").Text = $ProfileData.Rufnummer
        $win.FindName("txtHandynummer").Text = $ProfileData.Handynummer
        $win.FindName("txtTitle").Text = $ProfileData.Title
        $win.FindName("txtAmtsbez").Text = $ProfileData.Amtsbez
        $win.FindName("txtLaufgruppe").Text = $ProfileData.Laufgruppe
        $win.FindName("txtDepartment").Text = $ProfileData.Department
        $win.FindName("txtExpDate").Text = $ProfileData.ExpDate
        $win.FindName("txtEntryDate").Text = $ProfileData.EntryDate
        $win.FindName("txtTicketNr").Text = $ProfileData.TicketNr
        $win.FindName("txtRolle").Text = $ProfileData.Rolle
        $win.FindName("txtSonderkenn").Text = $ProfileData.Sonderkenn
        $win.FindName("txtFunktion").Text = $ProfileData.Funktion
        $win.FindName("txtDescription").Text = $ProfileData.Description
        $win.FindName("txtRefUser").Text = $ProfileData.RefUser
        $win.FindName("txtAdditionalGroups").Text = ($ProfileData.AdditionalGroups -join ", ")
        $win.FindName("txtFixedGroups").Text = ($ProfileData.FixedAdditionalGroups -join ", ")
        $win.FindName("txtGender").Text = $ProfileData.Gender

        $script:profileAccepted = $false
        $btnOK.Add_Click({ $script:profileAccepted = $true; $win.Close() })
        $btnCancel.Add_Click({ $script:profileAccepted = $false; $win.Close() })

        $win.Owner = $window
        $win.ShowDialog() | Out-Null
        return $script:profileAccepted
    } catch {
        Write-Error "Fehler beim Laden der Profil-Vorschau: $_"
        return $false
    }
}

#endregion

#region Gruppenverwaltung
# ------------------------------------------------#
# 9) Gruppenverwaltung
# ------------------------------------------------#

# Event-Handler für Gruppensuche (Haupt-Button)
$btnSearchGroups.Add_Click({
    $deptFilter = $comboDept.Text.Trim()
    if (-not $deptFilter) {
        Write-ColoredLog "Bitte Department eingeben!" "Red"
        [System.Windows.MessageBox]::Show("Bitte Department eingeben!")
        return
    }

    try {
        # ArrayList zum Sammeln der Teil-Filter
        $filters = New-Object System.Collections.ArrayList

        # 1) IMMER: Originalstring
        [void] $filters.Add("(cn=*$deptFilter*)")

        # 2) Bindestrich-Variante (z.B. "Z II 5" -> "Z-II-5")
        $deptHyphen = $deptFilter -replace "\s","-"
        if ($deptHyphen -ne $deptFilter -and $deptHyphen) {
            [void] $filters.Add("(cn=*$deptHyphen*)")
        }

        # 3) „Römisch zu Ziffern“-Kurzform (z.B. "Z II 5" -> "ZII5" -> "Z25")
        $deptShort = $deptFilter -replace "\s+",""
        # Von großen „römischen“ Werten nach klein ersetzen:
        $deptShort = $deptShort -replace "(?i)XIII","13"
        $deptShort = $deptShort -replace "(?i)XII","12"
        $deptShort = $deptShort -replace "(?i)XI","11"
        $deptShort = $deptShort -replace "(?i)X","10"
        $deptShort = $deptShort -replace "(?i)IX","9"
        $deptShort = $deptShort -replace "(?i)VIII","8"
        $deptShort = $deptShort -replace "(?i)VII","7"
        $deptShort = $deptShort -replace "(?i)VI","6"
        $deptShort = $deptShort -replace "(?i)V","5"
        $deptShort = $deptShort -replace "(?i)IV","4"
        $deptShort = $deptShort -replace "(?i)III","3"
        $deptShort = $deptShort -replace "(?i)II","2"
        $deptShort = $deptShort -replace "(?i)I","1"

        # Nur hinzufügen, wenn sich die Kurzform tatsächlich von Original unterscheidet und nicht leer ist
        if ($deptShort -and $deptShort -ne $deptFilter) {
            [void] $filters.Add("(cn=*$deptShort*)")
        }

        # OR-Filter zusammenbauen
        # Wenn mehrere Filter existieren, baue (|(cn=*A*)(cn=*B*)...),
        # sonst nimm einfach den einen Filter
        $ldapFilter = if ($filters.Count -gt 1) {
            "(|" + ($filters -join "") + ")"
        } else {
            $filters[0]
        }

        Write-ColoredLog "Starte Gruppensuche: $ldapFilter" "Black"
        $groups = Get-ADGroup -LDAPFilter $ldapFilter -ErrorAction Stop
        $sortedGroups = $groups | Sort-Object Name

        # UI aktualisieren
        Update-UI {
            $existingGroups = $lstGroups.Items | ForEach-Object { $_.Name }

            foreach ($grpItem in $sortedGroups) {
                if ($existingGroups -notcontains $grpItem.Name) {
                    # Falls du die C#-Klasse 'GroupItem' verwendest:
                    $groupItem = New-Object GroupItem
                    $groupItem.Name = $grpItem.Name
                    $groupItem.IsChecked = $false
                    $lstGroups.Items.Add($groupItem)
                }
            }

            if ($lstGroups.Items.Count -eq 0) {
                Write-ColoredLog "Keine Gruppen gefunden für '$deptFilter'" "Orange"
                [System.Windows.MessageBox]::Show("Keine Gruppen gefunden.")
            }
            else {
                Write-ColoredLog "Gruppensuche erfolgreich: $($lstGroups.Items.Count) Gruppen gefunden" "Green"
            }
        }
    }
    catch {
        Write-ColoredLog "Fehler bei der Gruppensuche: $($_.Exception.Message)" "Red"
        [System.Windows.MessageBox]::Show("Fehler bei der Gruppensuche: $($_.Exception.Message)")
    }
})



# Event-Handler für erweiterte Gruppensuche (Dropdown-Button)
if ($btnSearchGroups_DropDown) {
    $btnSearchGroups_DropDown.Add_Click({
        [xml]$xamlSearch = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Erweiterte Gruppensuche"
        Height="200" Width="400"
        Background="#E0ECF8"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Label Grid.Row="0" Content="Suchbegriff (CN oder Beschreibung):"/>
        <TextBox x:Name="txtSearchTerm" Grid.Row="1" Width="350" ToolTip="z. B. 'Verteiler' oder 'IT'"/>
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,10,0,0">
            <CheckBox x:Name="chkSearchDesc" Content="In Beschreibung suchen" Margin="0,0,10,0"/>
            <CheckBox x:Name="chkExactMatch" Content="Exakte Übereinstimmung"/>
        </StackPanel>
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnSearch" Content="Suchen" Width="100" Margin="0,0,10,0"/>
            <Button x:Name="btnCancel" Content="Abbrechen" Width="100"/>
        </StackPanel>
    </Grid>
</Window>
"@
        $reader = New-Object System.Xml.XmlNodeReader $xamlSearch
        $searchWin = [Windows.Markup.XamlReader]::Load($reader)
        
        $txtSearchTerm = $searchWin.FindName("txtSearchTerm")
        $chkSearchDesc = $searchWin.FindName("chkSearchDesc")
        $chkExactMatch = $searchWin.FindName("chkExactMatch")
        $btnSearch = $searchWin.FindName("btnSearch")
        $btnCancel = $searchWin.FindName("btnCancel")

        $btnSearch.Add_Click({
            $searchTerm = $txtSearchTerm.Text.Trim()
            if ($searchTerm) {
                try {
                    $ldapFilter = if ($chkExactMatch.IsChecked) {
                        if ($chkSearchDesc.IsChecked) { "(description=$searchTerm)" } else { "(cn=$searchTerm)" }
                    } else {
                        if ($chkSearchDesc.IsChecked) { "(description=*$searchTerm*)" } else { "(cn=*$searchTerm*)" }
                    }
                    Write-ColoredLog "Starte erweiterte Gruppensuche: $ldapFilter" "Black"
                    $groups = Get-ADGroup -LDAPFilter $ldapFilter -ErrorAction Stop
                    $sortedGroups = $groups | Sort-Object Name

                    Update-UI {
                        $existingGroups = $lstGroups.Items | ForEach-Object { $_.Name }
                        foreach ($grpItem in $sortedGroups) {
                            if ($existingGroups -notcontains $grpItem.Name) {
                                $groupItem = New-Object PSObject -Property @{ Name = $grpItem.Name; IsChecked = $false }
                                $lstGroups.Items.Add($groupItem)
                            }
                        }
                        if ($lstGroups.Items.Count -eq 0) {
                            Write-ColoredLog "Keine Gruppen gefunden für '$searchTerm'" "Orange"
                            [System.Windows.MessageBox]::Show("Keine Gruppen gefunden.")
                        } else {
                            Write-ColoredLog "Erweiterte Suche erfolgreich: $($lstGroups.Items.Count) Gruppen gefunden" "Green"
                        }
                    }
                    $searchWin.Close()
                } catch {
                    Write-ColoredLog "Fehler bei erweiterter Gruppensuche: $($_.Exception.Message)" "Red"
                    [System.Windows.MessageBox]::Show("Fehler bei der Gruppensuche: $($_.Exception.Message)")
                }
            } else {
                Write-ColoredLog "Bitte einen Suchbegriff eingeben!" "Red"
                [System.Windows.MessageBox]::Show("Bitte einen Suchbegriff eingeben!")
            }
        })

        $btnCancel.Add_Click({ $searchWin.Close() })
        $searchWin.ShowDialog() | Out-Null
    })
} else {
    Write-ColoredLog "Fehler: btnSearchGroups_DropDown konnte nicht initialisiert werden." "Red"
}
#region Button-Events
# ------------------------------------------------#
# 10) Button-Events
# ------------------------------------------------#

# Event-Handler für Abbrechen
$btnCancel.Add_Click({
    Update-UI {
        $txtLog.Document.Blocks.Clear()
        Update-Status "Bereit"
        $btnStart.IsEnabled = $true
        Update-ProgressControls -percent 0
    }
    AddToLog "Vorgang abgebrochen" "INFO"
})

# Event-Handler für Beenden
$btnExit.Add_Click({
    Update-UI {
        try {
            $window.Close()
            if ([System.Windows.Application]::Current) {
                [System.Windows.Application]::Current.Shutdown()
            }
        } catch {
            Write-Error "Fehler beim Beenden der Anwendung: $_"
            [System.Environment]::Exit(0)
        }
    }
})

# Event-Handler für Log kopieren
$btnCopyLog.Add_Click({
    $filteredLog = Get-FilteredLog
    [System.Windows.Clipboard]::SetText($filteredLog)
    [System.Windows.MessageBox]::Show("Gefiltertes Log wurde in die Zwischenablage kopiert.", "Info", 
        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
})

# Event-Handler für Neuer Benutzer
$btnNewUser.Add_Click({
    Update-UI {
        # Felder zurücksetzen
        $txtUser.Text = ""
        $txtGivenName.Text = ""
        $txtLastName.Text = ""
        $txtBuro.Text = ""
        $comboSite.SelectedItem = $null
        $txtRufnummer.Text = ""
        $txtHandynummer.Text = ""
        $txtTitle.Text = ""
        $comboAmts.Text = ""
        $comboLauf.Text = ""
        $comboDept.Text = ""
        $txtExp.Text = ""
        $txtAktiv.Text = ""
        $txtTicket.Text = ""
        $comboRolle.SelectedItem = $null
        $comboFunktion.Text = ""
        $comboSonder.Text = ""
        $txtDesc.Text = ""
        $txtRefUser.Text = ""
        $lstGroups.Items.Clear()

        $chkIVBB.IsChecked = $false
        $chkGVPL.IsChecked = $false
        $chkPhonebook.IsChecked = $false
        $chkVIP.IsChecked = $false
        $chkIsFemale.IsChecked = $false
        $chkAbgeordnet.IsChecked = $false
        $chkNatPerson.IsChecked = $false
        $chkConet.IsChecked = $false
        $chkExternAccount.IsChecked = $false
        $chkMailbox.IsChecked = $false
        $chkExtern.IsChecked = $false
        $chkVerstecken.IsChecked = $false
        $chkResMailbox.IsChecked = $false

        $txtLog.Document.Blocks.Clear()
        Update-Status "Bereit für neuen Benutzer"
        Update-ProgressControls -percent 0
    }
    AddToLog "Formular für neuen Benutzer zurückgesetzt" "INFO"
})

#endregion

#region Tastenkombinationen und Fenster-Events
# ------------------------------------------------#
# 11) Tastenkombinationen und Fenster-Events
# ------------------------------------------------#

# Tastenkombinationen hinzufügen
$window.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq "F5" -and $btnStart.IsEnabled) {
        $btnStart.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
    }
    if ($e.Key -eq "Escape") {
        $btnExit.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
    }
})

# Tastenkombinationen (Profil Speichern Strg+S, Profil Laden Strg+L, Strg+S für "Ausführen", Esc für "Abbrechen", Strg+N für "Neuer Benutzer")
$window.Add_KeyDown({
    if ($_.Key -eq "S" -and $_.KeyboardDevice.Modifiers -eq "Ctrl") {
        $btnSaveProfile.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
    }
    elseif ($_.Key -eq "L" -and $_.KeyboardDevice.Modifiers -eq "Ctrl") {
        $btnLoadProfile.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
    }
})


# Event-Handler für Fenster-Schließen
$window.Add_Closing({
    Update-UI {
        $txtLog.Document.Blocks.Clear()
        AddToLog "Anwendung wird beendet" "INFO"
    }
    if ($global:CustomApplication) {
        $global:CustomApplication.Shutdown()
    }
})

# Initiale Standardprofile laden
Update-UI {
    $comboStdProfile.Items.Clear()
    foreach ($key in $global:StdProfiles.PSObject.Properties.Name) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $key
        $comboStdProfile.Items.Add($item)
    }
    Update-Status "Bereit"
    $window.Title = "Add-User"
}

#endregion

# Fenster anzeigen
$window.ShowDialog() | Out-Null
if ($global:CustomApplication) {
    $global:CustomApplication.Run()
}
