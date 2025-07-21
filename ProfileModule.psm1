function Save-UserProfile {
    param(
        [hashtable]$UI,
        [string]   $BasePath
    )
    if (-not (Test-Path $BasePath)) { New-Item -ItemType Directory -Path $BasePath -Force | Out-Null }
    $fixedGroups = @()
    if ($UI.txtUser.Text.Trim()) {
        try {
            $adUser = Get-ADUser -Identity $UI.txtUser.Text.Trim() -Properties MemberOf -ErrorAction Stop
            $fixedGroups = $adUser.MemberOf | ForEach-Object {
                ($_ -match '^CN=([^,]+)') ? $Matches[1] : $_
            }
        } catch { }
    }

    $profile = [ordered]@{
        UserID         = $UI.txtUser.Text.Trim()
        Gender         = ($UI.comboGender.SelectedItem)?.Content ?? 'Mann'
        NatPerson      = if ($UI.chkNatPerson.IsChecked) { 'j' } else { 'n' }
        GivenName      = $UI.txtGivenName.Text.Trim()
        LastName       = $UI.txtLastName.Text.Trim()
        Buro           = $UI.txtBuro.Text.Trim()
        Site           = ($UI.comboSite.SelectedItem)?.Content ?? ''
        Department     = $UI.comboDept.Text.Trim()
        Rufnummer      = $UI.txtRufnummer.Text.Trim()
        Handynummer    = $UI.txtHandynummer.Text.Trim()
        Title          = $UI.txtTitle.Text.Trim()
        Amtsbez        = $UI.comboAmts.Text.Trim()
        Laufgruppe     = $UI.comboLauf.Text.Trim()
        ExpDate        = $UI.txtExp.Text.Trim()
        EntryDate      = $UI.txtAktiv.Text.Trim()
        TicketNr       = $UI.txtTicket.Text.Trim()
        Rolle          = ($UI.comboRolle.SelectedItem)?.Content ?? ''
        Sonderkenn     = $UI.comboSonder.Text.Trim()
        Funktion       = $UI.comboFunktion.Text.Trim()
        Description    = $UI.txtDesc.Text.Trim()
        RefUser        = $UI.txtRefUser.Text.Trim()
        AdditionalGroups      = @($UI.lstGroups.Items | Where-Object { $_.IsChecked } | ForEach-Object { $_.Name })
        FixedAdditionalGroups = $fixedGroups
        IVBB           = if ($UI.chkIVBB.IsChecked)  { 'j' } else { 'n' }
        GVPL           = if ($UI.chkGVPL.IsChecked)  { 'j' } else { 'n' }
        VIP            = if ($UI.chkVIP.IsChecked)   { 'j' } else { 'n' }
        Extern         = $UI.chkExtern.IsChecked
        Verstecken     = $UI.chkVerstecken.IsChecked
        Phonebook      = $UI.chkPhonebook.IsChecked
        Abgeordnet     = if ($UI.chkAbgeordnet.IsChecked) { 'j' } else { 'n' }
        Conet          = if ($UI.chkConet.IsChecked) { 'j' } else { 'n' }
        ExternAccount  = if ($UI.chkExternAccount.IsChecked){ 'j' } else { 'n' }
        ResMailbox     = if ($UI.chkResMailbox.IsChecked){ 'j' } else { 'n' }
        Mailbox        = if ($UI.chkMailbox.IsChecked) { 'j' } else { 'n' }
    }

    $file = Join-Path $BasePath ("{0}_{1}.json" -f $profile.UserID,$profile.TicketNr)
    $dlg  = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.InitialDirectory = $BasePath
    $dlg.Filter = 'JSON (*.json)|*.json'
    $dlg.FileName = [IO.Path]::GetFileName($file)

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $profile | ConvertTo-Json -Depth 5 | Out-File -FilePath $dlg.FileName -Encoding UTF8
        [Windows.MessageBox]::Show("Profil gespeichert: $($dlg.FileName)")
    }
}

function Load-UserProfile {
    param(
        [hashtable]$UI,
        [string]   $BasePath
    )
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.InitialDirectory = $BasePath
    $dlg.Filter = 'JSON (*.json)|*.json'

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $json = Get-Content $dlg.FileName -Raw -Encoding UTF8 | ConvertFrom-Json

    if (-not (Show-ProfilePreview -ProfileData $json)) { return }

    $UI.txtUser.Text      = $json.UserID
    $UI.comboGender.SelectedItem = $UI.comboGender.Items | Where-Object { $_.Content -eq $json.Gender }
    $UI.chkNatPerson.IsChecked   = ($json.NatPerson -eq 'j')
    $UI.txtGivenName.Text = $json.GivenName
    $UI.txtLastName.Text  = $json.LastName
    $UI.txtBuro.Text      = $json.Buro
    $UI.comboSite.SelectedItem = $UI.comboSite.Items | Where-Object { $_.Content -eq $json.Site }
    $UI.comboDept.Text    = $json.Department
    $UI.txtRufnummer.Text = $json.Rufnummer
    $UI.txtHandynummer.Text= $json.Handynummer
    $UI.txtTitle.Text     = $json.Title
    $UI.comboAmts.Text    = $json.Amtsbez
    $UI.comboLauf.Text    = $json.Laufgruppe
    $UI.txtExp.Text       = $json.ExpDate
    $UI.txtAktiv.Text     = $json.EntryDate
    $UI.txtTicket.Text    = $json.TicketNr
    $UI.comboRolle.SelectedItem = $UI.comboRolle.Items | Where-Object { $_.Content -eq $json.Rolle }
    $UI.comboSonder.Text  = $json.Sonderkenn
    $UI.comboFunktion.Text= $json.Funktion
    $UI.txtDesc.Text      = $json.Description
    $UI.txtRefUser.Text   = $json.RefUser

    $UI.lstGroups.Items.Clear()
    $all = @($json.FixedAdditionalGroups) + @($json.AdditionalGroups)
    foreach ($grp in ($all | Sort-Object -Unique)) {
        $item = New-Object GroupItem
        $item.Name      = $grp
        $item.IsChecked = ($json.AdditionalGroups -contains $grp)
        $UI.lstGroups.Items.Add($item)
    }

    $UI.chkIVBB.IsChecked       = ($json.IVBB -eq 'j')
    $UI.chkGVPL.IsChecked       = ($json.GVPL -eq 'j')
    $UI.chkVIP.IsChecked        = ($json.VIP  -eq 'j')
    $UI.chkExtern.IsChecked     = $json.Extern
    $UI.chkVerstecken.IsChecked = $json.Verstecken
    $UI.chkPhonebook.IsChecked  = $json.Phonebook
    $UI.chkAbgeordnet.IsChecked = ($json.Abgeordnet -eq 'j')
    $UI.chkConet.IsChecked      = ($json.Conet -eq 'j')
    $UI.chkExternAccount.IsChecked = ($json.ExternAccount -eq 'j')
    $UI.chkResMailbox.IsChecked = ($json.ResMailbox -eq 'j')
    $UI.chkMailbox.IsChecked    = ($json.Mailbox -eq 'j')

    [Windows.MessageBox]::Show("Profil erfolgreich übernommen!")
}

Export-ModuleMember -Function Save-UserProfile, Load-UserProfile
