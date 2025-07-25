Set-StrictMode -Version Latest
. "$PSScriptRoot\TestHelpers.ps1"

Describe "ProcessUserCreation Businesslogic" {
    BeforeAll {
        Import-AddUserRoot
        Use-TestConfig
    }

    Context "Neuer Benutzer (Happy Path)" {
        BeforeEach {
            Mock Get-ADUser { $null } -Verifiable
            Mock new-ADAccountSettings { @{ dummy = $true } } -Verifiable
            Mock Set-NewADAccount { $true } -Verifiable
            Mock WriteJobLog { } -Verifiable
            Mock append-database { $true } -Verifiable
            Mock write-XMLLog { $true } -Verifiable
        }

        It "legt neuen Benutzer an und loggt SUCCESS" {
            $null = ProcessUserCreation `
                -UserID 'u123' -givenName 'Max' -lastName 'Muster' `
                -gender 'Mann' -Buro '' -Rufnummer '' -Handynummer '' `
                -titleValue '' -amtsbez '' -laufgruppe '' -roleSelection '' `
                -Site 'STR' -ExpDate '' -desc '' -Department 'Z II 5' `
                -TicketNr 'T-1' -EntryDate '' -sonderkenn '' -funktion '' `
                -refUser '' -isIVBB 'n' -isGVPL 'n' -isVIP 'n' -isFemale '' `
                -isExtern:$false -isVerstecken:$false -isPhonebook:$false `
                -isNatPerson '' -isResMailbox '' -isAbgeordnet '' -isConet '' `
                -isExternAccount '' -makeMailbox '' `
                -departmentOGMapping @{} -departmentMapping @{} `
                -AdditionalGroups @() -LogTextbox $null -InJob `
                -ProgressCallback { }

            Assert-VerifiableMocks
            # Prüfe ob SUCCESS geloggt wurde (wir könnten WriteJobLog aufzeichnen)
            Should -Invoke WriteJobLog -ParameterFilter { $msg -like '*erfolgreich erstellt*' -and $category -eq 'SUCCESS' } -Times 1
        }
    }

    Context "Benutzer existiert bereits" {
        BeforeEach {
            Mock Get-ADUser { @{ SamAccountName = 'u123' } } -Verifiable
            Mock WriteJobLog { } -Verifiable
        }

        It "loggt WARN und bricht im Job-Mode ab" {
            ProcessUserCreation -UserID 'u123' -givenName 'a' -lastName 'b' `
                -gender '' -Buro '' -Rufnummer '' -Handynummer '' `
                -titleValue '' -amtsbez '' -laufgruppe '' -roleSelection '' `
                -Site 'STR' -ExpDate '' -desc '' -Department '' `
                -TicketNr '' -EntryDate '' -sonderkenn '' -funktion '' `
                -refUser '' -isIVBB '' -isGVPL '' -isVIP '' -isFemale '' `
                -isExtern:$false -isVerstecken:$false -isPhonebook:$false `
                -isNatPerson '' -isResMailbox '' -isAbgeordnet '' -isConet '' `
                -isExternAccount '' -makeMailbox '' `
                -departmentOGMapping @{} -departmentMapping @{} `
                -AdditionalGroups @() -LogTextbox $null -InJob `
                -ProgressCallback { } | Out-Null

            Assert-VerifiableMocks
            Should -Invoke WriteJobLog -ParameterFilter { $msg -like '*existiert bereits*' -and $category -eq 'WARN' } -Times 1
        }
    }

    Context "Fehler beim Set-NewADAccount" {
        BeforeEach {
            Mock Get-ADUser { $null }
            Mock new-ADAccountSettings { @{ dummy = $true } }
            Mock Set-NewADAccount { throw "BOOM" }
            Mock WriteJobLog { } -Verifiable
        }

        It "wirft Exception und loggt ERROR" {
            { 
                ProcessUserCreation -UserID 'u999' -givenName 'a' -lastName 'b' `
                    -gender '' -Buro '' -Rufnummer '' -Handynummer '' `
                    -titleValue '' -amtsbez '' -laufgruppe '' -roleSelection '' `
                    -Site 'STR' -ExpDate '' -desc '' -Department '' `
                    -TicketNr '' -EntryDate '' -sonderkenn '' -funktion '' `
                    -refUser '' -isIVBB '' -isGVPL '' -isVIP '' -isFemale '' `
                    -isExtern:$false -isVerstecken:$false -isPhonebook:$false `
                    -isNatPerson '' -isResMailbox '' -isAbgeordnet '' -isConet '' `
                    -isExternAccount '' -makeMailbox '' `
                    -departmentOGMapping @{} -departmentMapping @{} `
                    -AdditionalGroups @() -LogTextbox $null -InJob `
                    -ProgressCallback { } | Out-Null
            } | Should -Throw

            Should -Invoke WriteJobLog -ParameterFilter { $category -eq 'ERROR' } -Times 1
        }
    }
}

