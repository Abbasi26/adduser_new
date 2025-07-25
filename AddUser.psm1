$here = Split-Path -Parent $PSCommandPath
$pub  = Join-Path $here 'Modules/Public'
$int  = Join-Path $here 'Modules/Internal'

if (Test-Path $int) {
    Get-ChildItem $int -Filter *.psm1 -File | ForEach-Object {
        Import-Module $_.FullName -Force
    }
}
if (Test-Path $pub) {
    Get-ChildItem $pub -Filter *.psm1 -File | ForEach-Object {
        Import-Module $_.FullName -Force
    }
}

$publicFunctions = Get-Module |
    Where-Object { $_.Path -like (Join-Path $pub '*') } |
    ForEach-Object { $_.ExportedFunctions.Keys } |
    Sort-Object -Unique

if ($publicFunctions) {
    Export-ModuleMember -Function $publicFunctions
} else {
    Export-ModuleMember -Function *
}
