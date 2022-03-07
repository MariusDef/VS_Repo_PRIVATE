Import-Module ActiveDirectory

$Members = Get-ADGroupMember -Identity "DA-Clients"

 ForEach ($Member in $Members) {
    Write-Host $Member.SamAccountName
    Add-ADGroupMember -Identity "SW_DirectAccess" -Members $Member.SamAccountName
 }