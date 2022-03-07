$AllUsers = Get-ADUser -Filter {(enabled -eq $true)} -Properties displayName, manager, extensionAttribute2 -SearchBase "OU=Internal User,OU=Usermanagement,DC=cellent,DC=int"#| Select-String displayName, manager, extensionAttribute2


$AllUsersManaged = $AllUsers #| 
        #Select-Object @{Label = "DisplayName";Expression = {$_.displayName}}, 
        #@{Label = "Manager";Expression = {%{(Get-AdUser $_.Manager -Properties Mail).Mail}}},
        #@{Label = "extensionAttribute2";Expression = {$_.extensionAttribute2}} #| Export-Csv -Path E:\temp\AllUsers-withManagerandExtensionAttribute.csv -NoTypeInformation
        


"Username; Manager L1 MailAdress; Heat-Approver L1; L1 Name; L2 MailAdress; Heat-Approver L2" |Out-File -FilePath E:\Temp\AllUsers-withManager-plus-SubManager.txt

foreach($Item in $AllUsersManaged)
{
    try
    {
    $AllUsersSub = Get-ADUser -Filter {distinguishedName -eq $Item.Manager} -Properties displayName, manager, extensionAttribute2 -SearchBase "OU=Internal User,OU=Usermanagement,DC=cellent,DC=int"
    "$($Item.DisplayName); $($(Get-AdUser $($item.Manager) -Properties Mail).Mail); $($item.extensionAttribute2); $($AllUsersSub.displayName); $($(Get-AdUser $($AllUsersSub.Manager) -Properties Mail).Mail); $($AllUsersSub.extensionAttribute2)" |Out-File -FilePath E:\Temp\AllUsers-withManager-plus-SubManager.txt -Append
    }
    catch
    {}
#    if ($Item.Manager -ne $Item.extensionAttribute2)
#    {
#        Write-Host $Item.DisplayName';'$Item.Manager';'$Item.extensionAttribute2
#    }
}

#| Export-Csv -Path E:\temp\export.csv -NoTypeInformation

