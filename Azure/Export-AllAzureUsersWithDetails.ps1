Import-Module AzureAD
#connect-azuread

$ExportCSV = "C:\Temp\AzureUsers_active.csv"

$Measure_getUsers = Measure-Command {
    $Users = get-azureaduser -all $true -Filter "UserType eq 'Member' and AccountEnabled eq true and startswith(userprincipalname,'Marius.Deffner')" | `
        Select-Object -Property CompanyName,Country,Department,GivenName,Surname,JobTitle,Mail, `
            Mobile,TelephoneNumber,POstalCode,StreetAddress,City,UserPrincipalName,Manager
    $users | Out-Null
}

"Count Rows in Users-Object: " + $users.Count
"Get Users (total minutes): " + $Measure_getUsers.TotalMinutes

$Measure_getManagers = Measure-Command {
    foreach ($user in $Users)
    {
        $User.Manager = (Get-AzureADUserManager -ObjectId $user.UserPrincipalName).UserPrincipalName
        
    }
}

$users | Export-Csv $ExportCSV -Force

"Get Managers (total minutes): " + $Measure_getManagers.TotalMinutes